import re
import pandas as pd
from io import BytesIO


ERP_COLS = {
    "품목코드": "품목코드",
    "품목명": "품목명",
    "규격": "규격",
    "단위": "단위",
    "이동수량": "이동수량",
    "비고": "비고",
}

ALPS_COLS = {
    "ERP코드": "ERP코드",
    "제품명": "제품명",
    "출고수량": "출고수량",
    "단위": "단위",
    "상태": "상태",
    "규격": "규격",
}


def _extract_date_from_note(note: str) -> str | None:
    """비고 컬럼에서 날짜 추출 (예: '이동처리(2026-05-12)' → '2026-05-12')."""
    if not isinstance(note, str):
        return None
    m = re.search(r"\d{4}-\d{2}-\d{2}", note)
    return m.group(0) if m else None


def load_erp(file, sheet_name=None) -> pd.DataFrame:
    """ERP 엑셀 파일을 읽어 DataFrame으로 반환."""
    kwargs = {"dtype": str}
    if sheet_name:
        kwargs["sheet_name"] = sheet_name

    try:
        df = pd.read_excel(file, **kwargs)
    except Exception:
        # sheet_name 지정 실패 시 첫 번째 시트로 재시도
        df = pd.read_excel(file, dtype=str)

    df.columns = df.columns.str.strip()

    missing = [c for c in ERP_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"ERP 파일에 필수 컬럼 없음: {missing}\n실제 컬럼: {list(df.columns)}")

    df["이동수량"] = pd.to_numeric(df["이동수량"], errors="coerce").fillna(0)
    df["_날짜"] = df["비고"].apply(_extract_date_from_note)

    # 날짜 없는 행: 이동일자 컬럼이 있으면 대체
    if "이동일자" in df.columns and df["_날짜"].isna().any():
        mask = df["_날짜"].isna()
        df.loc[mask, "_날짜"] = df.loc[mask, "이동일자"].astype(str).apply(
            lambda x: re.search(r"\d{4}[-/]\d{2}[-/]\d{2}", x).group(0).replace("/", "-")
            if re.search(r"\d{4}[-/]\d{2}[-/]\d{2}", x) else None
        )

    return df


def load_alps(file, sheet_name=None) -> pd.DataFrame:
    """Alps 엑셀 파일을 읽어 DataFrame으로 반환."""
    kwargs = {"dtype": str}
    if sheet_name:
        kwargs["sheet_name"] = sheet_name

    try:
        df = pd.read_excel(file, **kwargs)
    except Exception:
        df = pd.read_excel(file, dtype=str)

    df.columns = df.columns.str.strip()

    # (급)품목코드 → ERP코드 대체
    if "ERP코드" not in df.columns and "(급)품목코드" in df.columns:
        df = df.rename(columns={"(급)품목코드": "ERP코드"})

    missing = [c for c in ALPS_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Alps 파일에 필수 컬럼 없음: {missing}\n실제 컬럼: {list(df.columns)}")

    df["출고수량"] = pd.to_numeric(df["출고수량"], errors="coerce").fillna(0)

    # PAK 단위 → EA 자동 환산
    # 제품명에 <PAK 단위> 또는 <N봉 단위> 명시된 경우만 적용
    def _pak_ratio(row):
        name = str(row.get("제품명", ""))
        # <PAK 단위> 명시된 경우: 제품명에서 *MEA 추출
        if "<PAK 단위>" in name or "<pak 단위>" in name.lower():
            m = re.search(r"\*\s*(\d+)\s*EA", name, re.IGNORECASE)
            return int(m.group(1)) if m else 1
        # <N봉 단위> 명시된 경우: N 추출
        m2 = re.search(r"<\s*(\d+)\s*봉\s*단위>", name)
        if m2:
            return int(m2.group(1))
        return 1

    df["_pak_ratio"] = df.apply(_pak_ratio, axis=1)
    df["출고수량_EA"] = df["출고수량"] * df["_pak_ratio"]

    return df


def assign_d_levels(erp_df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """비고 날짜 기준으로 D1/D2/D3 분류."""
    dates = sorted(erp_df["_날짜"].dropna().unique())
    d_map = {}
    for i, d in enumerate(dates[:3]):
        d_map[d] = f"D{i+1}"

    erp_df = erp_df.copy()
    erp_df["_D레벨"] = erp_df["_날짜"].map(d_map)
    return erp_df, d_map


def run_check(erp_df: pd.DataFrame, alps_df: pd.DataFrame, erp_filter: str = "TCT 케이터링 이동처리") -> dict:
    """ERP와 Alps 수량 비교 실행."""
    # ERP 비고 필터링
    if erp_filter:
        mask = erp_df["비고"].str.contains(erp_filter, na=False)
        erp_df = erp_df[mask].copy()

    erp_df, d_map = assign_d_levels(erp_df)

    # ERP: 품목코드별 D레벨 합산
    erp_grouped = (
        erp_df.groupby(["품목코드", "품목명", "규격", "단위", "_D레벨"])["이동수량"]
        .sum()
        .unstack(fill_value=0)
        .reset_index()
    )

    for d in ["D1", "D2", "D3"]:
        if d not in erp_grouped.columns:
            erp_grouped[d] = 0

    erp_grouped["ERP합산"] = erp_grouped[["D1", "D2", "D3"]].sum(axis=1)

    # Alps: 취소·변경 상태 제외 후 합산
    # Alps출고수량 = 원본 수량 합계 (화면 표시용)
    # Alps환산수량 = PAK/봉 환산 적용 수량 (비교 계산용)
    alps_valid = alps_df[~alps_df["상태"].isin(["취소", "변경"])].copy()
    alps_grouped = (
        alps_valid.groupby("ERP코드")
        .agg(
            Alps출고수량=("출고수량", "sum"),       # 원본 수량 (표시용)
            Alps환산수량=("출고수량_EA", "sum"),     # 환산 수량 (비교용)
            Alps단위=("단위", lambda x: x.mode().iloc[0] if not x.empty else ""),
            Alps제품명=("제품명", lambda x: x.mode().iloc[0] if not x.empty else ""),
        )
        .reset_index()
    )

    # 병합
    merged = erp_grouped.merge(
        alps_grouped,
        left_on="품목코드",
        right_on="ERP코드",
        how="outer",
    )

    merged["ERP합산"] = merged["ERP합산"].fillna(0)
    merged["Alps출고수량"] = merged["Alps출고수량"].fillna(0)
    merged["Alps환산수량"] = merged["Alps환산수량"].fillna(0)

    # 단위 혼재 감지
    def _unit_mixed(row):
        erp_u = str(row.get("단위", "")).strip()
        alps_u = str(row.get("Alps단위", "")).strip()
        return erp_u != alps_u and erp_u and alps_u

    merged["단위혼재"] = merged.apply(_unit_mixed, axis=1).astype(bool)

    # 환산비 없이 1차 비교 (Alps환산수량 기준으로 차이 계산)
    merged["환산비"] = 1.0
    merged["ERP환산수량"] = merged["ERP합산"].astype(float)
    merged["차이"] = merged["ERP합산"] - merged["Alps환산수량"]
    merged["일치여부"] = (merged["차이"].abs() < 0.001).astype(bool)

    # 단위혼재 + 불일치 항목에 대해 제안 환산비 계산
    def _suggest_ratio(row):
        """ERP합산/Alps출고수량 가 정수에 가까우면 그 값, 아니면 규격에서 추출"""
        if row["Alps출고수량"] <= 0:
            return 1.0
        ratio_calc = row["ERP합산"] / row["Alps출고수량"]
        if abs(ratio_calc - round(ratio_calc)) < 0.05:
            return round(ratio_calc)
        # 규격에서 추출
        nums = re.findall(r"\*\s*(\d+(?:\.\d+)?)\s*ea", str(row.get("규격", "")), re.IGNORECASE)
        return float(nums[-1]) if nums else 1.0

    mask_mixed_miss = merged["단위혼재"] & ~merged["일치여부"]
    merged.loc[mask_mixed_miss, "환산비"] = merged[mask_mixed_miss].apply(_suggest_ratio, axis=1)

    # 품목명 보완
    merged["품목명"] = merged["품목명"].fillna(merged["Alps제품명"])
    merged["규격"] = merged["규격"].fillna("")

    # 요약
    total_erp = merged["ERP합산"].sum()
    total_alps = merged["Alps출고수량"].sum()
    match_count = merged["일치여부"].sum()
    mismatch_count = (~merged["일치여부"]).sum()
    mixed_unit_count = merged["단위혼재"].sum()

    match_qty = merged.loc[merged["일치여부"], "ERP합산"].sum()
    mismatch_qty = merged.loc[~merged["일치여부"], "ERP합산"].sum()

    mismatch_df = merged[~merged["일치여부"]].copy().sort_values("차이", key=abs, ascending=False)

    # 환산기준 시트용: 자동 환산 규칙 추출
    ratio_rules = alps_valid[alps_valid["_pak_ratio"] > 1][
        ["ERP코드", "제품명", "규격", "단위", "_pak_ratio"]
    ].drop_duplicates(["ERP코드", "_pak_ratio"])

    return {
        "merged": merged,
        "mismatch": mismatch_df,
        "d_map": d_map,
        "_ratio_rules": ratio_rules,
        "custom_ratios": {},
        "summary": {
            "total_erp": total_erp,
            "total_alps": total_alps,
            "match_count": int(match_count),
            "mismatch_count": int(mismatch_count),
            "mixed_unit_count": int(mixed_unit_count),
            "match_qty": match_qty,
            "mismatch_qty": mismatch_qty,
            "match_pct": match_qty / total_erp * 100 if total_erp else 0,
        },
    }


def apply_custom_ratios(result: dict, custom_ratios: dict) -> dict:
    """사용자 지정 환산비를 적용해 결과 재계산. custom_ratios = {품목코드: ratio}"""
    merged = result["merged"].copy()

    for code, ratio in custom_ratios.items():
        mask = merged["품목코드"] == code
        r = float(ratio) if ratio else 1.0
        merged.loc[mask, "환산비"] = r
        merged.loc[mask, "ERP환산수량"] = merged.loc[mask, "ERP합산"] / r if r > 0 else merged.loc[mask, "ERP합산"]

    merged["차이"] = merged["ERP환산수량"] - merged["Alps환산수량"]
    merged["일치여부"] = (merged["차이"].abs() < 0.001).astype(bool)

    total_erp = merged["ERP합산"].sum()
    match_qty = merged.loc[merged["일치여부"], "ERP합산"].sum()
    mismatch_qty = merged.loc[~merged["일치여부"], "ERP합산"].sum()

    return {
        **result,
        "merged": merged,
        "mismatch": merged[~merged["일치여부"]].copy().sort_values("차이", key=abs, ascending=False),
        "custom_ratios": custom_ratios,
        "summary": {
            **result["summary"],
            "match_count": int(merged["일치여부"].sum()),
            "mismatch_count": int((~merged["일치여부"]).sum()),
            "match_qty": match_qty,
            "mismatch_qty": mismatch_qty,
            "match_pct": match_qty / total_erp * 100 if total_erp else 0,
        },
    }


def to_excel_bytes(result: dict) -> bytes:
    """결과를 Excel 파일로 변환."""
    buf = BytesIO()
    d_map = result["d_map"]
    date_labels = {v: k for k, v in d_map.items()}

    with pd.ExcelWriter(buf, engine="xlsxwriter", engine_kwargs={"options": {"nan_inf_to_errors": True}}) as writer:
        wb = writer.book

        # ── 스타일 ──────────────────────────────────────────────
        hdr_fmt = wb.add_format({"bold": True, "bg_color": "#1F4E79", "font_color": "white", "border": 1, "align": "center"})
        match_fmt = wb.add_format({"bg_color": "#E2EFDA", "border": 1})
        mismatch_fmt = wb.add_format({"bg_color": "#FCE4D6", "border": 1})
        mixed_fmt = wb.add_format({"bg_color": "#FFF2CC", "border": 1})
        num_fmt = wb.add_format({"num_format": "#,##0", "border": 1})
        num_match_fmt = wb.add_format({"num_format": "#,##0", "border": 1, "bg_color": "#E2EFDA"})
        num_mismatch_fmt = wb.add_format({"num_format": "#,##0", "border": 1, "bg_color": "#FCE4D6"})

        # ── 검증결과 시트 ────────────────────────────────────────
        merged = result["merged"].copy()
        s = result["summary"]

        d1_label = date_labels.get("D1", "D1")
        d2_label = date_labels.get("D2", "D2")
        d3_label = date_labels.get("D3", "D3")

        cols = [
            "품목코드", "품목명", "규격", "단위",
            f"D1\n({d1_label})", f"D2\n({d2_label})", f"D3\n({d3_label})",
            "합산\n(D1+D2+D3)", "Alps출고수량", "차이\n(ERP-Alps)", "일치여부", "단위혼재",
        ]
        export_df = merged.rename(columns={
            "D1": f"D1\n({d1_label})",
            "D2": f"D2\n({d2_label})",
            "D3": f"D3\n({d3_label})",
            "ERP합산": "합산\n(D1+D2+D3)",
            "차이": "차이\n(ERP-Alps)",
        })[cols]

        export_df["일치여부"] = export_df["일치여부"].map({True: "✓ 일치", False: "✗ 불일치"})
        export_df["단위혼재"] = export_df["단위혼재"].map({True: "⚠ 혼재", False: ""})

        ws = writer.book.add_worksheet("검증결과")
        writer.sheets["검증결과"] = ws

        for c, col in enumerate(cols):
            ws.write(0, c, col, hdr_fmt)
            ws.set_column(c, c, 14)

        ws.set_column(1, 1, 30)  # 품목명
        ws.set_column(2, 2, 20)  # 규격

        for r, row in export_df.iterrows():
            is_match = row["일치여부"] == "✓ 일치"
            is_mixed = row["단위혼재"] == "⚠ 혼재"
            row_fmt = match_fmt if is_match else (mixed_fmt if is_mixed else mismatch_fmt)
            num_row_fmt = num_match_fmt if is_match else num_mismatch_fmt
            for c, col in enumerate(cols):
                val = row[col]
                if c >= 4 and c <= 9:
                    ws.write(r + 1, c, val, num_row_fmt)
                else:
                    ws.write(r + 1, c, val, row_fmt)

        # ── 불일치목록 시트 ──────────────────────────────────────
        mismatch_export = result["mismatch"].rename(columns={
            "D1": f"D1({d1_label})",
            "D2": f"D2({d2_label})",
            "D3": f"D3({d3_label})",
            "ERP합산": "합산(D1+D2+D3)",
            "차이": "차이(ERP-Alps)",
        })
        mismatch_export["단위혼재"] = mismatch_export["단위혼재"].map({True: "⚠ 혼재", False: ""})
        mismatch_export.to_excel(writer, sheet_name="불일치목록", index=False)

        # ── 환산기준 시트 ─────────────────────────────────────────
        # 자동 환산 규칙: Alps _pak_ratio > 1 인 개별 행 기준
        ratio_rules = result.get("_ratio_rules", pd.DataFrame())
        custom_ratios = result.get("custom_ratios", {})
        merged = result["merged"]

        rows = []
        # 자동 환산 항목
        if not ratio_rules.empty:
            for _, r in ratio_rules.iterrows():
                rows.append({
                    "품목코드": r["ERP코드"],
                    "제품명(Alps)": r["제품명"],
                    "규격": r.get("규격", ""),
                    "ERP단위": "",
                    "Alps단위": r["단위"],
                    "환산비(1 Alps단위 = N EA)": int(r["_pak_ratio"]),
                    "구분": "자동(제품명 기준)",
                })

        # 수동 환산 항목
        for code, ratio in custom_ratios.items():
            m_row = merged[merged["품목코드"] == code]
            if not m_row.empty:
                mr = m_row.iloc[0]
                rows.append({
                    "품목코드": code,
                    "제품명(Alps)": mr["품목명"],
                    "규격": mr["규격"],
                    "ERP단위": mr["단위"],
                    "Alps단위": mr["Alps단위"],
                    "환산비(1 Alps단위 = N EA)": ratio,
                    "구분": "수동 설정",
                })

        ratio_sheet = pd.DataFrame(rows).drop_duplicates(["품목코드", "Alps단위"]) if rows else \
            pd.DataFrame(columns=["품목코드", "제품명(Alps)", "규격", "ERP단위", "Alps단위",
                                   "환산비(1 Alps단위 = N EA)", "구분"])

        ratio_sheet.to_excel(writer, sheet_name="환산기준", index=False)
        ws_r = writer.sheets["환산기준"]
        ws_r.set_column(0, 0, 14)
        ws_r.set_column(1, 1, 40)
        ws_r.set_column(2, 2, 20)
        ws_r.set_column(6, 6, 12)

    return buf.getvalue()
