import streamlit as st
import pandas as pd
from io import BytesIO
from logic import load_erp, load_alps, run_check, to_excel_bytes, apply_custom_ratios

st.set_page_config(
    page_title="ERP ↔ Alps 수량 검증",
    page_icon="📦",
    layout="wide",
)

st.title("📦 ERP ↔ Alps 출고수량 검증")
st.caption("ERP(D1+D2+D3) 이동수량과 Alps 출고수량을 비교합니다.")

# ── 파일 업로드 ──────────────────────────────────────────────────────────────
st.markdown("---")
col1, col2 = st.columns(2)

with col1:
    st.subheader("① ERP 파일")
    erp_file = st.file_uploader(
        "ERP 엑셀 파일을 업로드하세요",
        type=["xlsx", "xlsm", "xls"],
        key="erp",
    )
    erp_sheet = st.text_input("시트명 (비워두면 첫 번째 시트)", value="ERP_D1", key="erp_sheet")

with col2:
    st.subheader("② Alps 파일")
    alps_file = st.file_uploader(
        "Alps 엑셀 파일을 업로드하세요",
        type=["xlsx", "xlsm", "xls"],
        key="alps",
    )
    alps_sheet = st.text_input("시트명 (비워두면 첫 번째 시트)", value="Alps_원본", key="alps_sheet")

st.markdown("---")

# ── 단일 파일 모드 ────────────────────────────────────────────────────────────
with st.expander("하나의 파일에 두 시트가 모두 있는 경우 (xlsm)"):
    combined_file = st.file_uploader(
        "ERP_D1 + Alps_원본 시트가 포함된 파일 업로드",
        type=["xlsx", "xlsm"],
        key="combined",
    )
    c1, c2 = st.columns(2)
    with c1:
        erp_sheet_c = st.text_input("ERP 시트명", value="ERP_D1", key="erp_sheet_c")
    with c2:
        alps_sheet_c = st.text_input("Alps 시트명", value="Alps_원본", key="alps_sheet_c")

# ── ERP 비고 필터 ─────────────────────────────────────────────────────────────
erp_filter = st.text_input(
    "ERP 비고 필터 (해당 텍스트 포함된 행만 집계)",
    value="TCT 케이터링 이동처리",
    help="비워두면 전체 행 집계. 기본값: 'TCT 케이터링 이동처리'",
)

# ── 검증 실행 ─────────────────────────────────────────────────────────────────
run_btn = st.button("🔍 검증 실행", type="primary", use_container_width=True)

if run_btn:
    erp_data, alps_data = None, None

    # 단일 파일 우선
    if combined_file:
        try:
            combined_bytes = BytesIO(combined_file.read())
            erp_data = load_erp(BytesIO(combined_bytes.getvalue()), sheet_name=erp_sheet_c or None)
            alps_data = load_alps(BytesIO(combined_bytes.getvalue()), sheet_name=alps_sheet_c or None)
        except Exception as e:
            st.error(f"파일 읽기 실패: {e}")
    elif erp_file and alps_file:
        try:
            erp_data = load_erp(erp_file, sheet_name=erp_sheet or None)
            alps_data = load_alps(alps_file, sheet_name=alps_sheet or None)
        except Exception as e:
            st.error(f"파일 읽기 실패: {e}")
    else:
        st.warning("파일을 업로드해 주세요.")

    if erp_data is not None and alps_data is not None:
        with st.spinner("검증 중..."):
            try:
                result = run_check(erp_data, alps_data, erp_filter=erp_filter.strip())
            except Exception as e:
                st.error(f"검증 실패: {e}")
                st.stop()

        s = result["summary"]
        d_map = result["d_map"]
        date_labels = {v: k for k, v in d_map.items()}

        # ── 헤더 요약 ──────────────────────────────────────────────────────────
        st.success("✅ 검증 완료")

        date_str = " / ".join(
            f"{d_map[k]} → {k}"
            for k in sorted(d_map.keys())
        )
        st.caption(f"날짜 분류: {date_str}")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("ERP 이동수량 합계", f"{s['total_erp']:,.0f}")
        m2.metric("Alps 출고수량 합계", f"{s['total_alps']:,.0f}")
        m3.metric(
            "✓ 일치",
            f"{s['match_count']}건",
            f"{s['match_pct']:.1f}%",
        )
        m4.metric(
            "✗ 불일치",
            f"{s['mismatch_count']}건",
            f"{100 - s['match_pct']:.1f}%",
            delta_color="inverse",
        )

        # ── 환산비 조정 ────────────────────────────────────────────────────────
        need_ratio = result["merged"][result["merged"]["단위혼재"] & ~result["merged"]["일치여부"]].copy()
        if not need_ratio.empty:
            with st.expander(f"⚙ 단위 환산비 조정 ({len(need_ratio)}건) — 확인 후 '환산 적용' 클릭"):
                st.caption("ERP 수량 ÷ 환산비 = Alps 비교 수량 | 자동 계산된 값이며 직접 수정 가능합니다.")
                ratio_df = need_ratio[["품목코드", "품목명", "규격", "단위", "Alps단위", "ERP합산", "Alps출고수량", "환산비"]].copy()
                ratio_df["ERP÷환산비"] = (ratio_df["ERP합산"] / ratio_df["환산비"]).round(2)

                edited = st.data_editor(
                    ratio_df,
                    column_config={
                        "품목코드":    st.column_config.TextColumn(disabled=True),
                        "품목명":      st.column_config.TextColumn(disabled=True),
                        "규격":        st.column_config.TextColumn(disabled=True),
                        "단위":        st.column_config.TextColumn(disabled=True),
                        "Alps단위":    st.column_config.TextColumn(disabled=True),
                        "ERP합산":     st.column_config.NumberColumn(disabled=True, format="%,.0f"),
                        "Alps출고수량": st.column_config.NumberColumn(disabled=True, format="%,.0f"),
                        "환산비":      st.column_config.NumberColumn(min_value=0.01, step=1.0, format="%.0f"),
                        "ERP÷환산비":  st.column_config.NumberColumn(disabled=True, format="%.2f"),
                    },
                    use_container_width=True,
                    hide_index=True,
                    key="ratio_editor",
                )

                if st.button("🔄 환산 적용 후 재검증", type="primary"):
                    custom = dict(zip(edited["품목코드"], edited["환산비"]))
                    result = apply_custom_ratios(result, custom)
                    s = result["summary"]
                    st.success(f"재검증 완료 — 일치 {s['match_count']}건 ({s['match_pct']:.1f}%) / 불일치 {s['mismatch_count']}건")
                    st.session_state["result"] = result
                    st.rerun()

        st.markdown("---")

        # ── 탭 ─────────────────────────────────────────────────────────────────
        tab1, tab2, tab3 = st.tabs(["📋 전체 결과", "❌ 불일치 목록", "📊 날짜별 집계"])

        merged = result["merged"].copy()
        mismatch = result["mismatch"].copy()

        d1_col = f"D1 ({date_labels.get('D1', '')})"
        d2_col = f"D2 ({date_labels.get('D2', '')})"
        d3_col = f"D3 ({date_labels.get('D3', '')})"

        merged = merged.rename(columns={
            "D1": d1_col, "D2": d2_col, "D3": d3_col,
            "ERP합산": "합산(D1+D2+D3)",
        })
        mismatch = mismatch.rename(columns={
            "D1": d1_col, "D2": d2_col, "D3": d3_col,
            "ERP합산": "합산(D1+D2+D3)",
        })

        display_cols = [
            "품목코드", "품목명", "규격", "단위",
            d1_col, d2_col, d3_col,
            "합산(D1+D2+D3)", "Alps출고수량", "차이", "일치여부", "단위혼재",
        ]

        def style_row(row):
            if row["일치여부"]:
                return ["background-color: #e8f5e9"] * len(row)
            elif row["단위혼재"]:
                return ["background-color: #fff9c4"] * len(row)
            else:
                return ["background-color: #fce4d6"] * len(row)

        with tab1:
            display = merged[display_cols].copy()
            display["일치여부"] = display["일치여부"].map({True: "✓", False: "✗"})
            display["단위혼재"] = display["단위혼재"].map({True: "⚠", False: ""})
            st.dataframe(
                display.style.apply(style_row, axis=1),
                use_container_width=True,
                height=500,
            )

        with tab2:
            if mismatch.empty:
                st.success("불일치 항목이 없습니다! 🎉")
            else:
                st.warning(f"불일치 항목: {len(mismatch)}건 (ERP 기준 {s['mismatch_qty']:,.0f})")
                display_mm = mismatch[display_cols].copy()
                display_mm["일치여부"] = display_mm["일치여부"].map({True: "✓", False: "✗"})
                display_mm["단위혼재"] = display_mm["단위혼재"].map({True: "⚠", False: ""})
                st.dataframe(
                    display_mm.style.apply(style_row, axis=1),
                    use_container_width=True,
                    height=400,
                )

        with tab3:
            if d_map:
                st.markdown("#### ERP 날짜별 이동수량 합계")
                daily = {}
                for date, level in sorted(d_map.items()):
                    col_name = f"{level} ({date})"
                    if col_name in merged.columns:
                        daily[col_name] = merged[col_name].sum()

                daily_df = pd.DataFrame(
                    [{"날짜/구분": k, "이동수량 합계": v} for k, v in daily.items()]
                )
                st.table(daily_df.style.format({"이동수량 합계": "{:,.0f}"}))

        # ── 결과 다운로드 ──────────────────────────────────────────────────────
        st.markdown("---")
        excel_bytes = to_excel_bytes(result)
        st.download_button(
            label="📥 결과 Excel 다운로드",
            data=excel_bytes,
            file_name="ERP_Alps_검증결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.session_state["result"] = result
