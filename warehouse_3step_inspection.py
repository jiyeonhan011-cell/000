#!/usr/bin/env python3
"""
창고이동 3단계 검수 자동화
────────────────────────────────────────────────────────────────
Step 1. 이동처리 파일(F열) vs 라벨발행 파일(I열) 출고수량 비교
         매칭 키: ERP코드 + 날짜(이동처리 비고날짜 = 라벨발행 배송일)
         * 취소/변경 행은 라벨발행에서 자동 제외

Step 2. 라벨발행 파일(L열 급품목코드) vs 작업내역 파일(G열 제품코드) 수량 비교
         매칭 키: 급품목코드 + 배송일
         * 작업내역 파일은 창고이동 2회이므로 K열 수량 ÷ 2 적용

결과: 검수요약 / 1단계_수량불일치 / 1단계_창고이동만 / 1단계_라벨만
      2단계_수량불일치 / 2단계_라벨만 / 2단계_작업내역만
      1단계_일치 / 2단계_일치
────────────────────────────────────────────────────────────────
"""

import sys, os, re
from pathlib import Path
from collections import defaultdict

try:
    import xlrd, openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    os.system("pip install xlrd openpyxl -q")
    import xlrd, openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter


# ── 스타일 ───────────────────────────────────────────────────────
BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

def hdr(cell, bg="1F4E79"):
    cell.font  = Font(color="FFFFFF", bold=True, size=10)
    cell.fill  = PatternFill("solid", fgColor=bg)
    cell.alignment = CENTER
    cell.border = BORDER

def dat(cell, align=LEFT, bold=False, bg=None):
    cell.font      = Font(bold=bold, size=10)
    cell.alignment = align
    cell.border    = BORDER
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)


# ── 유틸 ─────────────────────────────────────────────────────────
def norm(name: str) -> str:
    """괄호 접두사·공백 정규화"""
    name = re.sub(r'^\([^)]+\)\s*', '', str(name))
    return re.sub(r'\s+', ' ', name).strip()

def clean_code(v) -> str:
    """코드에서 zero-width 문자 제거"""
    return re.sub(r'[​‌‍﻿]', '', str(v or '')).strip()


# ── STEP 1: 이동처리 파일 로드 ────────────────────────────────────
def load_warehouse(path):
    """TCT 케이터링 이동처리만 → (비고날짜, ERP코드) : {qty, 품목명, ...}"""
    try:
        wb = xlrd.open_workbook(path)
    except Exception as e:
        print(f"[오류] 이동처리 파일: {e}"); sys.exit(1)
    ws = wb.sheet_by_index(0)
    TARGET = "TCT 케이터링 이동처리"
    qty_map  = defaultdict(float)
    info_map = {}
    for i in range(1, ws.nrows):
        note = str(ws.cell_value(i, 8))
        if TARGET not in note:
            continue
        m = re.search(r'\((\d{4}-\d{2}-\d{2})\)', note)
        date = m.group(1) if m else str(ws.cell_value(i, 3)).strip()
        erp  = str(ws.cell_value(i, 9)).strip()
        name = str(ws.cell_value(i, 10)).strip()
        qty  = float(ws.cell_value(i, 5) or 0)
        key  = (date, erp)
        qty_map[key] += qty
        if key not in info_map:
            info_map[key] = {
                "품목코드": erp,
                "품목명":   name,
                "규격":     str(ws.cell_value(i, 11)).strip(),
                "단위":     str(ws.cell_value(i, 12)).strip(),
            }
    return qty_map, info_map


# ── STEP 1: 라벨발행 파일 로드 ────────────────────────────────────
def load_label(path):
    """
    취소/변경 제외 → (배송일, ERP코드) : {I출고수량, L급품목코드, ...}
    + (배송일, 급품목코드) : {I출고수량, ...}  ← Step2용
    """
    try:
        wb = openpyxl.load_workbook(path)
    except Exception as e:
        print(f"[오류] 라벨발행 파일: {e}"); sys.exit(1)
    ws = wb.active

    step1_qty  = defaultdict(float)
    step1_info = {}
    step2_qty  = defaultdict(float)
    step2_info = {}
    canceled = 0

    for i in range(2, ws.max_row + 1):
        state = ws.cell(i, 16).value
        if state in ('취소', '변경'):
            canceled += 1
            continue
        배송일   = str(ws.cell(i, 4).value  or '').strip()
        제품명   = str(ws.cell(i, 7).value  or '').strip()
        출고수량 = float(ws.cell(i, 9).value or 0)
        단위     = str(ws.cell(i, 10).value or '').strip()
        규격     = str(ws.cell(i, 11).value or '').strip()
        급코드   = clean_code(ws.cell(i, 12).value)   # L열
        erp코드  = str(ws.cell(i, 13).value  or '').strip()  # M열

        # Step1 키: 배송일 + ERP코드
        k1 = (배송일, erp코드)
        step1_qty[k1] += 출고수량
        if k1 not in step1_info:
            step1_info[k1] = {
                "ERP코드": erp코드, "급품목코드": 급코드,
                "제품명": 제품명, "규격": 규격, "단위": 단위,
            }

        # Step2 키: 배송일 + 급품목코드
        k2 = (배송일, 급코드)
        step2_qty[k2] += 출고수량
        if k2 not in step2_info:
            step2_info[k2] = {
                "급품목코드": 급코드, "ERP코드": erp코드,
                "제품명": 제품명, "규격": 규격, "단위": 단위,
            }

    print(f"  라벨발행 취소/변경 제외: {canceled}건")
    return step1_qty, step1_info, step2_qty, step2_info


# ── STEP 2: 작업내역 파일 로드 ───────────────────────────────────
def load_catering(path):
    """(배송일, G열제품코드) : {K열수량÷2, 제품명, ...}"""
    try:
        wb = openpyxl.load_workbook(path)
    except Exception as e:
        print(f"[오류] 작업내역 파일: {e}"); sys.exit(1)
    ws = wb.active
    qty_map  = defaultdict(float)
    info_map = {}
    for i in range(4, ws.max_row + 1):
        배송일 = str(ws.cell(i, 3).value or '').strip()
        if not 배송일 or 배송일 == '센터명':
            continue
        prod_code = clean_code(ws.cell(i, 7).value)
        prod_name = str(ws.cell(i, 8).value or '').strip()
        qty       = float(ws.cell(i, 11).value or 0)
        key = (배송일, prod_code)
        qty_map[key] += qty
        if key not in info_map:
            info_map[key] = {
                "제품코드": prod_code, "제품명": prod_name,
                "규격": str(ws.cell(i, 9).value or '').strip(),
                "단위": str(ws.cell(i, 10).value or '').strip(),
            }
    # 창고이동 2회 → ÷2
    qty_half = {k: v / 2 for k, v in qty_map.items()}
    return qty_half, info_map


# ── 비교 함수 ────────────────────────────────────────────────────
def compare(a_qty, a_info, b_qty, b_info, a_col, b_col):
    """
    두 dict를 비교하여 일치/수량불일치/a만/b만 분류.
    반환: (matched, qty_diff, a_only, b_only)
    """
    all_keys = set(a_qty) | set(b_qty)
    matched, qty_diff, a_only, b_only = [], [], [], []

    for key in sorted(all_keys):
        date, code = key
        aq = a_qty.get(key)
        bq = b_qty.get(key)
        ai = a_info.get(key, {})
        bi = b_info.get(key, {})

        name = (ai.get("품목명") or ai.get("제품명") or
                bi.get("품목명") or bi.get("제품명") or
                ai.get("ERP코드") or "")
        row = {
            "날짜":   date,
            "코드":   code,
            "품목명": name,
            "규격":   ai.get("규격") or bi.get("규격", ""),
            "단위":   ai.get("단위") or bi.get("단위", ""),
            a_col:    aq,
            b_col:    bq,
            "차이":   (bq or 0) - (aq or 0),
        }
        if aq is not None and bq is not None:
            if abs(aq - bq) < 0.001:
                matched.append(row)
            else:
                qty_diff.append(row)
        elif aq is not None:
            a_only.append(row)
        else:
            b_only.append(row)

    return matched, qty_diff, a_only, b_only


# ── xlsx 저장 ────────────────────────────────────────────────────
SHEET_DEFS = [
    # (sheet_title, hdr_bg, row_bg, col_headers, col_keys)
]

def write_sheet(wb, title, rows, hdr_bg, row_bg,
                col_label_a, col_label_b, col_key_a, col_key_b,
                extra_cols=None):
    """범용 시트 작성기"""
    ws = wb.create_sheet(title)
    base_headers = ["날짜", "코드", "품목명", "규격", "단위",
                    col_label_a, col_label_b, "차이"]
    base_keys    = ["날짜", "코드", "품목명", "규격", "단위",
                    col_key_a,  col_key_b,  "차이"]
    widths       = [13, 16, 38, 18, 7, 16, 16, 10]

    for c, h in enumerate(base_headers, 1):
        hdr(ws.cell(row=1, column=c, value=h), hdr_bg)

    for r, item in enumerate(rows, 2):
        bg = row_bg if r % 2 == 0 else None
        for c, key in enumerate(base_keys, 1):
            val = item.get(key)
            if val is None: val = "-"
            cell = ws.cell(row=r, column=c, value=val)
            align = CENTER if c in (1, 5, 6, 7, 8) else LEFT
            cell_bg = bg
            if key == "차이" and isinstance(val, (int, float)) and abs(val) > 0.001:
                cell_bg = "FFD7D7"
            dat(cell, align, bg=cell_bg)

    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(base_headers))}1"
    return ws


def save_all(s1_matched, s1_diff, s1_wh_only, s1_label_only,
             s2_matched, s2_diff, s2_label_only, s2_cat_only,
             output_path):
    wb = openpyxl.Workbook()

    # ── 요약 시트 ────────────────────────────────────────────────
    ws_sum = wb.active
    ws_sum.title = "검수요약"
    ws_sum.column_dimensions["A"].width = 42
    ws_sum.column_dimensions["B"].width = 16

    ws_sum.merge_cells("A1:B1")
    c = ws_sum.cell(1, 1, "창고이동 3단계 검수 결과")
    hdr(c, "1F4E79")

    s1_common = len(s1_matched) + len(s1_diff)
    s2_common = len(s2_matched) + len(s2_diff)

    rows = [
        ("── STEP 1: 이동처리 vs 라벨발행(출고수량) ──", ""),
        ("  전체 비교 항목 (날짜+ERP코드)",             len(s1_matched)+len(s1_diff)+len(s1_wh_only)+len(s1_label_only)),
        ("  ✅ 일치",                                   len(s1_matched)),
        ("  ❌ 수량 불일치 (양쪽 존재)",                 len(s1_diff)),
        ("  ⚠️  이동처리에만 있는 항목",                 len(s1_wh_only)),
        ("  ⚠️  라벨발행에만 있는 항목",                 len(s1_label_only)),
        ("  일치율 (공통항목 기준)",
         f"{len(s1_matched)/s1_common*100:.1f}%" if s1_common else "-"),
        ("", ""),
        ("── STEP 2: 라벨발행(L열) vs 작업내역(G열) ──", ""),
        ("  전체 비교 항목 (날짜+급품목코드)",             len(s2_matched)+len(s2_diff)+len(s2_label_only)+len(s2_cat_only)),
        ("  ✅ 일치",                                   len(s2_matched)),
        ("  ❌ 수량 불일치 (양쪽 존재)",                 len(s2_diff)),
        ("  ⚠️  라벨발행에만 있는 항목",                 len(s2_label_only)),
        ("  ⚠️  작업내역에만 있는 항목",                 len(s2_cat_only)),
        ("  일치율 (공통항목 기준)",
         f"{len(s2_matched)/s2_common*100:.1f}%" if s2_common else "-"),
    ]
    for r, (label, val) in enumerate(rows, 2):
        c1 = ws_sum.cell(r, 1, label)
        c2 = ws_sum.cell(r, 2, val if val != "" else None)
        bold = label.startswith("──")
        bg = "D9E1F2" if bold else None
        dat(c1, LEFT,   bold=bold, bg=bg)
        dat(c2, CENTER, bold=bold, bg=bg)

    # ── STEP 1 시트들 ────────────────────────────────────────────
    write_sheet(wb, "1단계_수량불일치", s1_diff,
                "C00000", "FCE4D6",
                "이동처리(F열)", "라벨발행(I열)",
                "이동처리(F열)", "라벨발행(I열)")

    write_sheet(wb, "1단계_이동처리만", s1_wh_only,
                "843C0C", "FFF2CC",
                "이동처리(F열)", "라벨발행(I열)",
                "이동처리(F열)", "라벨발행(I열)")

    write_sheet(wb, "1단계_라벨발행만", s1_label_only,
                "7030A0", "EAD1F5",
                "이동처리(F열)", "라벨발행(I열)",
                "이동처리(F열)", "라벨발행(I열)")

    write_sheet(wb, "1단계_일치", s1_matched,
                "375623", "E2EFDA",
                "이동처리(F열)", "라벨발행(I열)",
                "이동처리(F열)", "라벨발행(I열)")

    # ── STEP 2 시트들 ────────────────────────────────────────────
    write_sheet(wb, "2단계_수량불일치", s2_diff,
                "C00000", "FCE4D6",
                "라벨발행(I열)", "작업내역(K÷2)",
                "라벨발행(I열)", "작업내역(K÷2)")

    write_sheet(wb, "2단계_라벨발행만", s2_label_only,
                "843C0C", "FFF2CC",
                "라벨발행(I열)", "작업내역(K÷2)",
                "라벨발행(I열)", "작업내역(K÷2)")

    write_sheet(wb, "2단계_작업내역만", s2_cat_only,
                "7030A0", "EAD1F5",
                "라벨발행(I열)", "작업내역(K÷2)",
                "라벨발행(I열)", "작업내역(K÷2)")

    write_sheet(wb, "2단계_일치", s2_matched,
                "375623", "E2EFDA",
                "라벨발행(I열)", "작업내역(K÷2)",
                "라벨발행(I열)", "작업내역(K÷2)")

    wb.save(output_path)
    print(f"\n저장 완료: {output_path}")


# ── 메인 ─────────────────────────────────────────────────────────
def main():
    WH_PATH  = "/root/.claude/uploads/4b31195f-0362-5c7f-baf9-3f04e5c8214d/21498ca7-20260608133431.xls"
    LBL_PATH = "/root/.claude/uploads/4b31195f-0362-5c7f-baf9-3f04e5c8214d/cdbc6ee1-Alps_________260609_1144.xlsx"
    CAT_PATH = "/root/.claude/uploads/4b31195f-0362-5c7f-baf9-3f04e5c8214d/544784f8-_____20260608_1420.xlsx"
    OUT_PATH = "창고이동_3단계검수결과.xlsx"

    if len(sys.argv) >= 4:
        WH_PATH, LBL_PATH, CAT_PATH = sys.argv[1], sys.argv[2], sys.argv[3]
    if len(sys.argv) >= 5:
        OUT_PATH = sys.argv[4]

    print("=" * 60)
    print("[STEP 1] 이동처리 파일 로드 중...")
    wh_qty, wh_info = load_warehouse(WH_PATH)
    print(f"  TCT 케이터링 이동처리 항목: {len(wh_qty)}건 (날짜+ERP코드 기준)")

    print("\n[STEP 1] 라벨발행 파일 로드 중 (취소/변경 자동 제외)...")
    lbl_s1_qty, lbl_s1_info, lbl_s2_qty, lbl_s2_info = load_label(LBL_PATH)
    print(f"  유효 라벨 항목(Step1): {len(lbl_s1_qty)}건")
    print(f"  유효 라벨 항목(Step2): {len(lbl_s2_qty)}건")

    print("\n[STEP 2] 작업내역 파일 로드 중 (수량 ÷2 적용)...")
    cat_qty, cat_info = load_catering(CAT_PATH)
    print(f"  작업내역 항목: {len(cat_qty)}건 (날짜+제품코드 기준)")

    print("\n[비교 수행]")
    s1_matched, s1_diff, s1_wh_only, s1_label_only = compare(
        wh_qty, wh_info, lbl_s1_qty, lbl_s1_info,
        "이동처리(F열)", "라벨발행(I열)"
    )
    s2_matched, s2_diff, s2_label_only, s2_cat_only = compare(
        lbl_s2_qty, lbl_s2_info, cat_qty, cat_info,
        "라벨발행(I열)", "작업내역(K÷2)"
    )

    s1_common = len(s1_matched) + len(s1_diff)
    s2_common = len(s2_matched) + len(s2_diff)

    print("\n" + "=" * 60)
    print("[ STEP 1 결과: 이동처리 vs 라벨발행 ]")
    print(f"  ✅ 일치:          {len(s1_matched):>5}건  ({len(s1_matched)/s1_common*100:.1f}%, 공통기준)" if s1_common else "")
    print(f"  ❌ 수량불일치:    {len(s1_diff):>5}건  ({len(s1_diff)/s1_common*100:.1f}%)" if s1_common else "")
    print(f"  ⚠️  이동처리만:   {len(s1_wh_only):>5}건")
    print(f"  ⚠️  라벨발행만:   {len(s1_label_only):>5}건")

    print("\n[ STEP 2 결과: 라벨발행(L) vs 작업내역(G) ]")
    print(f"  ✅ 일치:          {len(s2_matched):>5}건  ({len(s2_matched)/s2_common*100:.1f}%, 공통기준)" if s2_common else "")
    print(f"  ❌ 수량불일치:    {len(s2_diff):>5}건  ({len(s2_diff)/s2_common*100:.1f}%)" if s2_common else "")
    print(f"  ⚠️  라벨발행만:   {len(s2_label_only):>5}건")
    print(f"  ⚠️  작업내역만:   {len(s2_cat_only):>5}건")
    print("=" * 60)

    save_all(s1_matched, s1_diff, s1_wh_only, s1_label_only,
             s2_matched, s2_diff, s2_label_only, s2_cat_only,
             OUT_PATH)


if __name__ == "__main__":
    main()
