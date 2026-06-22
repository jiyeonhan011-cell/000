#!/usr/bin/env python3
"""
창고이동 3단계 검수 - Streamlit 버전
실행: streamlit run app.py
"""

import os, re, sys, io, tempfile
from pathlib import Path
from collections import defaultdict

try:
    import streamlit as st
    import xlrd, openpyxl
    from openpyxl.styles import Font as XFont, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    os.system(f"{sys.executable} -m pip install streamlit xlrd openpyxl -q")
    import streamlit as st
    import xlrd, openpyxl
    from openpyxl.styles import Font as XFont, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

# ── 검수 로직 ──────────────────────────────────────────────────────

BORDER_STYLE = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)

def clean_code(v):
    return re.sub(u'[​‌‍﻿ ]', '', str(v or '')).strip()



# 파일 간 단위가 달라 수량이 다르게 보이는 품목 — 환산비 적용 후 일치하면 "단위다름(정상)"
# {코드: (환산배수, 설명)}  — 한 쪽 수량 * 배수 ≈ 다른 쪽 수량이면 정상
UNIT_DIFF_NORMAL = {
    "NA603095":    (168, "새찬 오이피클 일회용/중국산 80g (1BOX=168EA)"),
    "1000464464":  (168, "새찬 오이피클 일회용/중국산 80g (1BOX=168EA) - 작업내역코드"),
    "153325":      (168, "새찬 오이피클 일회용/중국산 80g (1BOX=168EA) - 작업내역코드"),
    "482644":      (168, "새찬 오이피클 일회용/중국산 80g (1BOX=168EA) - 작업내역코드"),
    "NB701107":    (30,  "청수식품 맑은물 냉면육수 300g/동치미맛 (1BOX=30EA)"),
    "NB706070":    (30,  "영미 동치미 냉면육수 310g (1BOX=30EA)"),
    "NE101062":    (6,   "삼원 튀김스프링 회오리감자 700g/냉동 (1BOX=6PAK)"),
    "NE101075":    (6,   "화영푸드 회오리감자 800g/냉동 (1BOX=6PAK)"),
    "NH101013":    (10,  "푸드림 미니바 슈가스틱 (1BOX=10EA)"),
    "NH512021":    (10,  "식예원 가쓰오맛 후리가께 50g (1PAC=10EA)"),
    "NH610019":    (20,  "신진 신화당 50g (1PAK=20EA)"),
    "NI307040":    (24,  "진로 하이트제로 350ml/무알콜 (1BOX=24EA)"),
    "NI309007":    (12,  "랭거스 망고쥬스 449ml (1BOX=12EA)"),
    "NL102043":    (2,   "유진 카사바칩 1.2kg (1BOX=2EA)"),
}

def _unit_diff_ok(row, col_a, col_b):
    """환산비 적용 후 두 수량이 일치하는지 확인"""
    info = UNIT_DIFF_NORMAL.get(row["코드"])
    if not info: return False
    factor = info[0]
    a = row.get(col_a) or 0
    b = row.get(col_b) or 0
    return abs(a * factor - b) < 0.001 or abs(b * factor - a) < 0.001

def load_warehouse(path):
    wb  = xlrd.open_workbook(path)
    ws  = wb.sheet_by_index(0)
    TARGET = "TCT 케이터링 이동처리"
    qty_map, info_map, dates_map = defaultdict(float), {}, defaultdict(set)
    for i in range(1, ws.nrows):
        note = str(ws.cell_value(i, 8))
        if TARGET not in note: continue
        m    = re.search(r'\((\d{4}-\d{2}-\d{2})\)', note)
        date = m.group(1) if m else str(ws.cell_value(i, 3)).strip()
        erp  = str(ws.cell_value(i, 9)).strip()
        qty  = float(ws.cell_value(i, 5) or 0)
        qty_map[erp] += qty
        dates_map[erp].add(date)
        if erp not in info_map:
            info_map[erp] = {
                "품목명": str(ws.cell_value(i,10)).strip(),
                "규격":   str(ws.cell_value(i,11)).strip(),
                "단위":   str(ws.cell_value(i,12)).strip(),
            }
    for erp in info_map:
        info_map[erp]["이동날짜"] = ", ".join(sorted(dates_map[erp]))
    return qty_map, info_map

BOX_UNITS = {'BOX','PAC','PAK','B','박스'}

def _to_ea(qty, 단위, code):
    """단위가 BOX계열이고 환산배수가 등록된 품목이면 EA로 환산"""
    if 단위.upper() in BOX_UNITS and code in UNIT_DIFF_NORMAL:
        return qty * UNIT_DIFF_NORMAL[code][0]
    return qty

def load_label(path):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    s1_qty, s1_qty_cmp = defaultdict(float), defaultdict(float)
    s2_qty, s2_qty_cmp = defaultdict(float), defaultdict(float)
    s1_info, s2_info = {}, {}
    canceled = 0
    for i in range(2, ws.max_row + 1):
        state = ws.cell(i,16).value
        if state in ('취소','변경'):
            canceled += 1; continue
        제품명   = str(ws.cell(i,7).value  or '').strip()
        try:
            출고수량 = float(ws.cell(i,9).value or 0)
        except (ValueError, TypeError):
            continue
        단위     = str(ws.cell(i,10).value or '').strip()
        규격     = str(ws.cell(i,11).value or '').strip()
        급코드   = clean_code(ws.cell(i,12).value)
        erp코드  = str(ws.cell(i,13).value or '').strip()
        s1_qty[erp코드] += 출고수량                              # 원본 (합계 표시용)
        s1_qty_cmp[erp코드] += _to_ea(출고수량, 단위, erp코드)  # 환산 (비교용)
        if erp코드 not in s1_info:
            s1_info[erp코드] = {"ERP코드":erp코드,"급품목코드":급코드,"제품명":제품명,"규격":규격,"단위":단위}
        s2_qty[급코드] += 출고수량
        s2_qty_cmp[급코드] += _to_ea(출고수량, 단위, 급코드)
        if 급코드 not in s2_info:
            s2_info[급코드] = {"급품목코드":급코드,"ERP코드":erp코드,"제품명":제품명,"규격":규격,"단위":단위}
    return s1_qty, s1_qty_cmp, s1_info, s2_qty, s2_qty_cmp, s2_info, canceled

def load_catering(path):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    qty_map, info_map = defaultdict(float), {}
    for i in range(4, ws.max_row + 1):
        d = ws.cell(i,3).value
        if not d or str(d)=='센터명': continue
        code = clean_code(ws.cell(i,7).value)
        qty  = float(ws.cell(i,11).value or 0)
        단위  = str(ws.cell(i,10).value or '').strip()
        qty_map[code] += qty
        if code not in info_map:
            info_map[code] = {
                "제품코드":code,"제품명":str(ws.cell(i,8).value or '').strip(),
                "규격":str(ws.cell(i,9).value or '').strip(),
                "단위":단위,
            }
    return {k:v/2 for k,v in qty_map.items()}, info_map

def load_prework(path):
    if not path or not os.path.exists(path): return {}, set()
    wb = openpyxl.load_workbook(path)
    # 선작업 코드
    items = {}
    if '선작업-pre_qty' in wb.sheetnames:
        ws = wb['선작업-pre_qty']
        for i in range(2, ws.max_row + 1):
            code = clean_code(ws.cell(i,2).value)
            if not code: continue
            if code not in items:
                items[code] = {"품목명":str(ws.cell(i,3).value or ''),"급식사":str(ws.cell(i,1).value or ''),"선작업qty":0.0,"보정후":0.0}
            items[code]["선작업qty"] += float(ws.cell(i,7).value or 0)
            items[code]["보정후"]    += float(ws.cell(i,8).value or 0)
    # BOX-EA 환산 코드
    box_ea = set()
    if 'BOX-EA환산' in wb.sheetnames:
        ws2 = wb['BOX-EA환산']
        for i in range(2, ws2.max_row + 1):
            code = clean_code(ws2.cell(i,2).value)
            if code: box_ea.add(code)
    return items, box_ea

def compare(a_qty, a_info, b_qty, b_info, col_a, col_b):
    matched, qty_diff, a_only, b_only = [], [], [], []
    for code in sorted(set(a_qty)|set(b_qty)):
        aq, bq = a_qty.get(code), b_qty.get(code)
        ai, bi = a_info.get(code,{}), b_info.get(code,{})
        name = ai.get("품목명") or ai.get("제품명") or bi.get("품목명") or bi.get("제품명") or ""
        row = {"코드":code,"품목명":name,
               "규격":ai.get("규격") or bi.get("규격",""),
               "단위":ai.get("단위") or bi.get("단위",""),
               "이동날짜":ai.get("이동날짜",""),
               col_a:aq, col_b:bq, "차이":(bq or 0)-(aq or 0)}
        if aq is not None and bq is not None:
            (matched if abs(aq-bq)<0.001 else qty_diff).append(row)
        elif aq is not None: a_only.append(row)
        else: b_only.append(row)
    return matched, qty_diff, a_only, b_only

def split_prework(matched, qty_diff, lbl_only, cat_only, prework_codes, lbl_col):
    def _split(rows):
        pre, normal = [], []
        for row in rows:
            if row["코드"] in prework_codes:
                r = dict(row)
                r["선작업qty"] = prework_codes[row["코드"]]["선작업qty"]
                r["보정후"]    = prework_codes[row["코드"]]["보정후"]
                r["급식사"]    = prework_codes[row["코드"]]["급식사"]
                lbl_qty = row.get(lbl_col) or 0
                r["보정후차이"] = lbl_qty - r["보정후"]
                pre.append(r)
            else: normal.append(row)
        return pre, normal
    pm,m = _split(matched); pd,d = _split(qty_diff)
    pl,l = _split(lbl_only); pc,c = _split(cat_only)
    return pm+pd+pl+pc, m, d, l, c

def hdr_style(cell, bg="1F4E79"):
    cell.font  = XFont(color="FFFFFF", bold=True, size=10)
    cell.fill  = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER_STYLE

def dat_style(cell, center=False, bold=False, bg=None):
    cell.font      = XFont(bold=bold, size=10)
    cell.alignment = Alignment(horizontal="center" if center else "left", vertical="center", wrap_text=True)
    cell.border    = BORDER_STYLE
    if bg: cell.fill = PatternFill("solid", fgColor=bg)

def write_xl_sheet(wb, title, rows, hdr_bg, row_bg, col_a, col_b, show_date=True, extra_cols=None):
    ws = wb.create_sheet(title)
    base_h = ["코드","품목명","규격","단위"]
    base_k = ["코드","품목명","규격","단위"]
    if show_date: base_h.append("이동날짜"); base_k.append("이동날짜")
    base_h += [col_a, col_b, "차이"]
    base_k += [col_a, col_b, "차이"]
    if extra_cols:
        for eh,ek in extra_cols: base_h.append(eh); base_k.append(ek)
    w_base = [16,38,18,7]
    if show_date: w_base.append(22)
    w_base += [16,16,10]
    if extra_cols: w_base += [14]*len(extra_cols)
    for c,h in enumerate(base_h,1): hdr_style(ws.cell(1,c,h), hdr_bg)
    for r,item in enumerate(rows,2):
        bg = row_bg if r%2==0 else None
        for c,key in enumerate(base_k,1):
            val = item.get(key); val = "-" if val is None else val
            cell = ws.cell(r,c,val)
            center = key in ("코드","단위",col_a,col_b,"차이","이동날짜","선작업qty","보정후","보정후차이")
            cell_bg = bg
            if key in ("차이","보정후차이") and isinstance(val,(int,float)) and abs(val)>0.001: cell_bg="FFD7D7"
            dat_style(cell, center=center, bg=cell_bg)
    for c,w in enumerate(w_base,1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(base_h))}1"


def run_inspection(wh_path, lbl_path, cat_path, pre_path):
    wh_qty, wh_info = load_warehouse(wh_path)
    ls1q, ls1q_cmp, ls1i, ls2q, ls2q_cmp, ls2i, canceled = load_label(lbl_path)
    cat_qty, cat_info = load_catering(cat_path)
    prework, box_ea_codes = load_prework(pre_path) if pre_path else ({}, set())

    def is_box_ea(row):
        return row["코드"] in box_ea_codes or "혼용코드" in str(row.get("품목명",""))

    def restore_orig(rows, col_display, orig_qty, col_a, col_b):
        """비교는 환산값으로 했지만 표시는 원본값으로 복원"""
        for r in rows:
            orig = orig_qty.get(r["코드"])
            if orig is not None:
                r[col_display] = orig
                r["차이"] = (r.get(col_b) or 0) - (r.get(col_a) or 0)
        return rows

    s1m,s1d,s1w,s1l = compare(wh_qty,wh_info,ls1q_cmp,ls1i,"이동처리(F열)","라벨발행(I열)")
    for lst in (s1m, s1d, s1w, s1l):
        restore_orig(lst, "라벨발행(I열)", ls1q, "이동처리(F열)", "라벨발행(I열)")

    # BOX-EA 환산 품목을 수량불일치에서 분리 (파일 코드 + 품목명 혼용코드 자동감지)
    s1box = [r for r in s1d if is_box_ea(r)]
    s1d   = [r for r in s1d if not is_box_ea(r)]
    # 단위 다름: 환산비 적용 후 수량 일치하면 정상, 불일치하면 수량불일치로 유지
    s1unit = [r for r in s1d if _unit_diff_ok(r, "이동처리(F열)", "라벨발행(I열)")]
    s1d    = [r for r in s1d if not _unit_diff_ok(r, "이동처리(F열)", "라벨발행(I열)")]

    s2m_all,s2d_all,s2l_all,s2c_all = compare(ls2q,ls2i,cat_qty,cat_info,"라벨발행(I열)","작업내역(K÷2)")

    if prework:
        s2pre,s2m,s2d,s2l,s2c = split_prework(s2m_all,s2d_all,s2l_all,s2c_all,prework,"라벨발행(I열)")
    else:
        s2pre,s2m,s2d,s2l,s2c = [],s2m_all,s2d_all,s2l_all,s2c_all

    # BOX-EA 환산 품목을 2단계 수량불일치에서도 분리 (파일 코드 + 품목명 혼용코드 자동감지)
    s2box = [r for r in s2d if is_box_ea(r)]
    s2d   = [r for r in s2d if not is_box_ea(r)]
    # 단위 다름: 환산비 적용 후 수량 일치하면 정상, 불일치하면 수량불일치로 유지
    s2unit = [r for r in s2d if _unit_diff_ok(r, "라벨발행(I열)", "작업내역(K÷2)")]
    s2d    = [r for r in s2d if not _unit_diff_ok(r, "라벨발행(I열)", "작업내역(K÷2)")]

    wb = openpyxl.Workbook()
    ws_sum = wb.active; ws_sum.title = "검수요약"
    ws_sum.column_dimensions["A"].width=46; ws_sum.column_dimensions["B"].width=16
    ws_sum.merge_cells("A1:B1")
    hdr_style(ws_sum.cell(1,1,"창고이동 3단계 검수 결과"), "1F4E79")
    s1c = len(s1m)+len(s1d); s2c_cnt = len(s2m)+len(s2d)
    summary_rows = [
        ("── STEP 1: 이동처리(F열) vs 라벨발행(I열) ──",""),
        ("  전체 항목 (ERP코드 기준)", len(s1m)+len(s1d)+len(s1w)+len(s1l)+len(s1box)+len(s1unit)),
        ("  ✅ 일치", len(s1m)),("  ❌ 수량 불일치", len(s1d)),
        ("  ✅ BOX/EA환산(정상)", len(s1box)),
        ("  ✅ 단위다름(정상)", len(s1unit)),
        ("  ⚠️ 이동처리에만", len(s1w)),("  ⚠️ 라벨발행에만", len(s1l)),
        ("  일치율(공통기준)", f"{len(s1m)/s1c*100:.1f}%" if s1c else "-"),
        ("",""),
        ("── STEP 2: 라벨발행(L열) vs 작업내역(G열) ──",""),
        ("  전체 항목 (급품목코드 기준)", len(s2m)+len(s2d)+len(s2l)+len(s2c)+len(s2pre)+len(s2box)+len(s2unit)),
        ("  ✅ 일치", len(s2m)),("  ❌ 수량 불일치", len(s2d)),
        ("  ✅ BOX/EA환산(정상)", len(s2box)),
        ("  ✅ 단위다름(정상)", len(s2unit)),
        ("  ⚠️ 라벨발행에만", len(s2l)),("  ⚠️ 작업내역에만", len(s2c)),
        ("  ✅ 선작업(정상)", len(s2pre)),
        ("  일치율(공통기준)", f"{len(s2m)/s2c_cnt*100:.1f}%" if s2c_cnt else "-"),
    ]
    for r,(label,val) in enumerate(summary_rows,2):
        bold = label.startswith("──"); bg="D9E1F2" if bold else None
        dat_style(ws_sum.cell(r,1,label), bold=bold, bg=bg)
        dat_style(ws_sum.cell(r,2,val if val!="" else None), center=True, bold=bold, bg=bg)

    A,B = "이동처리(F열)","라벨발행(I열)"
    C,D = "라벨발행(I열)","작업내역(K÷2)"
    write_xl_sheet(wb,"1단계_수량불일치",s1d,"C00000","FCE4D6",A,B)
    write_xl_sheet(wb,"1단계_이동처리만",s1w,"843C0C","FFF2CC",A,B)
    write_xl_sheet(wb,"1단계_라벨발행만",s1l,"7030A0","EAD1F5",A,B,show_date=False)
    write_xl_sheet(wb,"1단계_일치",      s1m,"375623","E2EFDA",A,B)
    if s1box:
        write_xl_sheet(wb,"1단계_BOX_EA환산(정상)",s1box,"4472C4","DAE8FC",A,B)
    write_xl_sheet(wb,"2단계_수량불일치",s2d,"C00000","FCE4D6",C,D,show_date=False)
    write_xl_sheet(wb,"2단계_라벨발행만",s2l,"843C0C","FFF2CC",C,D,show_date=False)
    write_xl_sheet(wb,"2단계_작업내역만",s2c,"7030A0","EAD1F5",C,D,show_date=False)
    write_xl_sheet(wb,"2단계_일치",      s2m,"375623","E2EFDA",C,D,show_date=False)
    if s2pre:
        write_xl_sheet(wb,"2단계_선작업(정상)",s2pre,"4472C4","DAE8FC",C,D,show_date=False,
                       extra_cols=[("급식사","급식사"),("선작업qty","선작업qty"),("보정후","보정후"),("보정후차이","보정후차이")])
    if s2box:
        write_xl_sheet(wb,"2단계_BOX_EA환산(정상)",s2box,"4472C4","DAE8FC",C,D,show_date=False)
    if s1unit:
        write_xl_sheet(wb,"1단계_단위다름(정상)",s1unit,"6B6B6B","EFEFEF",A,B)
    if s2unit:
        write_xl_sheet(wb,"2단계_단위다름(정상)",s2unit,"6B6B6B","EFEFEF",C,D,show_date=False)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    A,B = "이동처리(F열)","라벨발행(I열)"
    C,D = "라벨발행(I열)","작업내역(K÷2)"
    def sumq(rows, col): return sum(r[col] for r in rows if isinstance(r.get(col), (int,float)))
    all_s1 = s1m+s1d+s1w+s1l+s1box+s1unit
    all_s2 = s2m+s2d+s2l+s2c+s2pre+s2box+s2unit

    return buf, {
        "s1_matched":len(s1m),"s1_diff":len(s1d),"s1_wh":len(s1w),"s1_lbl":len(s1l),
        "s1_box":len(s1box),"s1_unit":len(s1unit),
        "s1_rate":f"{len(s1m)/s1c*100:.1f}" if s1c else "0",
        "s1_wh_qty": sumq(all_s1, A), "s1_lbl_qty": sumq(all_s1, B),
        "s2_matched":len(s2m),"s2_diff":len(s2d),"s2_lbl":len(s2l),"s2_cat":len(s2c),
        "s2_pre":len(s2pre), "s2_box":len(s2box),"s2_unit":len(s2unit),
        "s2_rate":f"{len(s2m)/s2c_cnt*100:.1f}" if s2c_cnt else "0",
        "s2_lbl_qty": sumq(all_s2, C), "s2_cat_qty": sumq(all_s2, D),
        "canceled": canceled,
        # 실제 데이터
        "rows": {
            "s1_diff": s1d, "s1_wh": s1w, "s1_lbl": s1l, "s1_box": s1box, "s1_matched": s1m,
            "s1_unit": s1unit,
            "s2_diff": s2d, "s2_lbl": s2l, "s2_cat": s2c, "s2_pre": s2pre, "s2_matched": s2m,
            "s2_box": s2box, "s2_unit": s2unit,
        },
        "cols": {"A": A, "B": B, "C": C, "D": D},
    }


# ── Streamlit UI ───────────────────────────────────────────────────

st.set_page_config(
    page_title="창고이동 3단계 검수",
    page_icon="📦",
    layout="wide",
)

st.markdown("""
<style>
    .main { max-width: 900px; margin: 0 auto; }
    .stFileUploader > label { font-size: 1rem; font-weight: 600; }
    div[data-testid="metric-container"] {
        background: #f0f4ff; border-radius: 8px; padding: 12px;
    }
    .result-box {
        background: #f8f9fa; border-radius: 10px; padding: 20px;
        border: 1px solid #dee2e6; margin-top: 16px;
    }
</style>
""", unsafe_allow_html=True)

st.title("📦 창고이동 3단계 검수 프로그램")
st.caption("이동처리 → 라벨발행 → 작업내역 자동 비교")

st.divider()

def find_latest(folder, prefix):
    """폴더에서 prefix로 시작하는 파일 중 가장 최근 것 반환"""
    p = Path(folder)
    if not p.is_dir():
        return None
    files = [f for f in p.iterdir() if f.name.startswith(prefix) and f.suffix.lower() in ('.xls','.xlsx')]
    return str(max(files, key=lambda f: f.stat().st_mtime)) if files else None

CONFIG_FILE = Path(__file__).parent / "config.json"

def load_config():
    if CONFIG_FILE.exists():
        import json
        try: return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except: pass
    return {}

def save_config(data):
    import json
    CONFIG_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

cfg = load_config()
saved_path = cfg.get("folder_path", "")

folder_path = st.text_input(
    "📁 검수 파일 폴더 경로",
    value=saved_path,
    placeholder="예: C:\\검수파일  또는  C:/검수파일",
    help="이동처리 / 라벨발행 / 작업내역 / 선작업 하위폴더가 있는 상위 폴더 경로를 입력하세요"
)

if folder_path and folder_path != saved_path:
    save_config({**cfg, "folder_path": folder_path})

if folder_path:
    for sub in ("이동처리", "라벨발행", "작업내역", "선작업"):
        (Path(folder_path) / sub).mkdir(parents=True, exist_ok=True)

wh_path = lbl_path = cat_path = pre_path = None
if folder_path:
    wh_path  = find_latest(Path(folder_path) / "이동처리",  "이동처리")
    lbl_path = find_latest(Path(folder_path) / "라벨발행",  "라벨발행")
    cat_path = find_latest(Path(folder_path) / "작업내역",  "작업내역")
    pre_path = find_latest(Path(folder_path) / "선작업",    "선작업")

    c1, c2, c3, c4 = st.columns(4)
    c1.success(f"✓ {Path(wh_path).name}")  if wh_path  else c1.error("이동처리 파일 없음")
    c2.success(f"✓ {Path(lbl_path).name}") if lbl_path else c2.error("라벨발행 파일 없음")
    c3.success(f"✓ {Path(cat_path).name}") if cat_path else c3.error("작업내역 파일 없음")
    c4.info(f"✓ {Path(pre_path).name}")    if pre_path else c4.info("선작업 파일 없음 (선택)")

st.divider()

ready = bool(wh_path and lbl_path and cat_path)
run_btn = st.button("🔍 검수 시작", type="primary", use_container_width=True, disabled=not ready)

if not folder_path:
    st.warning("위에 폴더 경로를 입력하면 파일을 자동으로 찾아드립니다.")
elif not ready:
    st.warning("이동처리 / 라벨발행 / 작업내역 파일을 찾지 못했습니다. 폴더 구조를 확인해주세요.")

if run_btn:
    with st.spinner("검수 중..."):
        try:
            buf, result = run_inspection(wh_path, lbl_path, cat_path, pre_path)

            import pandas as pd

            def to_df(rows, cols, extra=None):
                if not rows: return None
                keys = ["코드","품목명","규격","단위"] + [c for c in cols if c in rows[0]] + ["차이"]
                if extra: keys += [k for k in extra if k in rows[0]]
                keys = [k for k in keys if k in rows[0]]
                return pd.DataFrame([{k: r.get(k,"-") for k in keys} for r in rows])

            st.success("✅ 검수 완료!")
            rows = result["rows"]
            A,B = result["cols"]["A"], result["cols"]["B"]
            C,D = result["cols"]["C"], result["cols"]["D"]

            # ── STEP 1 ──
            st.markdown("### STEP 1: 이동처리 vs 라벨발행")
            c1,c2 = st.columns(2)
            c1.metric("이동처리 이동수량 합계", f"{result['s1_wh_qty']:,.0f}")
            c2.metric("라벨발행 출고수량 합계", f"{result['s1_lbl_qty']:,.0f}",
                      delta=f"{result['s1_lbl_qty']-result['s1_wh_qty']:+,.0f}")
            c1,c2,c3,c4,c5,c6 = st.columns(6)
            c1.metric("✅ 일치",          result["s1_matched"])
            c2.metric("❌ 수량불일치",    result["s1_diff"])
            c3.metric("✅ BOX/EA환산",    result["s1_box"])
            c4.metric("✅ 단위다름(정상)", result["s1_unit"])
            c5.metric("⚠️ 이동처리만",   result["s1_wh"])
            c6.metric("⚠️ 라벨발행만",   result["s1_lbl"])
            st.caption(f"일치율: {result['s1_rate']}%  |  취소/변경 제외: {result['canceled']}건")

            if rows["s1_diff"]:
                with st.expander(f"❌ 수량불일치 품목 ({result['s1_diff']}건)", expanded=True):
                    st.dataframe(to_df(rows["s1_diff"], [A,B]), use_container_width=True, hide_index=True)
            if rows["s1_wh"]:
                with st.expander(f"⚠️ 이동처리에만 있는 품목 ({result['s1_wh']}건)"):
                    st.dataframe(to_df(rows["s1_wh"], [A,B]), use_container_width=True, hide_index=True)
            if rows["s1_lbl"]:
                with st.expander(f"⚠️ 라벨발행에만 있는 품목 ({result['s1_lbl']}건)"):
                    st.dataframe(to_df(rows["s1_lbl"], [A,B]), use_container_width=True, hide_index=True)
            if rows["s1_box"]:
                s1box_err = [r for r in rows["s1_box"] if abs(r.get("차이") or 0) > 0.001]
                with st.expander(f"✅ BOX/EA환산 ({result['s1_box']}건) {'⚠️ 수량 확인 필요 ' + str(len(s1box_err)) + '건' if s1box_err else ''}",
                                 expanded=bool(s1box_err)):
                    if s1box_err:
                        st.warning("아래 항목은 차이가 있습니다. 환산 오류 여부를 확인해주세요.")
                    st.dataframe(to_df(rows["s1_box"], [A,B]), use_container_width=True, hide_index=True)
            if rows["s1_unit"]:
                with st.expander(f"✅ 단위다름(정상) ({result['s1_unit']}건)"):
                    st.caption("파일 간 단위가 달라 수량이 다르게 표시되지만 정상 처리된 품목입니다.")
                    st.dataframe(to_df(rows["s1_unit"], [A,B]), use_container_width=True, hide_index=True)

            st.divider()

            # ── STEP 2 ──
            st.markdown("### STEP 2: 라벨발행 vs 작업내역")
            c1,c2 = st.columns(2)
            c1.metric("라벨발행 출고수량 합계", f"{result['s2_lbl_qty']:,.0f}")
            c2.metric("작업내역 수량 합계(÷2)", f"{result['s2_cat_qty']:,.0f}",
                      delta=f"{result['s2_cat_qty']-result['s2_lbl_qty']:+,.0f}")
            c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
            c1.metric("✅ 일치",          result["s2_matched"])
            c2.metric("❌ 수량불일치",    result["s2_diff"])
            c3.metric("✅ BOX/EA환산",    result["s2_box"])
            c4.metric("✅ 단위다름(정상)", result["s2_unit"])
            c5.metric("⚠️ 라벨발행만",   result["s2_lbl"])
            c6.metric("⚠️ 작업내역만",   result["s2_cat"])
            c7.metric("✅ 선작업",        result["s2_pre"])
            st.caption(f"일치율: {result['s2_rate']}%")

            if rows["s2_diff"]:
                with st.expander(f"❌ 수량불일치 품목 ({result['s2_diff']}건)", expanded=True):
                    st.dataframe(to_df(rows["s2_diff"], [C,D]), use_container_width=True, hide_index=True)
            if rows["s2_lbl"]:
                with st.expander(f"⚠️ 라벨발행에만 있는 품목 ({result['s2_lbl']}건)"):
                    st.dataframe(to_df(rows["s2_lbl"], [C,D]), use_container_width=True, hide_index=True)
            if rows["s2_cat"]:
                with st.expander(f"⚠️ 작업내역에만 있는 품목 ({result['s2_cat']}건)"):
                    st.dataframe(to_df(rows["s2_cat"], [C,D]), use_container_width=True, hide_index=True)
            if rows["s2_pre"]:
                with st.expander(f"✅ 선작업(정상) ({result['s2_pre']}건)"):
                    st.caption("보정후차이 = 라벨발행(I열) - 보정후 (0이면 선작업 반영 시 정상)")
                    st.dataframe(to_df(rows["s2_pre"], [C,D], extra=["급식사","선작업qty","보정후","보정후차이"]),
                                 use_container_width=True, hide_index=True)
            if rows["s2_box"]:
                s2box_err = [r for r in rows["s2_box"] if abs(r.get("차이") or 0) > 0.001]
                with st.expander(f"✅ BOX/EA환산 ({result['s2_box']}건) {'⚠️ 수량 확인 필요 ' + str(len(s2box_err)) + '건' if s2box_err else ''}",
                                 expanded=bool(s2box_err)):
                    if s2box_err:
                        st.warning("아래 항목은 차이가 있습니다. 환산 오류 여부를 확인해주세요.")
                    st.dataframe(to_df(rows["s2_box"], [C,D]), use_container_width=True, hide_index=True)
            if rows["s2_unit"]:
                with st.expander(f"✅ 단위다름(정상) ({result['s2_unit']}건)"):
                    st.caption("파일 간 단위가 달라 수량이 다르게 표시되지만 정상 처리된 품목입니다.")
                    st.dataframe(to_df(rows["s2_unit"], [C,D]), use_container_width=True, hide_index=True)

            st.divider()

            # ── 최종 요약 ──
            total_normal = result["s1_matched"] + result["s1_diff"] + result["s1_box"] + result["s1_unit"] + result["s2_pre"] + result["s2_box"] + result["s2_unit"]
            total_issue  = result["s1_wh"] + result["s1_lbl"] + result["s2_diff"] + result["s2_lbl"] + result["s2_cat"]
            qty_diff_val = result["s1_lbl_qty"] - result["s1_wh_qty"]

            st.markdown("### 📋 최종 결과 요약")
            if total_issue == 0:
                st.success("✅ 정상 — STEP 1, STEP 2 모두 확인이 필요한 항목이 없습니다.")
            else:
                st.error(f"🚨 확인 필요 — STEP 1, STEP 2 합계 {total_issue}건의 항목을 확인해야 합니다.")
            c1,c2,c3 = st.columns(3)
            c1.metric("✅ 정상 품목 수", total_normal,
                      help="일치 + 수량불일치 + BOX/EA환산 + 선작업")
            c2.metric("🚨 확인 필요 품목 수", total_issue,
                      help="이동처리만/라벨발행만/작업내역만/2단계 수량불일치")
            c3.metric("라벨발행 - 이동수량 차이", f"{qty_diff_val:+,.0f}")

            st.download_button(
                label="📥 결과 엑셀 다운로드",
                data=buf,
                file_name="창고이동_3단계검수결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

        except Exception as e:
            import traceback
            st.error(f"오류가 발생했습니다:\n\n```\n{traceback.format_exc()}\n```")
