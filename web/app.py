#!/usr/bin/env python3
"""
창고이동 3단계 검수 웹 프로그램
실행: python3 app.py  →  브라우저에서 http://localhost:5000 접속
"""

import os, re, uuid, threading
from pathlib import Path
from collections import defaultdict
from flask import Flask, render_template, request, jsonify, send_file, session

try:
    import xlrd, openpyxl
    from openpyxl.styles import Font as XFont, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    os.system("pip install xlrd openpyxl -q")
    import xlrd, openpyxl
    from openpyxl.styles import Font as XFont, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = "warehouse-inspection-2026"

UPLOAD_DIR = Path("uploads")
RESULT_DIR = Path("results")
UPLOAD_DIR.mkdir(exist_ok=True)
RESULT_DIR.mkdir(exist_ok=True)

# ── 검수 로직 ─────────────────────────────────────────────────────

BORDER_STYLE = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)

def clean_code(v):
    return re.sub(u'[​‌‍﻿ ]', '', str(v or '')).strip()

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

def load_label(path):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    s1_qty, s1_info, s2_qty, s2_info = defaultdict(float), {}, defaultdict(float), {}
    canceled = 0
    for i in range(2, ws.max_row + 1):
        state = ws.cell(i,16).value
        if state in ('취소','변경'):
            canceled += 1; continue
        제품명   = str(ws.cell(i,7).value  or '').strip()
        출고수량 = float(ws.cell(i,9).value or 0)
        단위     = str(ws.cell(i,10).value or '').strip()
        규격     = str(ws.cell(i,11).value or '').strip()
        급코드   = clean_code(ws.cell(i,12).value)
        erp코드  = str(ws.cell(i,13).value or '').strip()
        s1_qty[erp코드] += 출고수량
        if erp코드 not in s1_info:
            s1_info[erp코드] = {"ERP코드":erp코드,"급품목코드":급코드,"제품명":제품명,"규격":규격,"단위":단위}
        s2_qty[급코드] += 출고수량
        if 급코드 not in s2_info:
            s2_info[급코드] = {"급품목코드":급코드,"ERP코드":erp코드,"제품명":제품명,"규격":규격,"단위":단위}
    return s1_qty, s1_info, s2_qty, s2_info, canceled

def load_catering(path):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    qty_map, info_map = defaultdict(float), {}
    for i in range(4, ws.max_row + 1):
        d = ws.cell(i,3).value
        if not d or str(d)=='센터명': continue
        code = clean_code(ws.cell(i,7).value)
        qty  = float(ws.cell(i,11).value or 0)
        qty_map[code] += qty
        if code not in info_map:
            info_map[code] = {
                "제품코드":code,"제품명":str(ws.cell(i,8).value or '').strip(),
                "규격":str(ws.cell(i,9).value or '').strip(),
                "단위":str(ws.cell(i,10).value or '').strip(),
            }
    return {k:v/2 for k,v in qty_map.items()}, info_map

def load_prework(path):
    if not path or not os.path.exists(path): return {}
    wb = openpyxl.load_workbook(path)
    ws = wb['선작업-pre_qty'] if '선작업-pre_qty' in wb.sheetnames else wb.active
    items = {}
    for i in range(2, ws.max_row + 1):
        code = clean_code(ws.cell(i,2).value)
        if not code: continue
        if code not in items:
            items[code] = {"품목명":str(ws.cell(i,3).value or ''),"급식사":str(ws.cell(i,1).value or ''),"선작업qty":0.0,"보정후":0.0}
        items[code]["선작업qty"] += float(ws.cell(i,7).value or 0)
        items[code]["보정후"]    += float(ws.cell(i,8).value or 0)
    return items

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

def split_prework(matched, qty_diff, lbl_only, cat_only, prework_codes):
    def _split(rows):
        pre, normal = [], []
        for row in rows:
            if row["코드"] in prework_codes:
                r = dict(row)
                r["선작업qty"] = prework_codes[row["코드"]]["선작업qty"]
                r["보정후"]    = prework_codes[row["코드"]]["보정후"]
                r["급식사"]    = prework_codes[row["코드"]]["급식사"]
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
            center = key in ("코드","단위",col_a,col_b,"차이","이동날짜","선작업qty","보정후")
            cell_bg = bg
            if key=="차이" and isinstance(val,(int,float)) and abs(val)>0.001: cell_bg="FFD7D7"
            dat_style(cell, center=center, bg=cell_bg)
    for c,w in enumerate(w_base,1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(base_h))}1"


def run_inspection(wh_path, lbl_path, cat_path, pre_path, out_path):
    wh_qty, wh_info = load_warehouse(wh_path)
    ls1q,ls1i,ls2q,ls2i,canceled = load_label(lbl_path)
    cat_qty, cat_info = load_catering(cat_path)
    prework = load_prework(pre_path) if pre_path else {}

    s1m,s1d,s1w,s1l = compare(wh_qty,wh_info,ls1q,ls1i,"이동처리(F열)","라벨발행(I열)")
    s2m_all,s2d_all,s2l_all,s2c_all = compare(ls2q,ls2i,cat_qty,cat_info,"라벨발행(I열)","작업내역(K÷2)")

    if prework:
        s2pre,s2m,s2d,s2l,s2c = split_prework(s2m_all,s2d_all,s2l_all,s2c_all,prework)
    else:
        s2pre,s2m,s2d,s2l,s2c = [],s2m_all,s2d_all,s2l_all,s2c_all

    wb = openpyxl.Workbook()
    ws_sum = wb.active; ws_sum.title = "검수요약"
    ws_sum.column_dimensions["A"].width=46; ws_sum.column_dimensions["B"].width=16
    ws_sum.merge_cells("A1:B1")
    hdr_style(ws_sum.cell(1,1,"창고이동 3단계 검수 결과"), "1F4E79")
    s1c = len(s1m)+len(s1d); s2c_cnt = len(s2m)+len(s2d)
    summary_rows = [
        ("── STEP 1: 이동처리(F열) vs 라벨발행(I열) ──",""),
        ("  전체 항목 (ERP코드 기준)", len(s1m)+len(s1d)+len(s1w)+len(s1l)),
        ("  ✅ 일치", len(s1m)),("  ❌ 수량 불일치", len(s1d)),
        ("  ⚠️ 이동처리에만", len(s1w)),("  ⚠️ 라벨발행에만", len(s1l)),
        ("  일치율(공통기준)", f"{len(s1m)/s1c*100:.1f}%" if s1c else "-"),
        ("",""),
        ("── STEP 2: 라벨발행(L열) vs 작업내역(G열) ──",""),
        ("  전체 항목 (급품목코드 기준)", len(s2m)+len(s2d)+len(s2l)+len(s2c)+len(s2pre)),
        ("  ✅ 일치", len(s2m)),("  ❌ 수량 불일치", len(s2d)),
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
    write_xl_sheet(wb,"2단계_수량불일치",s2d,"C00000","FCE4D6",C,D,show_date=False)
    write_xl_sheet(wb,"2단계_라벨발행만",s2l,"843C0C","FFF2CC",C,D,show_date=False)
    write_xl_sheet(wb,"2단계_작업내역만",s2c,"7030A0","EAD1F5",C,D,show_date=False)
    write_xl_sheet(wb,"2단계_일치",      s2m,"375623","E2EFDA",C,D,show_date=False)
    if s2pre:
        write_xl_sheet(wb,"2단계_선작업(정상)",s2pre,"4472C4","DAE8FC",C,D,show_date=False,
                       extra_cols=[("선작업qty","선작업qty"),("보정후","보정후"),("급식사","급식사")])
    wb.save(out_path)

    return {
        "s1_matched":len(s1m),"s1_diff":len(s1d),"s1_wh":len(s1w),"s1_lbl":len(s1l),
        "s1_rate":f"{len(s1m)/s1c*100:.1f}" if s1c else "0",
        "s2_matched":len(s2m),"s2_diff":len(s2d),"s2_lbl":len(s2l),"s2_cat":len(s2c),
        "s2_pre":len(s2pre),
        "s2_rate":f"{len(s2m)/s2c_cnt*100:.1f}" if s2c_cnt else "0",
        "canceled": canceled,
        "wh_total": len(wh_qty),
    }


# ── Flask 라우트 ───────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/inspect", methods=["POST"])
def inspect():
    try:
        sid = str(uuid.uuid4())[:8]
        files_saved = {}

        for key in ["warehouse", "label", "catering", "prework"]:
            f = request.files.get(key)
            if f and f.filename:
                ext = Path(f.filename).suffix
                dest = UPLOAD_DIR / f"{sid}_{key}{ext}"
                f.save(dest)
                files_saved[key] = str(dest)

        if not all(k in files_saved for k in ["warehouse","label","catering"]):
            return jsonify({"error": "필수 파일(①②③)을 모두 업로드하세요."}), 400

        out_path = str(RESULT_DIR / f"{sid}_결과.xlsx")
        result = run_inspection(
            files_saved["warehouse"],
            files_saved["label"],
            files_saved["catering"],
            files_saved.get("prework"),
            out_path,
        )
        result["file_id"] = sid
        return jsonify(result)

    except Exception as e:
        import traceback
        return jsonify({"error": traceback.format_exc()}), 500

@app.route("/download/<sid>")
def download(sid):
    path = RESULT_DIR / f"{sid}_결과.xlsx"
    if not path.exists():
        return "파일을 찾을 수 없습니다.", 404
    return send_file(path, as_attachment=True,
                     download_name="창고이동_3단계검수결과.xlsx")

if __name__ == "__main__":
    print("=" * 50)
    print("창고이동 3단계 검수 웹 프로그램")
    print("브라우저에서 접속: http://localhost:5000")
    print("=" * 50)
    app.run(debug=False, port=5000)
