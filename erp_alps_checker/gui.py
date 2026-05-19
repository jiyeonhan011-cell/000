import sys
import os
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox, QTabWidget, QFrame, QLineEdit,
    QProgressBar, QGroupBox, QGridLayout, QSizePolicy
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QColor, QFont, QPalette, QIcon

sys.path.insert(0, os.path.dirname(__file__))
from logic import load_erp, load_alps, run_check, to_excel_bytes


# ── 색상 ──────────────────────────────────────────────────────────────────────
CLR_BG       = "#1e1e2e"
CLR_SURFACE  = "#2a2a3e"
CLR_ACCENT   = "#7c3aed"
CLR_ACCENT2  = "#a855f7"
CLR_TEXT     = "#e2e8f0"
CLR_SUBTEXT  = "#94a3b8"
CLR_MATCH    = "#166534"
CLR_MISMATCH = "#991b1b"
CLR_MIXED    = "#854d0e"
CLR_MATCH_BG    = "#dcfce7"
CLR_MISMATCH_BG = "#fee2e2"
CLR_MIXED_BG    = "#fef9c3"
CLR_BORDER   = "#374151"


def _styled_btn(text, color=CLR_ACCENT, height=42):
    btn = QPushButton(text)
    btn.setFixedHeight(height)
    btn.setCursor(Qt.PointingHandCursor)
    btn.setStyleSheet(f"""
        QPushButton {{
            background: {color};
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: bold;
            padding: 0 20px;
        }}
        QPushButton:hover {{ background: {CLR_ACCENT2}; }}
        QPushButton:pressed {{ background: #6d28d9; }}
        QPushButton:disabled {{ background: #4b5563; color: #9ca3af; }}
    """)
    return btn


def _card(title=""):
    box = QGroupBox(title)
    box.setStyleSheet(f"""
        QGroupBox {{
            background: {CLR_SURFACE};
            border: 1px solid {CLR_BORDER};
            border-radius: 10px;
            margin-top: 14px;
            font-size: 13px;
            font-weight: bold;
            color: {CLR_TEXT};
            padding: 10px;
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            left: 14px;
            padding: 0 6px;
            color: {CLR_ACCENT2};
        }}
    """)
    return box


# ── 백그라운드 스레드 ──────────────────────────────────────────────────────────
class WorkerThread(QThread):
    done = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, erp_path, alps_path, erp_sheet, alps_sheet):
        super().__init__()
        self.erp_path   = erp_path
        self.alps_path  = alps_path
        self.erp_sheet  = erp_sheet
        self.alps_sheet = alps_sheet

    def run(self):
        try:
            erp  = load_erp(self.erp_path,  sheet_name=self.erp_sheet  or None)
            alps = load_alps(self.alps_path, sheet_name=self.alps_sheet or None)
            self.done.emit(run_check(erp, alps))
        except Exception as e:
            self.error.emit(str(e))


# ── 메인 윈도우 ───────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ERP ↔ Alps 출고수량 검증")
        self.setMinimumSize(1200, 800)
        self.result = None
        self._build_ui()
        self._apply_theme()

    # ── UI 구성 ───────────────────────────────────────────────────────────────
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(20, 16, 20, 16)
        root.setSpacing(14)

        # 헤더
        hdr = QLabel("📦  ERP ↔ Alps  출고수량 검증")
        hdr.setStyleSheet(f"font-size: 22px; font-weight: bold; color: {CLR_TEXT};")
        root.addWidget(hdr)

        # 파일 선택 카드
        file_card = _card("파일 선택")
        fc_layout = QGridLayout(file_card)
        fc_layout.setSpacing(10)

        self.erp_edit  = self._file_row(fc_layout, 0, "ERP 파일", "ERP_D1")
        self.alps_edit = self._file_row(fc_layout, 1, "Alps 파일", "Alps_원본")

        # 출력 폴더
        fc_layout.addWidget(QLabel("결과 저장 폴더"), 2, 0)
        self.out_edit = QLineEdit()
        self.out_edit.setPlaceholderText("비워두면 ERP 파일과 같은 폴더에 저장")
        self.out_edit.setReadOnly(True)
        fc_layout.addWidget(self.out_edit, 2, 1)
        out_btn = QPushButton("폴더 선택")
        out_btn.setFixedWidth(100)
        out_btn.clicked.connect(self._pick_out_dir)
        fc_layout.addWidget(out_btn, 2, 2)

        root.addWidget(file_card)

        # 실행 버튼 + 프로그레스
        run_row = QHBoxLayout()
        self.run_btn = _styled_btn("🔍  검증 실행", height=48)
        self.run_btn.setFixedWidth(200)
        self.run_btn.clicked.connect(self._on_run)
        run_row.addWidget(self.run_btn)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setVisible(False)
        self.progress.setFixedHeight(48)
        self.progress.setStyleSheet(f"""
            QProgressBar {{ border: none; border-radius: 8px; background: {CLR_SURFACE}; }}
            QProgressBar::chunk {{ background: {CLR_ACCENT}; border-radius: 8px; }}
        """)
        run_row.addWidget(self.progress)
        root.addLayout(run_row)

        # 요약 카드 (숨김)
        self.summary_card = _card("검증 결과 요약")
        sc = QHBoxLayout(self.summary_card)
        self.m_erp   = self._metric(sc, "ERP 이동수량", "-")
        self._sep(sc)
        self.m_alps  = self._metric(sc, "Alps 출고수량", "-")
        self._sep(sc)
        self.m_match = self._metric(sc, "✓ 일치", "-", CLR_MATCH)
        self._sep(sc)
        self.m_miss  = self._metric(sc, "✗ 불일치", "-", CLR_MISMATCH)
        self._sep(sc)
        self.m_mix   = self._metric(sc, "⚠ 단위혼재", "-", CLR_MIXED)
        self.summary_card.setVisible(False)
        root.addWidget(self.summary_card)

        # 탭 테이블
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet(f"""
            QTabWidget::pane {{ border: 1px solid {CLR_BORDER}; border-radius: 8px; background: {CLR_SURFACE}; }}
            QTabBar::tab {{ background: {CLR_SURFACE}; color: {CLR_SUBTEXT}; padding: 8px 20px;
                            border-top-left-radius: 6px; border-top-right-radius: 6px; margin-right: 4px; }}
            QTabBar::tab:selected {{ background: {CLR_ACCENT}; color: white; font-weight: bold; }}
        """)
        self.tbl_all      = self._make_table()
        self.tbl_mismatch = self._make_table()
        self.tabs.addTab(self.tbl_all,      "전체 결과")
        self.tabs.addTab(self.tbl_mismatch, "불일치 목록")
        self.tabs.setVisible(False)
        root.addWidget(self.tabs, stretch=1)

        # 하단 버튼
        bot = QHBoxLayout()
        bot.addStretch()
        self.save_btn = _styled_btn("💾  Excel 저장", color="#059669")
        self.save_btn.setEnabled(False)
        self.save_btn.clicked.connect(self._on_save)
        bot.addWidget(self.save_btn)
        root.addLayout(bot)

        # 상태바
        self.statusBar().setStyleSheet(f"color:{CLR_SUBTEXT}; font-size:12px;")
        self.statusBar().showMessage("파일을 선택하고 검증 실행 버튼을 누르세요.")

    def _file_row(self, grid, row, label, sheet_default):
        grid.addWidget(QLabel(label), row, 0)
        edit = QLineEdit()
        edit.setReadOnly(True)
        edit.setPlaceholderText(f"{label} Excel 파일 경로")
        grid.addWidget(edit, row, 1)
        btn = QPushButton("파일 선택")
        btn.setFixedWidth(100)
        btn.clicked.connect(lambda _, e=edit: self._pick_file(e))
        grid.addWidget(btn, row, 2)
        sheet = QLineEdit(sheet_default)
        sheet.setFixedWidth(120)
        sheet.setPlaceholderText("시트명")
        grid.addWidget(sheet, row, 3)
        edit._sheet = sheet
        return edit

    def _metric(self, layout, title, value, color=None):
        frame = QFrame()
        vbox = QVBoxLayout(frame)
        vbox.setAlignment(Qt.AlignCenter)
        vbox.setSpacing(4)
        t = QLabel(title)
        t.setAlignment(Qt.AlignCenter)
        t.setStyleSheet(f"color:{CLR_SUBTEXT}; font-size:12px;")
        v = QLabel(value)
        v.setAlignment(Qt.AlignCenter)
        v.setStyleSheet(f"color:{color or CLR_TEXT}; font-size:22px; font-weight:bold;")
        vbox.addWidget(t)
        vbox.addWidget(v)
        layout.addWidget(frame, stretch=1)
        return v

    def _sep(self, layout):
        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setStyleSheet(f"color:{CLR_BORDER};")
        layout.addWidget(sep)

    def _make_table(self):
        t = QTableWidget()
        t.setEditTriggers(QTableWidget.NoEditTriggers)
        t.setSelectionBehavior(QTableWidget.SelectRows)
        t.setAlternatingRowColors(False)
        t.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        t.horizontalHeader().setStretchLastSection(True)
        t.verticalHeader().setVisible(False)
        t.setStyleSheet(f"""
            QTableWidget {{ background:{CLR_SURFACE}; color:{CLR_TEXT}; gridline-color:{CLR_BORDER};
                            font-size:13px; border:none; }}
            QHeaderView::section {{ background:{CLR_BG}; color:{CLR_ACCENT2}; font-weight:bold;
                                    padding:6px; border:1px solid {CLR_BORDER}; }}
            QTableWidget::item {{ padding: 4px 8px; }}
        """)
        return t

    def _apply_theme(self):
        self.setStyleSheet(f"""
            QMainWindow, QWidget {{ background:{CLR_BG}; color:{CLR_TEXT}; font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif; }}
            QLineEdit {{ background:{CLR_BG}; color:{CLR_TEXT}; border:1px solid {CLR_BORDER}; border-radius:6px; padding:6px 10px; font-size:13px; }}
            QPushButton {{ background:{CLR_SURFACE}; color:{CLR_TEXT}; border:1px solid {CLR_BORDER}; border-radius:6px; padding:6px 14px; font-size:13px; }}
            QPushButton:hover {{ background:{CLR_BORDER}; }}
            QLabel {{ color:{CLR_TEXT}; font-size:13px; }}
            QStatusBar {{ background:{CLR_SURFACE}; }}
        """)

    # ── 파일 선택 ─────────────────────────────────────────────────────────────
    def _pick_file(self, edit):
        path, _ = QFileDialog.getOpenFileName(self, "파일 선택", "", "Excel 파일 (*.xlsx *.xlsm *.xls)")
        if path:
            edit.setText(path)

    def _pick_out_dir(self):
        d = QFileDialog.getExistingDirectory(self, "결과 저장 폴더 선택")
        if d:
            self.out_edit.setText(d)

    # ── 검증 실행 ─────────────────────────────────────────────────────────────
    def _on_run(self):
        erp_path  = self.erp_edit.text().strip()
        alps_path = self.alps_edit.text().strip()

        if not erp_path or not alps_path:
            QMessageBox.warning(self, "파일 미선택", "ERP 파일과 Alps 파일을 모두 선택해 주세요.")
            return

        self.run_btn.setEnabled(False)
        self.progress.setVisible(True)
        self.statusBar().showMessage("검증 중...")

        self._worker = WorkerThread(
            erp_path, alps_path,
            self.erp_edit._sheet.text().strip(),
            self.alps_edit._sheet.text().strip(),
        )
        self._worker.done.connect(self._on_done)
        self._worker.error.connect(self._on_error)
        self._worker.start()

    def _on_done(self, result):
        self.result = result
        self.progress.setVisible(False)
        self.run_btn.setEnabled(True)

        s  = result["summary"]
        dm = result["d_map"]

        # 요약 업데이트
        self.m_erp.setText(f"{s['total_erp']:,.0f}")
        self.m_alps.setText(f"{s['total_alps']:,.0f}")
        self.m_match.setText(f"{s['match_count']}건\n{s['match_pct']:.1f}%")
        self.m_miss.setText(f"{s['mismatch_count']}건")
        self.m_mix.setText(f"{s['mixed_unit_count']}건")
        self.summary_card.setVisible(True)

        date_str = "  /  ".join(f"{v}={k}" for k, v in sorted(dm.items(), key=lambda x: x[1]))
        self.statusBar().showMessage(f"완료   |   날짜: {date_str}")

        # 테이블 채우기
        dl = {v: k for k, v in dm.items()}
        cols_base = ["품목코드", "품목명", "규격", "단위"]
        d_cols = [f"D1({dl.get('D1','')})", f"D2({dl.get('D2','')})", f"D3({dl.get('D3','')})"]
        cols_end = ["합산(D1+D2+D3)", "Alps출고수량", "차이", "일치"]
        headers = cols_base + d_cols + cols_end

        merged   = result["merged"]
        mismatch = result["mismatch"]

        self._fill_table(self.tbl_all,      merged,   headers, dl)
        self._fill_table(self.tbl_mismatch, mismatch, headers, dl)
        self.tabs.setTabText(0, f"전체 결과  ({len(merged)}건)")
        self.tabs.setTabText(1, f"불일치 목록  ({len(mismatch)}건)")
        self.tabs.setVisible(True)

        self.save_btn.setEnabled(True)

        # 자동 저장
        self._auto_save()

        # 불일치 알림 팝업
        if s["mismatch_count"] > 0:
            QTimer.singleShot(200, lambda: self._notify_mismatch(s))

    def _fill_table(self, tbl, df, headers, dl):
        tbl.clear()
        tbl.setColumnCount(len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        tbl.setRowCount(len(df))

        for r, (_, row) in enumerate(df.iterrows()):
            is_match = bool(row.get("일치여부", False))
            is_mixed = bool(row.get("단위혼재", False))

            if is_match:
                bg = QColor(CLR_MATCH_BG)
            elif is_mixed:
                bg = QColor(CLR_MIXED_BG)
            else:
                bg = QColor(CLR_MISMATCH_BG)

            d1k = f"D1({dl.get('D1','')})"
            d2k = f"D2({dl.get('D2','')})"
            d3k = f"D3({dl.get('D3','')})"

            vals = [
                str(row.get("품목코드", "") or ""),
                str(row.get("품목명", "") or ""),
                str(row.get("규격", "") or ""),
                str(row.get("단위", "") or ""),
                f"{row.get('D1', 0):,.0f}",
                f"{row.get('D2', 0):,.0f}",
                f"{row.get('D3', 0):,.0f}",
                f"{row.get('ERP합산', 0):,.0f}",
                f"{row.get('Alps출고수량', 0):,.0f}",
                f"{row.get('차이', 0):+,.0f}",
                "✓" if is_match else ("⚠" if is_mixed else "✗"),
            ]

            for c, val in enumerate(vals):
                item = QTableWidgetItem(val)
                item.setBackground(bg)
                item.setForeground(QColor("#1a1a1a"))
                item.setTextAlignment(Qt.AlignCenter if c >= 4 else Qt.AlignVCenter | Qt.AlignLeft)
                tbl.setItem(r, c, item)

    def _on_error(self, msg):
        self.progress.setVisible(False)
        self.run_btn.setEnabled(True)
        QMessageBox.critical(self, "오류", f"검증 실패:\n{msg}")
        self.statusBar().showMessage("오류 발생")

    # ── 저장 ─────────────────────────────────────────────────────────────────
    def _auto_save(self):
        if not self.result:
            return
        path = self._resolve_save_path()
        try:
            with open(path, "wb") as f:
                f.write(to_excel_bytes(self.result))
            self.statusBar().showMessage(f"결과 자동 저장 완료 → {path}")
            self._last_saved = path
        except Exception as e:
            self.statusBar().showMessage(f"자동 저장 실패: {e}")

    def _on_save(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "저장 위치 선택", self._resolve_save_path(),
            "Excel 파일 (*.xlsx)"
        )
        if not path:
            return
        try:
            with open(path, "wb") as f:
                f.write(to_excel_bytes(self.result))
            QMessageBox.information(self, "저장 완료", f"저장되었습니다:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "저장 실패", str(e))

    def _resolve_save_path(self):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = f"ERP_Alps_검증결과_{ts}.xlsx"
        out_dir = self.out_edit.text().strip()
        if not out_dir:
            erp_path = self.erp_edit.text().strip()
            out_dir = os.path.dirname(erp_path) if erp_path else os.path.expanduser("~")
        return os.path.join(out_dir, fname)

    # ── 불일치 알림 ──────────────────────────────────────────────────────────
    def _notify_mismatch(self, s):
        msg = QMessageBox(self)
        msg.setWindowTitle("⚠ 불일치 항목 발견")
        msg.setIcon(QMessageBox.Warning)
        msg.setText(
            f"<b>{s['mismatch_count']}건의 불일치 항목이 있습니다.</b><br><br>"
            f"ERP 기준 불일치 수량: <b>{s['mismatch_qty']:,.0f}</b><br>"
            f"일치율: <b>{s['match_pct']:.1f}%</b>"
        )
        msg.exec_()


# ── 진입점 ────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())
