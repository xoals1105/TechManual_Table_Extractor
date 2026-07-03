"""비개발자용 PyQt6 GUI.

기능:
  - 한글 파일(.hwp/.hwpx) 또는 폴더 선택
  - 모드 선택: 모든 표 추출(--all) / 규칙에 맞는 표만
  - 규칙 편집기: 표 형태로 규칙을 편집하고 config/target_tables.yaml 에 저장
  - 실행 로그 실시간 표시, 완료 후 결과 폴더 열기

실행: run_gui.bat  또는  .venv\\Scripts\\python -m src.gui
"""
from __future__ import annotations

import logging
import os
import sys
from pathlib import Path

import yaml
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QRadioButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .main import collect_input_files, load_config, process_file

ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "config" / "target_tables.yaml"
OUTPUT_DIR = ROOT / "output"

CONFIG_HEADER = """\
# 이 파일은 GUI(규칙 저장)에서 자동 생성되었습니다.
# 점수: 헤더 일치율 x50, 제목 키워드 +30, 열 개수 +10 → threshold 이상이면 추출
"""


class GuiLogHandler(logging.Handler):
    """logger 출력 → GUI 로그 창으로 전달."""

    def __init__(self, signal):
        super().__init__()
        self._signal = signal
        self.setFormatter(logging.Formatter("%(asctime)s %(message)s", "%H:%M:%S"))

    def emit(self, record):
        try:
            self._signal.emit(self.format(record))
        except Exception:
            pass


class ExtractWorker(QThread):
    """백그라운드 추출 작업 (COM 사용 가능하도록 스레드 내 COM 초기화)."""

    log = pyqtSignal(str)
    finished_ok = pyqtSignal(str)   # 요약 메시지
    failed = pyqtSignal(str)        # 오류 메시지

    def __init__(self, input_path: Path, rules: list[dict], extract_all: bool,
                 include_title: bool):
        super().__init__()
        self.input_path = input_path
        self.rules = rules
        self.extract_all = extract_all
        self.include_title = include_title

    def run(self):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pythoncom = None

        logger = logging.getLogger("extractor")
        logger.setLevel(logging.INFO)
        gui_handler = GuiLogHandler(self.log)
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(OUTPUT_DIR / "extract.log", encoding="utf-8")
        file_handler.setFormatter(
            logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        logger.addHandler(gui_handler)
        logger.addHandler(file_handler)

        try:
            files = collect_input_files(self.input_path)
            self.log.emit(f"입력 파일 {len(files)}개, 규칙 {len(self.rules)}개"
                          f" (표 제목 {'포함' if self.include_title else '미포함'})")
            summary, failed = [], []
            for f in files:
                try:
                    summary.append(process_file(
                        f, self.rules, OUTPUT_DIR, self.extract_all,
                        include_title=self.include_title,
                    ))
                except Exception:
                    logger.exception(f"처리 실패: {f}")
                    failed.append(f.name)

            lines = [f"{s['file']}: 표 {s['total']}개 중 {s['extracted']}개 추출"
                     for s in summary]
            if failed:
                lines.append(f"실패: {', '.join(failed)}")
                self.failed.emit("\n".join(lines))
            else:
                self.finished_ok.emit("\n".join(lines) or "처리된 파일이 없습니다.")
        except Exception as e:
            self.failed.emit(str(e))
        finally:
            logger.removeHandler(gui_handler)
            logger.removeHandler(file_handler)
            file_handler.close()
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


RULE_COLUMNS = ["규칙 이름", "표 1행 헤더 (쉼표로 구분)", "제목 키워드 (쉼표, 선택)",
                "열 개수 (선택)", "기준 점수"]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("기술교범 표 추출기")
        self.resize(920, 720)
        self.worker: ExtractWorker | None = None

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # --- 1. 입력 선택 ---
        input_box = QGroupBox("1. 한글 파일 선택")
        input_layout = QHBoxLayout(input_box)
        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("추출할 .hwp/.hwpx 파일 또는 폴더를 선택하세요")
        btn_file = QPushButton("파일 선택...")
        btn_file.clicked.connect(self.pick_file)
        btn_dir = QPushButton("폴더 선택...")
        btn_dir.clicked.connect(self.pick_dir)
        input_layout.addWidget(self.input_edit, 1)
        input_layout.addWidget(btn_file)
        input_layout.addWidget(btn_dir)
        layout.addWidget(input_box)

        # --- 2. 추출 방식 ---
        mode_box = QGroupBox("2. 추출 방식")
        mode_layout = QHBoxLayout(mode_box)
        self.radio_all = QRadioButton("모든 표 추출 (문서 구조 확인용)")
        self.radio_rules = QRadioButton("아래 규칙에 맞는 표만 추출")
        self.radio_all.setChecked(True)
        self.chk_title = QCheckBox("엑셀 1행에 표 제목 넣기")
        mode_layout.addWidget(self.radio_all)
        mode_layout.addWidget(self.radio_rules)
        mode_layout.addStretch(1)
        mode_layout.addWidget(self.chk_title)
        layout.addWidget(mode_box)

        # --- 3. 규칙 편집기 ---
        rules_box = QGroupBox("3. 표 선별 규칙 (규칙에 맞는 표만 추출할 때 사용)")
        rules_layout = QVBoxLayout(rules_box)
        self.table = QTableWidget(0, len(RULE_COLUMNS))
        self.table.setHorizontalHeaderLabels(RULE_COLUMNS)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        rules_layout.addWidget(self.table)

        rule_btns = QHBoxLayout()
        btn_add = QPushButton("규칙 추가")
        btn_add.clicked.connect(self.add_rule_row)
        btn_del = QPushButton("선택 규칙 삭제")
        btn_del.clicked.connect(self.delete_rule_row)
        btn_save = QPushButton("규칙 저장")
        btn_save.clicked.connect(self.save_rules)
        rule_btns.addWidget(btn_add)
        rule_btns.addWidget(btn_del)
        rule_btns.addStretch(1)
        rule_btns.addWidget(btn_save)
        rules_layout.addLayout(rule_btns)
        hint = QLabel('예) 표 1행이 "품명 | 규격 | 수량" 이면 헤더 칸에  품명, 규격, 수량')
        hint.setStyleSheet("color: gray;")
        rules_layout.addWidget(hint)
        layout.addWidget(rules_box, 1)

        # --- 4. 실행 ---
        run_layout = QHBoxLayout()
        self.btn_run = QPushButton("추출 실행")
        self.btn_run.setMinimumHeight(36)
        self.btn_run.clicked.connect(self.run_extract)
        self.btn_open = QPushButton("결과 폴더 열기")
        self.btn_open.clicked.connect(self.open_output)
        run_layout.addWidget(self.btn_run, 1)
        run_layout.addWidget(self.btn_open)
        layout.addLayout(run_layout)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setPlaceholderText("실행하면 진행 상황이 여기에 표시됩니다.")
        layout.addWidget(self.log_view, 1)

        self.load_rules_into_table()

    # ---------------- 입력 선택 ----------------
    def pick_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "한글 파일 선택", str(ROOT),
            "한글 문서 (*.hwp *.hwpx);;모든 파일 (*)")
        if path:
            self.input_edit.setText(path)

    def pick_dir(self):
        path = QFileDialog.getExistingDirectory(self, "폴더 선택", str(ROOT))
        if path:
            self.input_edit.setText(path)

    # ---------------- 규칙 편집 ----------------
    def load_rules_into_table(self):
        try:
            rules, output_opts = load_config(CONFIG_PATH, require_rules=False)
        except FileNotFoundError:
            rules, output_opts = [], {}
        self.chk_title.setChecked(bool(output_opts.get("include_title", False)))
        self.table.setRowCount(0)
        for rule in rules:
            self.add_rule_row(rule)

    def add_rule_row(self, rule: dict | None = None):
        rule = rule if isinstance(rule, dict) else {}
        row = self.table.rowCount()
        self.table.insertRow(row)
        cols = rule.get("expected_cols")
        values = [
            rule.get("name", ""),
            ", ".join(rule.get("expected_headers") or []),
            ", ".join(rule.get("title_keywords") or []),
            "" if cols in (None, "") else str(cols),
            str(rule.get("threshold", 60)),
        ]
        for col, value in enumerate(values):
            self.table.setItem(row, col, QTableWidgetItem(value))

    def delete_rule_row(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True)
        if not rows:
            QMessageBox.information(self, "안내", "삭제할 규칙 행을 먼저 클릭하세요.")
            return
        for r in rows:
            self.table.removeRow(r)

    def _cell(self, row: int, col: int) -> str:
        item = self.table.item(row, col)
        return item.text().strip() if item else ""

    def rules_from_table(self) -> list[dict]:
        rules = []
        for row in range(self.table.rowCount()):
            headers = [h.strip() for h in self._cell(row, 1).split(",") if h.strip()]
            keywords = [k.strip() for k in self._cell(row, 2).split(",") if k.strip()]
            if not headers and not keywords:
                continue  # 빈 행 무시
            cols_text = self._cell(row, 3)
            try:
                threshold = float(self._cell(row, 4) or 60)
            except ValueError:
                threshold = 60
            rules.append({
                "name": self._cell(row, 0) or f"규칙{row + 1}",
                "expected_headers": headers,
                "title_keywords": keywords,
                "expected_cols": int(cols_text) if cols_text.isdigit() else None,
                "threshold": threshold,
            })
        return rules

    def save_rules(self):
        config = {
            "output": {"include_title": self.chk_title.isChecked()},
            "rules": self.rules_from_table(),
        }
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            f.write(CONFIG_HEADER)
            yaml.safe_dump(config, f, allow_unicode=True, sort_keys=False,
                           default_flow_style=False)
        QMessageBox.information(self, "저장 완료",
                                f"규칙 {len(config['rules'])}개를 저장했습니다.\n{CONFIG_PATH}")

    # ---------------- 실행 ----------------
    def run_extract(self):
        input_text = self.input_edit.text().strip().strip('"')
        if not input_text:
            QMessageBox.warning(self, "입력 필요", "한글 파일 또는 폴더를 먼저 선택하세요.")
            return
        input_path = Path(input_text)
        if not input_path.exists():
            QMessageBox.warning(self, "경로 오류", f"경로를 찾을 수 없습니다:\n{input_path}")
            return

        extract_all = self.radio_all.isChecked()
        rules = [] if extract_all else self.rules_from_table()
        if not extract_all and not rules:
            QMessageBox.warning(self, "규칙 필요",
                                "규칙이 없습니다. 규칙을 추가하거나 '모든 표 추출'을 선택하세요.")
            return

        self.log_view.clear()
        self.btn_run.setEnabled(False)
        self.btn_run.setText("실행 중... (큰 문서는 몇 분 걸릴 수 있습니다)")

        self.worker = ExtractWorker(input_path, rules, extract_all,
                                    self.chk_title.isChecked())
        self.worker.log.connect(self.log_view.appendPlainText)
        self.worker.finished_ok.connect(self.on_done)
        self.worker.failed.connect(self.on_failed)
        self.worker.start()

    def _reset_run_button(self):
        self.btn_run.setEnabled(True)
        self.btn_run.setText("추출 실행")

    def on_done(self, message: str):
        self._reset_run_button()
        QMessageBox.information(self, "완료", message + f"\n\n결과 위치: {OUTPUT_DIR}")

    def on_failed(self, message: str):
        self._reset_run_button()
        QMessageBox.warning(self, "일부 실패 또는 오류", message)

    def open_output(self):
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        os.startfile(str(OUTPUT_DIR))  # noqa: S606 (Windows 전용)


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
