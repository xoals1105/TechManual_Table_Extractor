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
from PyQt6.QtCore import QPoint, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QMouseEvent
from PyQt6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
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
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from .main import collect_input_files, load_config, process_file

ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "config" / "target_tables.yaml"
OUTPUT_DIR = ROOT / "output"

HELP_HTML = """
<h2>사용 순서</h2>
<ol>
<li><b>한글 파일 선택</b> — 추출할 .hwp / .hwpx 파일 또는 폴더를 고릅니다.</li>
<li><b>모든 표 추출</b> — 처음에는 이 모드를 선택하고 [추출 실행]을 누릅니다.
    [결과 폴더 열기]로 엑셀을 열어, 뽑고 싶은 표의 <b>맨 윗줄(1행)</b>을 확인합니다.</li>
<li><b>규칙 추가</b> — 아래 규칙 표에서 [규칙 추가]를 누르고
    「표 1행 헤더」에 쉼표로 입력합니다. (예: <code>품명, 규격, 수량</code>)
    입력 후 [규칙 저장]을 누릅니다.</li>
<li><b>규칙에 맞는 표만 추출</b> — 모드를 바꾼 뒤 [추출 실행]합니다.</li>
<li><b>결과 확인</b> — <code>output</code> 폴더에 <code>&lt;파일명&gt;_tables.xlsx</code> 가 생성됩니다.
    표 1개 = 엑셀 시트 1개입니다.</li>
</ol>

<h2>규칙 항목 설명</h2>
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">
<tr><th>항목</th><th>의미</th><th>어디서 확인</th><th>필수</th></tr>
<tr>
  <td><b>규칙 이름</b></td>
  <td>엑셀 시트 이름에 붙는 이름입니다. (예: <code>부품목록표_0</code>)</td>
  <td>사용자가 알아보기 쉬운 이름을 짓습니다.</td>
  <td>선택</td>
</tr>
<tr>
  <td><b>표 1행 헤더</b></td>
  <td>표 <b>맨 윗줄</b>에 있는 열 제목입니다. 선별의 핵심 조건입니다.</td>
  <td>「모든 표 추출」 결과 엑셀의 1행, 또는 한글 문서에서 표 첫 줄</td>
  <td><b>필수</b></td>
</tr>
<tr>
  <td><b>제목 키워드</b></td>
  <td>표 <b>바로 위</b> 문장·캡션에 나오는 단어입니다. 비슷한 표를 구분할 때 씁니다.</td>
  <td>한글에서 표 위 문단 (예: "부품 목록", "표 3-1")</td>
  <td>선택</td>
</tr>
<tr>
  <td><b>열 개수</b></td>
  <td>표가 몇 열인지 나타내는 숫자입니다.</td>
  <td>엑셀 1행의 열 개수 (예: 4)</td>
  <td>선택</td>
</tr>
<tr>
  <td><b>일치율 기준(%)</b></td>
  <td>표와 규칙의 <b>일치율</b>이 이 값 <b>이상</b>이면 해당 규칙으로 표를 뽑습니다.</td>
  <td>기본값 80. 조정 방법은 아래 참고</td>
  <td>기본 80</td>
</tr>
</table>

<p><b>참고</b></p>
<ul>
<li>표 <b>2행 이하 내용</b>(데이터)은 규칙에 넣지 않습니다. <b>1행 헤더만</b> 넣으세요.</li>
<li>규칙을 문서 안 표 순서대로 적을 필요는 없습니다.</li>
<li>헤더 입력 시 공백은 무시됩니다. (<code>품 명</code> = <code>품명</code>)</li>
<li>DRM이 걸린 .hwp 파일은 <b>한글 프로그램이 설치된 PC</b>에서만 처리됩니다.</li>
</ul>

<h2>일치율 기준(%) 사용법</h2>
<p>규칙에 <b>입력한 조건만</b>으로 만점을 계산하고, 표가 얼마나 맞는지를
<b>일치율 0~100%</b> 로 환산합니다. 일치율이 기준(%) 이상이면 표가 선정됩니다.</p>
<p><b>입력한 조건을 전부 만족하면 어떤 조합이든 100%</b> 입니다.
헤더만 입력해도, 헤더+키워드를 입력해도 전부 일치하면 100%가 됩니다.</p>

<h3>조건별 비중</h3>
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">
<tr><th>조건</th><th>비중</th><th>설명</th></tr>
<tr><td>헤더 일치</td><td>50</td>
    <td>입력한 헤더 중 표 1행과 일치하는 비율만큼 반영<br>
    예) 4개 중 3개 일치 → 비중의 75%만 인정</td></tr>
<tr><td>제목 키워드</td><td>30</td><td>키워드 중 하나라도 표 위·캡션에 있으면</td></tr>
<tr><td>열 개수</td><td>10</td><td>입력한 열 개수와 표 열 수가 같으면</td></tr>
</table>
<p>입력하지 않은 조건은 만점 계산에서 <b>제외</b>됩니다.</p>

<h3>일치율 예시 (기준 80%일 때)</h3>
<ul>
<li>헤더만 입력, 4/4 전부 일치 → <b>100%</b> → 선정</li>
<li>헤더만 입력, 4개 중 3개 일치 → <b>75%</b> → 미달</li>
<li>헤더+키워드 입력, 헤더 4/4 일치·키워드 불일치 → 50/80 = <b>63%</b> → 미달</li>
<li>헤더+키워드 입력, 둘 다 일치 → <b>100%</b> → 선정</li>
</ul>

<h3>기준(%) 조정</h3>
<ul>
<li><b>엉뚱한 표까지 뽑힐 때</b> → 90 ~ 100 으로 올립니다.</li>
<li><b>원하는 표가 빠질 때</b> → 60 ~ 70 으로 내리거나, 헤더·제목 키워드를 보강합니다.</li>
<li>기본값 <b>80</b>이면 "입력한 조건이 대부분 맞는 표"만 뽑힙니다.</li>
</ul>
<p>실행 후 <code>output\\extract.log</code> 에 표마다 일치율과 선정 이유가 기록됩니다.
결과가 이상하면 이 로그를 확인하세요.</p>
"""

# (열 헤더 라벨, 셀 툴팁)
RULE_COLUMN_SPECS: list[tuple[str, str]] = [
    (
        "규칙 이름 (선택)",
        "엑셀 시트 이름에 쓰입니다. 비워 두면 '규칙1', '규칙2' … 로 자동 지정됩니다.",
    ),
    (
        "표 1행 헤더 (필수)",
        "엑셀·한글에서 본 표 맨 윗줄을 쉼표로 입력하세요.\n"
        "예: 표 1행이 '품명 | 규격 | 수량' 이면 → 품명, 규격, 수량\n"
        "공백은 무시됩니다. ('품 명' = '품명')",
    ),
    (
        "제목 키워드 (선택)",
        "표 바로 위 문장·캡션에 있는 단어를 쉼표로 입력합니다.\n"
        "헤더만으로 표를 구분할 수 있으면 비워도 됩니다.\n"
        "예: 부품 목록, 부품목록",
    ),
    (
        "열 개수 (선택)",
        "표가 몇 열인지 숫자로 입력합니다. (예: 4)\n"
        "헤더가 비슷한 표가 여러 개일 때 보조로 사용합니다. 비워도 됩니다.",
    ),
    (
        "일치율 기준(%)",
        "표와 규칙의 일치율(0~100%)이 이 값 이상이면 표를 뽑습니다. 기본값 80.\n"
        "입력한 조건(헤더·키워드·열 개수)을 전부 만족하면 100%입니다.\n"
        "엉뚱한 표가 뽑히면 90~100으로 올리고, 빠지면 60~70으로 내립니다.",
    ),
]

CONFIG_HEADER = """\
# 이 파일은 GUI(규칙 저장)에서 자동 생성되었습니다.
# threshold: 일치율(%) 기준. 입력한 조건(헤더 50, 키워드 30, 열 개수 10 비중)을
# 전부 만족하면 100%이며, 일치율이 threshold 이상인 표만 추출됩니다.
"""


class _TitleBarButton(QPushButton):
    """제목 표시줄용 작은 버튼 (최소화/최대화/닫기/?)."""

    def __init__(self, text: str, *, close: bool = False, narrow: bool = False):
        super().__init__(text)
        w = 36 if narrow else 46
        self.setFixedSize(w, 32)
        self.setFlat(True)
        if close:
            self.setStyleSheet(
                "QPushButton { border: none; background: transparent; }"
                "QPushButton:hover { background: #e81123; color: white; }"
            )
        else:
            self.setStyleSheet(
                "QPushButton { border: none; background: transparent; }"
                "QPushButton:hover { background: #e5e5e5; }"
            )


class TitleBar(QWidget):
    """Windows 스타일 커스텀 제목 표시줄 — ? 는 최소화 버튼 왼쪽."""

    def __init__(self, window: QMainWindow, *, on_help):
        super().__init__(window)
        self._window = window
        self._drag_pos: QPoint | None = None
        self.setFixedHeight(32)
        self.setStyleSheet("background: #f0f0f0;")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 0, 0, 0)
        layout.setSpacing(0)

        self._title = QLabel(window.windowTitle())
        self._title.setStyleSheet("background: transparent;")
        layout.addWidget(self._title)
        layout.addStretch(1)

        btn_help = _TitleBarButton("?", narrow=True)
        btn_help.setToolTip("사용 방법 — 프로그램 사용법, 규칙 항목, 기준 점수 조정")
        btn_help.clicked.connect(on_help)
        layout.addWidget(btn_help)

        btn_min = _TitleBarButton("─")
        btn_min.setToolTip("최소화")
        btn_min.clicked.connect(window.showMinimized)
        layout.addWidget(btn_min)

        self._btn_max = _TitleBarButton("□")
        self._btn_max.setToolTip("최대화")
        self._btn_max.clicked.connect(self._toggle_maximize)
        layout.addWidget(self._btn_max)

        btn_close = _TitleBarButton("✕", close=True)
        btn_close.setToolTip("닫기")
        btn_close.clicked.connect(window.close)
        layout.addWidget(btn_close)

    def _toggle_maximize(self) -> None:
        if self._window.isMaximized():
            self._window.showNormal()
            self._btn_max.setText("□")
        else:
            self._window.showMaximized()
            self._btn_max.setText("❐")

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            target = self.childAt(event.position().toPoint())
            if isinstance(target, QPushButton):
                return super().mousePressEvent(event)
            self._drag_pos = event.globalPosition().toPoint() - self._window.pos()
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QMouseEvent) -> None:
        if self._drag_pos is not None and event.buttons() & Qt.MouseButton.LeftButton:
            self._window.move(event.globalPosition().toPoint() - self._drag_pos)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:
        self._drag_pos = None
        super().mouseReleaseEvent(event)

    def mouseDoubleClickEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self._toggle_maximize()
            event.accept()
            return
        super().mouseDoubleClickEvent(event)


class HelpDialog(QDialog):
    """사용 방법 상세 안내 창."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("사용 방법")
        self.resize(620, 520)

        layout = QVBoxLayout(self)
        browser = QTextBrowser()
        browser.setHtml(HELP_HTML)
        browser.setOpenExternalLinks(False)
        layout.addWidget(browser)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        buttons.accepted.connect(self.accept)
        close_btn = buttons.button(QDialogButtonBox.StandardButton.Close)
        if close_btn:
            close_btn.setText("닫기")
        layout.addWidget(buttons)


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
                 include_title: bool, include_footer: bool, visual_merge: bool,
                 sheet_layout: str):
        super().__init__()
        self.input_path = input_path
        self.rules = rules
        self.extract_all = extract_all
        self.include_title = include_title
        self.include_footer = include_footer
        self.visual_merge = visual_merge
        self.sheet_layout = sheet_layout

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
            layout_label = (
                "표당 시트 1개" if self.sheet_layout == "per_table" else "단일 시트 연속 배치"
            )
            self.log.emit(
                f"입력 파일 {len(files)}개, 규칙 {len(self.rules)}개"
                f" (표 제목: {'포함' if self.include_title else '미포함'},"
                f" 꼬리말: {'포함' if self.include_footer else '미포함'},"
                f" 숨은 가로선 병합: {'적용' if self.visual_merge else '미적용'},"
                f" 시트 배치: {layout_label})"
            )
            summary, failed = [], []
            for f in files:
                try:
                    summary.append(process_file(
                        f, self.rules, OUTPUT_DIR, self.extract_all,
                        include_title=self.include_title,
                        include_footer=self.include_footer,
                        visual_merge=self.visual_merge,
                        sheet_layout=self.sheet_layout,
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



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("기술교범 표 추출기")
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.resize(920, 720)
        self.worker: ExtractWorker | None = None

        shell = QWidget()
        shell.setObjectName("appShell")
        shell.setStyleSheet("#appShell { border: 1px solid #ababab; }")
        shell_layout = QVBoxLayout(shell)
        shell_layout.setContentsMargins(0, 0, 0, 0)
        shell_layout.setSpacing(0)

        self._title_bar = TitleBar(self, on_help=self.show_help)
        shell_layout.addWidget(self._title_bar)

        central = QWidget()
        shell_layout.addWidget(central, 1)
        self.setCentralWidget(shell)

        layout = QVBoxLayout(central)
        layout.setContentsMargins(8, 8, 8, 8)

        # --- 1. 입력 선택 ---
        input_box = QGroupBox("1. 한글 파일 선택")
        input_layout = QHBoxLayout(input_box)
        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("추출할 .hwp/.hwpx 파일 또는 폴더를 선택하세요")
        self.input_edit.setToolTip("처리할 한글 파일 하나, 또는 .hwp/.hwpx 가 들어 있는 폴더를 선택하세요.")
        btn_file = QPushButton("파일 선택...")
        btn_file.setToolTip("한글 파일(.hwp, .hwpx) 하나를 선택합니다.")
        btn_file.clicked.connect(self.pick_file)
        btn_dir = QPushButton("폴더 선택...")
        btn_dir.setToolTip("폴더 안의 모든 .hwp / .hwpx 파일을 한꺼번에 처리합니다.")
        btn_dir.clicked.connect(self.pick_dir)
        input_layout.addWidget(self.input_edit, 1)
        input_layout.addWidget(btn_file)
        input_layout.addWidget(btn_dir)
        layout.addWidget(input_box)

        # --- 2. 추출 방식 ---
        mode_box = QGroupBox("2. 추출 방식")
        mode_layout = QHBoxLayout(mode_box)
        self.radio_all = QRadioButton("모든 표 추출 (문서 구조 확인용)")
        self.radio_all.setToolTip(
            "문서 안의 모든 표를 엑셀로 뽑습니다.\n"
            "처음에 이 모드로 실행해 표 1행 헤더를 확인한 뒤 규칙을 만드세요.")
        self.radio_rules = QRadioButton("아래 규칙에 맞는 표만 추출")
        self.radio_rules.setToolTip("아래 규칙 표에 입력한 조건에 맞는 표만 엑셀로 뽑습니다.")
        self.radio_all.setChecked(True)
        mode_layout.addWidget(self.radio_all)
        mode_layout.addWidget(self.radio_rules)
        mode_layout.addStretch(1)
        layout.addWidget(mode_box)

        # --- 3. 엑셀 출력 옵션 ---
        output_box = QGroupBox("3. 엑셀 출력 옵션")
        output_box.setToolTip("이 옵션들은 [규칙 저장]을 누르면 설정 파일에 함께 저장됩니다.")
        output_layout = QHBoxLayout(output_box)

        sheet_label = QLabel("시트:")
        self.radio_sheet_per_table = QRadioButton("표 하나당 시트 1개")
        self.radio_sheet_per_table.setToolTip(
            "추출된 표마다 엑셀 시트를 하나씩 만듭니다. (기본)")
        self.radio_sheet_single = QRadioButton("한 시트에 연속 배치")
        self.radio_sheet_single.setToolTip(
            "추출된 모든 표를 '추출표' 시트 하나에 위에서 아래로 이어 붙입니다.\n"
            "표 사이에는 빈 행 2칸을 둡니다. 꼬리말이 있으면 표 바로 아래에 넣은 뒤 빈 행 2칸을 둡니다.")
        self.radio_sheet_per_table.setChecked(True)
        # 시트 라디오를 별도 그룹으로 묶어 추출 방식 라디오와 간섭 방지
        sheet_group = QButtonGroup(self)
        sheet_group.addButton(self.radio_sheet_per_table)
        sheet_group.addButton(self.radio_sheet_single)
        self._sheet_group = sheet_group

        self.chk_title = QCheckBox("표 제목 넣기")
        self.chk_title.setToolTip(
            "체크: 엑셀 1행에 표 제목, 3행부터 표 내용\n"
            "해제: 1행부터 표 내용만 (기본 권장)")
        self.chk_footer = QCheckBox("표 아래에 꼬리말 넣기")
        self.chk_footer.setToolTip(
            "표가 있는 쪽의 꼬리말을 엑셀 표 바로 아래 행에 넣습니다.\n"
            "줄바꿈·순서는 원문 그대로이며, .hwpx는 문서 XML에서 읽습니다.")
        self.chk_visual_merge = QCheckBox("숨은 가로선 셀 합치기")
        self.chk_visual_merge.setToolTip(
            "한글에서 가로 구분선이 없어 한 칸처럼 보이는 셀을,\n"
            "엑셀에서도 세로 병합으로 맞춥니다.")

        output_layout.addWidget(sheet_label)
        output_layout.addWidget(self.radio_sheet_per_table)
        output_layout.addWidget(self.radio_sheet_single)
        output_layout.addSpacing(24)
        output_layout.addWidget(self.chk_title)
        output_layout.addWidget(self.chk_footer)
        output_layout.addWidget(self.chk_visual_merge)
        output_layout.addStretch(1)
        layout.addWidget(output_box)

        # --- 4. 규칙 편집기 ---
        rules_box = QGroupBox("4. 표 선별 규칙 (규칙에 맞는 표만 추출할 때 사용)")
        rules_box.setToolTip("각 열 헤더에 마우스를 올리면 입력 방법을 볼 수 있습니다.")
        rules_layout = QVBoxLayout(rules_box)
        self.table = QTableWidget(0, len(RULE_COLUMN_SPECS))
        for col, (label, tip) in enumerate(RULE_COLUMN_SPECS):
            header_item = QTableWidgetItem(label)
            header_item.setToolTip(tip)
            self.table.setHorizontalHeaderItem(col, header_item)
        header = self.table.horizontalHeader()
        header.setToolTip("열 이름에 마우스를 올리면 각 항목 설명이 표시됩니다.")
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        rules_layout.addWidget(self.table)

        rule_btns = QHBoxLayout()
        btn_add = QPushButton("규칙 추가")
        btn_add.setToolTip("새 규칙 행을 추가합니다. 표 1행 헤더(필수)를 입력하세요.")
        btn_add.clicked.connect(self.add_rule_row)
        btn_del = QPushButton("선택 규칙 삭제")
        btn_del.setToolTip("표에서 선택한 규칙 행을 삭제합니다.")
        btn_del.clicked.connect(self.delete_rule_row)
        btn_save = QPushButton("규칙 저장")
        btn_save.setToolTip(
            "입력한 규칙과 '엑셀 출력 옵션'을 config/target_tables.yaml 에 저장합니다.")
        btn_save.clicked.connect(self.save_rules)
        rule_btns.addWidget(btn_add)
        rule_btns.addWidget(btn_del)
        rule_btns.addStretch(1)
        rule_btns.addWidget(btn_save)
        rules_layout.addLayout(rule_btns)
        hint = QLabel(
            '규칙이 헷갈리면 제목 표시줄 오른쪽 [?] (최소화 왼쪽)를 누르세요. '
            '각 칸에 마우스를 올려도 간단한 설명이 나옵니다.'
        )
        hint.setStyleSheet("color: gray;")
        hint.setWordWrap(True)
        rules_layout.addWidget(hint)
        layout.addWidget(rules_box, 1)

        # --- 5. 실행 ---
        run_layout = QHBoxLayout()
        self.btn_run = QPushButton("추출 실행")
        self.btn_run.setMinimumHeight(36)
        self.btn_run.setToolTip("선택한 파일·모드·규칙으로 표를 추출합니다. 큰 DRM 문서는 수 분 걸릴 수 있습니다.")
        self.btn_run.clicked.connect(self.run_extract)
        self.btn_open = QPushButton("결과 폴더 열기")
        self.btn_open.setToolTip(f"추출된 엑셀과 로그가 저장되는 폴더를 엽니다.\n{OUTPUT_DIR}")
        self.btn_open.clicked.connect(self.open_output)
        run_layout.addWidget(self.btn_run, 1)
        run_layout.addWidget(self.btn_open)
        layout.addLayout(run_layout)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setPlaceholderText("실행하면 진행 상황이 여기에 표시됩니다.")
        layout.addWidget(self.log_view, 1)

        self.load_rules_into_table()

    def show_help(self):
        HelpDialog(self).exec()

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
        self.chk_footer.setChecked(bool(output_opts.get("include_footer", True)))
        self.chk_visual_merge.setChecked(bool(output_opts.get("visual_merge", True)))
        sheet_layout = output_opts.get("sheet_layout", "per_table")
        if sheet_layout == "single_sheet":
            self.radio_sheet_single.setChecked(True)
        else:
            self.radio_sheet_per_table.setChecked(True)
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
            str(rule.get("threshold", 80)),
        ]
        for col, value in enumerate(values):
            item = QTableWidgetItem(value)
            item.setToolTip(RULE_COLUMN_SPECS[col][1])
            self.table.setItem(row, col, item)

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
                threshold = float(self._cell(row, 4) or 80)
            except ValueError:
                threshold = 80
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
            "output": {
                "include_title": self.chk_title.isChecked(),
                "include_footer": self.chk_footer.isChecked(),
                "visual_merge": self.chk_visual_merge.isChecked(),
                "sheet_layout": (
                    "single_sheet" if self.radio_sheet_single.isChecked() else "per_table"
                ),
            },
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

        self.worker = ExtractWorker(
            input_path, rules, extract_all,
            self.chk_title.isChecked(),
            self.chk_footer.isChecked(),
            self.chk_visual_merge.isChecked(),
            "single_sheet" if self.radio_sheet_single.isChecked() else "per_table",
        )
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
