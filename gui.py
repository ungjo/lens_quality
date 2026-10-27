r"""렌즈 품질 검사 GUI. 실행: .\.venv\Scripts\python.exe .\gui.py

검사와 파일 저장은 기존 모듈을 호출합니다. GUI 객체는 GUI 스레드에서만 다루며,
시간이 걸리는 검사/DB 조회/내보내기는 BackgroundTask에서 실행합니다.
"""

import argparse
from contextlib import closing
from dataclasses import dataclass, replace
from pathlib import Path
import sqlite3
import sys
import traceback

import pandas as pd
from PyQt6.QtCore import QThread, Qt, QUrl, pyqtSignal, pyqtSlot, QSignalBlocker
from PyQt6.QtGui import QDesktopServices, QFont, QPalette, QColor
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QComboBox, QTableView,
    QAbstractItemView, QHeaderView, QTabWidget, QSplitter, QPlainTextEdit,
    QProgressBar, QFrame, QMessageBox, QDialog, QDialogButtonBox,
)

from Lens import INPUT, RESULTS, DB_FILE, OUTPUT, InspectionReport, run_inspection, save_result_excel
from export_sql import export_sql
from table_model import DataFrameTableModel, LensFilterProxyModel
from view_db import read_table


TABLE_LABELS = {
    "inspection_results": "검사 결과",
    "measurements": "측정 데이터",
    "criteria": "검사 기준",
    "inspection_details": "항목별 상세",
}
REQUIRED_COLUMNS = {
    "measurements": {"제품번호", "소재종류", "제품명"},
    "criteria": {"검사항목", "적용소재", "하한", "상한"},
    "inspection_results": {"제품번호", "소재종류", "제품명", "판정", "판정사유"},
    "inspection_details": {"제품번호", "검사항목", "측정값", "하한", "상한", "항목판정", "사유"},
}


@dataclass
class DatabaseSnapshot:
    """현재 화면에 표시할 DB 조회 결과. 열린 DB 연결은 보관하지 않습니다."""

    db_path: Path
    frames: dict
    create_sql: dict
    insert_sql: dict
    sql_error: str = ""


def report_frames(report):
    return {
        "inspection_results": report.result,
        "measurements": report.measurements,
        "criteria": report.criteria,
        "inspection_details": report.detail,
    }


def read_database(db_path):
    """네 테이블과 SQL을 같은 시점의 읽기 전용 연결에서 조회합니다."""
    db_path = Path(db_path).resolve()
    if not db_path.is_file():
        raise FileNotFoundError(f"DB 파일을 찾을 수 없습니다: {db_path}")
    with closing(sqlite3.connect(db_path.as_uri() + "?mode=ro", uri=True)) as conn:
        conn.execute("BEGIN")
        frames, creates = {}, {}
        for table in TABLE_LABELS:
            row = conn.execute(
                "SELECT sql FROM sqlite_master WHERE type='table' AND name=?", (table,)
            ).fetchone()
            if row is None:
                raise ValueError(f"렌즈 검사 DB에 필요한 테이블이 없습니다: {table}")
            frames[table] = read_table(conn, table)
            missing = REQUIRED_COLUMNS[table] - set(frames[table].columns)
            if missing:
                raise ValueError(f"{table} 테이블에 필요한 컬럼이 없습니다: {', '.join(sorted(missing))}")
            creates[table] = row[0].rstrip(";") + ";"
        inserts = {table: [] for table in TABLE_LABELS}
        prefixes = {table: f'INSERT INTO "{table}" ' for table in TABLE_LABELS}
        for statement in conn.iterdump():
            for table, prefix in prefixes.items():
                if statement.startswith(prefix):
                    inserts[table].append(statement)
                    break
    return DatabaseSnapshot(
        db_path, frames, creates, {table: "\n".join(rows) for table, rows in inserts.items()}
    )


class BackgroundTask(QThread):
    """작업 결과만 시그널로 전달합니다. 여기서 창이나 표 모델을 수정하지 않습니다."""

    progress = pyqtSignal(str) #현재 진행 상황
    succeeded = pyqtSignal(str, object) #작업 완료와 결과 데이터
    failed = pyqtSignal(str, str, str) #오류 내용

    def __init__(self, operation, arguments, parent=None):
        super().__init__(parent)
        self.operation = operation
        self.arguments = arguments

    def run(self):
        try:
            if self.operation == "inspect": #Excel을 읽어 검사하고 결과 저장
                report = run_inspection(*self.arguments, progress_callback=self.progress.emit) #Lens.py의 run_inspection(검사함수) 호출, self : LensWindow객체를 의미
                # 검사/저장 성공과 SQL 미리보기 실패를 구분하여 결과를 잃지 않습니다.
                try:
                    snapshot = read_database(report.db_path)
                    snapshot.frames = report_frames(report)
                except Exception as exc:
                    snapshot = DatabaseSnapshot(
                        report.db_path, report_frames(report), {}, {}, str(exc)
                    )
                value = (report, snapshot)
            elif self.operation == "load": #기존 DB 조회
                self.progress.emit("저장된 DB 읽는 중")
                value = read_database(*self.arguments)
            elif self.operation == "sql": #DB 내용을 .sql 파일로 내보내기
                self.progress.emit("전체 SQL 파일 내보내는 중")
                value = export_sql(*self.arguments)
            elif self.operation == "excel": #검사한 결과를 Excel로 저장
                self.progress.emit("결과 Excel 저장 중")
                report, path = self.arguments
                save_result_excel(report, path)
                value = report
            else:
                raise ValueError(f"지원하지 않는 작업입니다: {self.operation}")
            self.succeeded.emit(self.operation, value)
        except Exception as exc:
            self.failed.emit(self.operation, str(exc), traceback.format_exc())


def plain_label(text="", object_name=None):
    label = QLabel(text)
    label.setTextFormat(Qt.TextFormat.PlainText)
    if object_name:
        label.setObjectName(object_name)
    return label


def configure_table(view):
    view.setAlternatingRowColors(False)
    view.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    view.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
    view.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    view.setWordWrap(False)
    view.setSortingEnabled(True)
    view.setShowGrid(True)
    view.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
    view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
    view.horizontalHeader().setMinimumSectionSize(65)
    view.verticalHeader().setDefaultSectionSize(30)


def size_columns(view, model):
    view.resizeColumnsToContents()
    for column in range(model.columnCount()):
        name = model.headerData(column, Qt.Orientation.Horizontal)
        maximum = 320 if name in ("판정사유", "사유") else 240
        view.setColumnWidth(column, max(85, min(view.columnWidth(column) + 12, maximum)))


class TablePanel(QWidget):
    """표 하나와 그 표에 적용 가능한 검색·필터를 묶습니다."""

    selection_changed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        filters = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("제품번호, 제품명 또는 값 검색")
        self.search.setClearButtonEnabled(True)
        self.status_filter = QComboBox()
        self.status_filter.addItems(["전체 판정", "합격", "불합격", "판정불가"])
        self.status_filter.setMinimumWidth(110)
        self.material_filter = QComboBox()
        self.material_filter.addItem("전체 소재")
        self.material_filter.setMinimumWidth(140)
        self.count_label = plain_label("표시 0 / 전체 0", "muted")
        filters.addWidget(self.search, 1)
        filters.addWidget(self.status_filter)
        filters.addWidget(self.material_filter)
        filters.addWidget(self.count_label)
        layout.addLayout(filters)

        self.model = DataFrameTableModel(parent=self) #self.model : DataFrame의 값과 셀 표시 정보 제공. DataFrameTableModel : 값,글자,배경색 제공
        self.proxy = LensFilterProxyModel(self) #self.proxy : 검색/필터/정렬을 적용, LensFilterProxyModel : 검색, 필터, 정렬
        self.proxy.setSourceModel(self.model)
        self.view = QTableView() #self.view : 사용자가 보는 표를 화면에 표시, QTableView : 화면에 표 표시
        self.view.setModel(self.proxy)
        configure_table(self.view)
        layout.addWidget(self.view)
        self.search.textChanged.connect(self.apply_filters)
        self.status_filter.currentIndexChanged.connect(self.apply_filters)
        self.material_filter.currentIndexChanged.connect(self.apply_filters)
        self.view.selectionModel().selectionChanged.connect(lambda *_: self.selection_changed.emit())
        self.proxy.modelReset.connect(self.update_count)
        self.proxy.layoutChanged.connect(self.update_count)
        self.proxy.rowsInserted.connect(self.update_count)
        self.proxy.rowsRemoved.connect(self.update_count)
        self.set_dataframe(pd.DataFrame())

    def set_dataframe(self, dataframe, details=None):
        blockers = [QSignalBlocker(widget) for widget in (self.search, self.status_filter, self.material_filter)]
        self.search.clear()
        self.status_filter.setCurrentIndex(0)
        self.material_filter.clear()
        self.material_filter.addItem("전체 소재")
        self.status_filter.setEnabled(any(name in dataframe.columns for name in ("판정", "항목판정")))
        self.material_filter.setEnabled("소재종류" in dataframe.columns)
        if "소재종류" in dataframe.columns:
            self.material_filter.addItems(sorted(dataframe["소재종류"].dropna().astype(str).unique()))
        self.model.set_dataframe(dataframe, details)
        self.proxy.sort(-1)
        self.proxy.set_filters()
        del blockers
        size_columns(self.view, self.model)
        self.update_count()

    def apply_filters(self, *_):
        # 필터 변경으로 다른 제품이 자동 선택되지 않도록 현재 선택을 초기화합니다.
        self.view.selectionModel().clear()
        status = self.status_filter.currentText() if self.status_filter.isEnabled() and self.status_filter.currentIndex() else None
        material = self.material_filter.currentText() if self.material_filter.isEnabled() and self.material_filter.currentIndex() else None
        self.proxy.set_filters(self.search.text(), status, material)
        self.update_count()
        self.selection_changed.emit()

    def update_count(self, *_):
        self.count_label.setText(f"표시 {self.proxy.rowCount()} / 전체 {self.model.rowCount()}")

    def selected_record(self):
        rows = self.view.selectionModel().selectedRows()
        if not rows:
            return None
        return self.proxy.record_at(rows[0].row())


class LensWindow(QMainWindow):
    """버튼 동작과 작업 결과를 연결하는 메인 창입니다."""

    def __init__(self, input_path=INPUT, output_dir=RESULTS):
        super().__init__()
        self.setWindowTitle("렌즈 품질 검사")
        self.setMinimumSize(960, 640)
        self.resize(1320, 930)
        self.report = None
        self.snapshot = None
        self._thread = None
        self._updating = False
        self._build_ui(input_path, output_dir)
        self._update_detail_layout()
        self._refresh_actions()

    def _build_ui(self, input_path, output_dir):
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(20, 14, 20, 12)
        layout.setSpacing(8)

        header = QHBoxLayout()
        heading = QVBoxLayout()
        heading.setSpacing(2)
        heading.addWidget(plain_label("렌즈 품질 검사", "title"))
        heading.addWidget(plain_label("파일 선택부터 판정 확인까지, 한 화면에서 진행하세요.", "muted"))
        header.addLayout(heading, 1)
        self.load_button = QPushButton("저장된 DB 열기")
        self.load_button.clicked.connect(self.choose_database)
        header.addWidget(self.load_button)
        help_button = QPushButton("도움말")
        help_button.clicked.connect(self.show_help)
        header.addWidget(help_button)
        layout.addLayout(header)

        settings = QFrame()
        settings.setObjectName("card")
        form = QGridLayout(settings)
        form.setContentsMargins(12, 10, 12, 10)
        self.input_edit = QLineEdit(str(input_path))
        self.output_edit = QLineEdit(str(output_dir))
        self.input_button = QPushButton("파일 선택")
        self.output_button = QPushButton("폴더 선택")
        self.input_button.clicked.connect(self.choose_input)
        self.output_button.clicked.connect(self.choose_output)
        form.addWidget(plain_label("원본 Excel"), 0, 0)
        form.addWidget(self.input_edit, 0, 1)
        form.addWidget(self.input_button, 0, 2)
        form.addWidget(plain_label("결과 저장 폴더"), 1, 0)
        form.addWidget(self.output_edit, 1, 1)
        form.addWidget(self.output_button, 1, 2)
        self.run_button = QPushButton("검사 실행") #버튼 객체 생성
        self.run_button.setObjectName("runButton") #스타일 적용시 구분할이름 지정
        self.run_button.setMinimumSize(140, 76) #버튼의 최소 크기를 지정
        self.run_button.clicked.connect(self.start_inspection) #클릭했을 때 실행할 메서드 등록. self : LensWindow객체를 의미
        form.addWidget(self.run_button, 0, 3, 2, 1)
        hint = plain_label("검사를 실행하면 선택한 폴더의 DB와 결과 Excel이 최신 데이터로 갱신됩니다.", "muted")
        form.addWidget(hint, 2, 1, 1, 3)
        form.setColumnStretch(1, 1)
        layout.addWidget(settings)

        summary = QHBoxLayout()
        self.count_labels = {}
        colors = {"전체": "#1e3a5f", "합격": "#28704a", "불합격": "#b34754", "판정불가": "#8a630e"}
        for name in colors:
            card = QFrame()
            card.setObjectName("card")
            box = QHBoxLayout(card)
            box.setContentsMargins(14, 8, 14, 8)
            box.addWidget(plain_label(name + (" 제품" if name == "전체" else ""), "muted"))
            number = plain_label("—")
            number.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            number.setStyleSheet(f"font-size: 25px; font-weight: 700; color: {colors[name]};")
            box.addWidget(number)
            self.count_labels[name] = number
            summary.addWidget(card, 1)
        layout.addLayout(summary)

        self.source_label = plain_label("아직 결과가 없습니다. 검사를 실행하거나 저장된 DB를 열어주세요.", "muted")
        self.source_label.setWordWrap(True)
        self.source_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(self.source_label)

        self.tabs = QTabWidget()
        self.panels = {}
        result_page = QWidget()
        result_layout = QVBoxLayout(result_page)
        result_layout.setContentsMargins(0, 0, 0, 0)
        splitter = QSplitter(Qt.Orientation.Vertical)
        self.result_splitter = splitter
        self.panels["inspection_results"] = TablePanel()
        splitter.addWidget(self.panels["inspection_results"])
        selected = QWidget()
        self.detail_container = selected
        selected_layout = QVBoxLayout(selected)
        selected_layout.setContentsMargins(10, 4, 10, 8)
        self.product_label = plain_label("제품을 선택하면 아래에 항목별 검사 결과가 표시됩니다.", "productTitle")
        self.reason_label = plain_label("", "muted")
        self.reason_label.setWordWrap(True)
        selected_layout.addWidget(self.product_label)
        selected_layout.addWidget(self.reason_label)
        self.detail_model = DataFrameTableModel(parent=self)
        self.detail_proxy = LensFilterProxyModel(self)
        self.detail_proxy.setSourceModel(self.detail_model)
        self.detail_view = QTableView()
        self.detail_view.setModel(self.detail_proxy)
        configure_table(self.detail_view)
        selected_layout.addWidget(self.detail_view)
        splitter.addWidget(selected)
        splitter.setSizes([330, 220])
        splitter.setChildrenCollapsible(False)
        result_layout.addWidget(splitter)
        self.compact_detail_button = QPushButton("선택 제품 상세 보기")
        self.compact_detail_button.clicked.connect(self.show_detail_dialog)
        result_layout.addWidget(self.compact_detail_button)
        self.tabs.addTab(result_page, "검사 결과")
        self.panels["inspection_results"].selection_changed.connect(self.show_selected_product)

        for name in ("measurements", "criteria", "inspection_details"):
            panel = TablePanel()
            self.panels[name] = panel
            self.tabs.addTab(panel, TABLE_LABELS[name])
        self._build_sql_tab()
        layout.addWidget(self.tabs, 1)

        actions = QHBoxLayout()
        self.open_excel_button = QPushButton("결과 Excel 열기")
        self.save_excel_button = QPushButton("Excel 다른 이름으로 저장")
        self.sql_export_button = QPushButton("SQL 내보내기")
        self.folder_button = QPushButton("결과 폴더 열기")
        self.open_excel_button.clicked.connect(self.open_result_excel)
        self.save_excel_button.clicked.connect(self.choose_excel_export)
        self.sql_export_button.clicked.connect(self.choose_sql_export)
        self.sql_export_button.setToolTip("화면 필터와 관계없이 현재 DB의 네 테이블 전체를 내보냅니다.")
        self.folder_button.clicked.connect(self.open_result_folder)
        for button in (self.open_excel_button, self.save_excel_button, self.sql_export_button, self.folder_button):
            actions.addWidget(button)
        actions.addStretch()
        layout.addLayout(actions)

        status = QHBoxLayout()
        self.status_label = plain_label("준비됨")
        self.status_label.setWordWrap(True)
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedWidth(180)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        status.addWidget(self.status_label, 1)
        status.addWidget(self.progress_bar)
        layout.addLayout(status)

    def _build_sql_tab(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        controls = QHBoxLayout()
        self.sql_table = QComboBox()
        self.sql_table.addItem("전체 테이블", None)
        for name, label in TABLE_LABELS.items():
            self.sql_table.addItem(f"{label} · {name}", name)
        self.sql_mode = QComboBox()
        self.sql_mode.addItems(["전체 SQL", "구조 · CREATE TABLE", "데이터 · INSERT INTO"])
        self.copy_sql_button = QPushButton("SQL 복사")
        self.copy_sql_button.clicked.connect(self.copy_sql)
        controls.addWidget(self.sql_table, 1)
        controls.addWidget(self.sql_mode)
        controls.addWidget(self.copy_sql_button)
        layout.addLayout(controls)
        layout.addWidget(plain_label("최근 검사 또는 DB 조회 시점의 내용입니다. 여기서는 SQL을 실행하지 않습니다.", "muted"))
        self.sql_text = QPlainTextEdit()
        self.sql_text.setReadOnly(True)
        self.sql_text.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.sql_text.setFont(QFont("Consolas", 10))
        self.sql_text.setPlaceholderText("검사를 실행하거나 저장된 DB를 열면 구조와 INSERT 문장이 표시됩니다.")
        layout.addWidget(self.sql_text)
        self.sql_table.currentIndexChanged.connect(self.update_sql_preview)
        self.sql_mode.currentIndexChanged.connect(self.update_sql_preview)
        self.tabs.addTab(page, "SQL 구조 · 데이터")

    def choose_input(self):
        path, _ = QFileDialog.getOpenFileName(self, "원본 Excel 선택", self.input_edit.text(), "Excel 파일 (*.xlsx)")
        if path:
            self.input_edit.setText(path)

    def choose_output(self):
        path = QFileDialog.getExistingDirectory(self, "결과 저장 폴더 선택", self.output_edit.text())
        if path:
            self.output_edit.setText(path)

    def choose_database(self):
        if self._thread is not None:
            return
        initial = self.snapshot.db_path if self.snapshot else DB_FILE
        path, _ = QFileDialog.getOpenFileName(self, "저장된 검사 DB 열기", str(initial), "SQLite DB (*.db *.sqlite *.sqlite3)")
        if path:
            self.load_database(path)

    def load_database(self, path): #DB를 선택하면 실행. run_inspection()은 호출하지 않음(->검사는 하지 않음)
        return self._start_task("load", (Path(path).resolve(),)) #load이므로 DB에 이미 저장되어있는 inspection_results와 inspection_details를 이용함.

    def start_inspection(self): #다른 작업이 실행 중인지 확인하기 위한 메서드
        if self._thread is not None:
            return
        try:
            if not self.input_edit.text().strip() or not self.output_edit.text().strip():
                raise ValueError("원본 Excel과 결과 저장 폴더를 모두 선택하세요.")
            source = Path(self.input_edit.text().strip()).resolve() #화면에 입력된 원본 Excel 경로를 가져옴
            destination = Path(self.output_edit.text().strip()).resolve() #화면에 입력된 결과 저장 폴더 경로를 가져옴
            if not source.is_file(): #파일이 존재하는지 확인
                raise FileNotFoundError(f"원본 Excel 파일을 찾을 수 없습니다: {source}")
            if destination.exists() and not destination.is_dir(): #파일을 저장할 위치가 올바른지 확인
                raise ValueError("결과 저장 위치에는 파일이 아닌 폴더를 선택하세요.")
        except (ValueError, OSError) as exc:
            self._message("입력 확인", str(exc), warning=True)
            return
        # 실패한 새 검사와 이전 결과가 섞이지 않도록 화면 상태부터 초기화합니다.
        self._clear_dataset() #이전 결과화면을 초기화함
        self._start_task("inspect", (source, destination)) # 화면에 입력한 원본 Excel경로와 결과 저장 폴더 경로를 이용해서 검사 작업 준비

    def _start_task(self, operation, arguments):
        if self._thread is not None:
            return False
        task = BackgroundTask(operation, arguments, self) #BackgroundTask : 검사나 파일 저장을 진행하는 동안 창이 멈추지 않도록 작업을 별도 스레드에서 실행하는 클래스
        self._thread = task  # finished 전까지 참조를 유지하여 실행 중 삭제를 막습니다.
        task.progress.connect(self._on_progress)
        task.succeeded.connect(self._on_success)
        task.failed.connect(self._on_failure)
        task.finished.connect(self._on_finished)
        self.progress_bar.setRange(0, 0)
        self.status_label.setText("작업을 시작합니다…")
        self._refresh_actions()
        task.start()
        return True

    @pyqtSlot(str)
    def _on_progress(self, text):
        self.status_label.setText(text)

    @pyqtSlot(str, object)
    def _on_success(self, operation, value):
        if operation == "inspect":
            self.report, snapshot = value
            self._set_snapshot(snapshot, "현재 검사 결과")
            self.tabs.setCurrentIndex(0)
            self._show_excel_status()
        elif operation == "load":
            # DB만으로 원본 Excel 경로나 Excel 저장 성공 여부를 추측하지 않습니다.
            self.report = None
            self._set_snapshot(value, "저장된 DB 조회")
            self.status_label.setText("DB 조회 완료 · 데이터와 SQL을 확인할 수 있습니다.")
        elif operation == "sql":
            self.status_label.setText(f"SQL 저장 완료: {value}")
        elif operation == "excel":
            self.report = value
            self._show_excel_status()
        self._refresh_actions()

    def _show_excel_status(self):
        if self.report.excel_saved:
            self.status_label.setText(f"저장 완료 · DB 및 Excel: {self.report.excel_path}")
        else:
            self.status_label.setText("DB 저장 완료 · Excel 저장 실패 — 아래 버튼으로 저장을 다시 시도하세요.")
            self._message(
                "Excel 저장 확인",
                "검사와 DB 저장은 완료됐습니다.\n"
                "열려 있는 결과 Excel을 닫거나 다른 파일 이름으로 다시 저장하세요.\n\n"
                + (self.report.excel_error or "Excel 파일을 저장하지 못했습니다."),
                warning=True,
            )

    @pyqtSlot(str, str, str)
    def _on_failure(self, operation, message, detail):
        label = {"inspect": "검사", "load": "DB 조회", "sql": "SQL 내보내기", "excel": "Excel 저장"}[operation]
        suffix = " · 이전에 조회한 결과를 유지합니다." if operation == "load" and self.snapshot else ""
        self.status_label.setText(f"{label} 실패{suffix}")
        if operation == "inspect":
            self.source_label.setText("검사에 실패하여 표시할 결과가 없습니다. 입력 파일과 오류 내용을 확인하세요.")
        self._message(f"{label} 실패", message, details=detail, warning=True)

    @pyqtSlot()
    def _on_finished(self):
        task = self._thread
        self._thread = None
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(1)
        if task is not None:
            task.deleteLater()
        self._refresh_actions()

    def _set_snapshot(self, snapshot, description): #네개의 표와 전체/함격/불합격 개수를 화면에 표시
        self._updating = True
        self.snapshot = snapshot
        details = snapshot.frames["inspection_details"]
        for name, panel in self.panels.items():
            frame = snapshot.frames[name]
            if name == "inspection_results":
                # 판정과 사유를 먼저 보여 주되 DB/원본 DataFrame의 열 순서는 유지합니다.
                leading = ["제품번호", "판정", "소재종류", "제품명", "판정사유"]
                frame = frame.loc[:, leading + [column for column in frame.columns if column not in leading]]
            panel.set_dataframe(frame, details if name == "inspection_results" else None)
        results = snapshot.frames["inspection_results"]
        self.count_labels["전체"].setText(str(len(results)))
        counts = results["판정"].value_counts() #함격,불합격,판정불가 판정의 개수 집계
        for name in ("합격", "불합격", "판정불가"):
            self.count_labels[name].setText(str(int(counts.get(name, 0))))
        source = f"{description} · 전체 검사 집계 (필터와 별개)\nDB: {snapshot.db_path}"
        tooltip = source
        if self.report is not None:
            tooltip += f"\n원본 Excel: {self.report.input_path}"
        self.source_label.setText(source)
        self.source_label.setToolTip(tooltip)
        self.sql_table.setCurrentIndex(0)
        self.sql_mode.setCurrentIndex(0)
        self.update_sql_preview()
        self._updating = False
        self.show_selected_product() #표에서 제품행을 클릭했을 때 검사 결과와 사유 표시

    def _clear_dataset(self):
        self._updating = True
        self.report = None
        self.snapshot = None
        for panel in self.panels.values():
            panel.set_dataframe(pd.DataFrame())
        for label in self.count_labels.values():
            label.setText("—")
        self.source_label.setText("새 검사를 진행하고 있습니다.")
        self.sql_text.clear()
        self._updating = False
        self.show_selected_product()

    def show_selected_product(self):
        if self._updating:
            return
        record = self.panels["inspection_results"].selected_record()
        if record is None or self.snapshot is None:
            self.product_label.setText("제품을 선택하면 아래에 항목별 검사 결과가 표시됩니다.")
            self.reason_label.clear()
            self.detail_model.set_dataframe()
            self._update_detail_layout()
            return
        product_id = str(record["제품번호"])
        full_details = self.snapshot.frames["inspection_details"]
        selected = full_details.loc[full_details["제품번호"].astype(str) == product_id]
        self.product_label.setText(f"{product_id} · {record.get('제품명', '')} · {record.get('판정', '')}")
        self.reason_label.setText(str(record.get("판정사유") or ""))
        self.detail_model.set_dataframe(selected)
        self.detail_proxy.sort(-1)
        size_columns(self.detail_view, self.detail_model)
        self._update_detail_layout()


    def _update_detail_layout(self):
        compact = self.height() < 800
        selected = self.detail_model.rowCount() > 0
        was_visible = not self.detail_container.isHidden()
        self.detail_container.setVisible(selected and not compact)
        self.compact_detail_button.setVisible(compact)
        self.compact_detail_button.setEnabled(selected)
        if selected and not compact and not was_visible:
            height = max(self.result_splitter.height(), 350)
            self.result_splitter.setSizes([int(height * 0.6), int(height * 0.4)])

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "compact_detail_button"):
            self._update_detail_layout()

    def show_detail_dialog(self):
        record = self.panels["inspection_results"].selected_record()
        if record is None or self.snapshot is None:
            return
        details = self.snapshot.frames["inspection_details"]
        frame = details.loc[details["제품번호"].astype(str) == str(record["제품번호"])]
        dialog = QDialog(self)
        dialog.setWindowTitle(f'{record["제품번호"]} · 항목별 검사 결과')
        dialog.resize(1020, 560)
        layout = QVBoxLayout(dialog)
        layout.addWidget(plain_label(self.product_label.text(), "productTitle"))
        reason = plain_label(self.reason_label.text(), "muted")
        reason.setWordWrap(True)
        layout.addWidget(reason)
        panel = TablePanel()
        panel.set_dataframe(frame)
        layout.addWidget(panel)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        dialog.exec()

    def update_sql_preview(self, *_): #SQL문장을 화면에 표시
        if self.snapshot is None:
            self.sql_text.clear()
            return
        if self.snapshot.sql_error:
            self.sql_text.setPlainText("SQL 미리보기를 읽지 못했습니다.\n" + self.snapshot.sql_error)
            return
        chosen = self.sql_table.currentData()
        tables = [chosen] if chosen else list(TABLE_LABELS)
        mode = self.sql_mode.currentIndex()
        parts = []
        for table in tables:
            parts.append(f"-- {TABLE_LABELS[table]} ({table})")
            if mode in (0, 1):
                parts.append(self.snapshot.create_sql.get(table, ""))
            if mode in (0, 2):
                parts.append(self.snapshot.insert_sql.get(table) or "-- 저장된 행이 없습니다.")
        self.sql_text.setPlainText("\n\n".join(parts))

    def choose_sql_export(self): #저장 경로를 받아 .sql파일로 내보내기
        if self._thread is not None or self.snapshot is None:
            return
        default = self.snapshot.db_path.with_suffix(".sql")
        name, _ = QFileDialog.getSaveFileName(self, "전체 SQL 내보내기", str(default), "SQL 파일 (*.sql)")
        if not name:
            return
        destination = Path(name)
        if not destination.suffix:
            destination = destination.with_suffix(".sql")
        if destination.suffix.lower() != ".sql":
            self._message("파일 이름 확인", ".sql 확장자로 저장하세요.", warning=True)
            return
        self._start_task("sql", (self.snapshot.db_path, destination.resolve()))

    def choose_excel_export(self): #검사 결과를 다른 이름으로 저장하거나 저장 재시도
        if self._thread is not None or self.report is None:
            return
        name, _ = QFileDialog.getSaveFileName(self, "결과 Excel 저장", str(self.report.excel_path), "Excel 파일 (*.xlsx)")
        if not name:
            return
        destination = Path(name)
        if not destination.suffix:
            destination = destination.with_suffix(".xlsx")
        if destination.suffix.lower() != ".xlsx":
            self._message("파일 이름 확인", ".xlsx 확장자로 저장하세요.", warning=True)
            return
        # 저장 함수가 상태를 수정하므로 작업 스레드에는 보고서의 별도 객체를 전달합니다.
        self._start_task("excel", (replace(self.report), destination.resolve()))

    def _refresh_actions(self): #작업 중이거나 필요한 결과가 업을 때 버튼 비활성화(프로그램 창을 처음 만들었을 때, 검사/조회/저장 작업 할때, 작업 결과를 정상적으로 받았을 때, 작업 스레드가 완전히 종료됐을 때, 작업 스레드가 완전히 종료됐을 때)
        idle = self._thread is None #실행중인 작업이 없으면 idle = True, 실행중인 작업이 있으면 idle = False
        for widget in (self.input_edit, self.output_edit, self.input_button, self.output_button, self.run_button, self.load_button):
            widget.setEnabled(idle) #True(실행중인 작업이 있을 때)면 클릭 가능, False(실행중인 작업이 업을 때)면 클릭 불가능.
        self.sql_export_button.setEnabled(idle and self.snapshot is not None)
        self.save_excel_button.setEnabled(idle and self.report is not None)
        self.open_excel_button.setEnabled(idle and self.report is not None and self.report.excel_saved)
        self.folder_button.setEnabled(idle and self.snapshot is not None)
        self.copy_sql_button.setEnabled(idle and self.snapshot is not None and not self.snapshot.sql_error)
        retry = self.report is not None and not self.report.excel_saved
        self.save_excel_button.setText("Excel 저장 다시 시도" if retry else "Excel 다른 이름으로 저장")
        self.save_excel_button.setToolTip("GUI에서 검사를 실행한 결과를 저장합니다.")
        if self.report is not None:
            self.open_excel_button.setToolTip(str(self.report.excel_path))
        if self.snapshot is not None:
            self.folder_button.setToolTip(str(self.snapshot.db_path.parent))

    def copy_sql(self):
        if self.snapshot is not None and not self.snapshot.sql_error:
            QApplication.clipboard().setText(self.sql_text.toPlainText())
            self.status_label.setText("현재 표시된 SQL을 복사했습니다.")

    def open_result_excel(self):
        if self.report is not None and self.report.excel_saved:
            self._open_path(self.report.excel_path)

    def open_result_folder(self):
        if self.snapshot is not None:
            self._open_path(self.snapshot.db_path.parent)

    def _open_path(self, path):
        path = Path(path).resolve()
        if not path.exists():
            self._message("파일 확인", f"경로를 찾을 수 없습니다: {path}", warning=True)
        elif not QDesktopServices.openUrl(QUrl.fromLocalFile(str(path))):
            self._message("열기 실패", f"연결된 프로그램으로 열 수 없습니다: {path}", warning=True)

    def _message(self, title, text, details="", warning=False):
        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setTextFormat(Qt.TextFormat.PlainText)
        box.setText(text)
        box.setIcon(QMessageBox.Icon.Warning if warning else QMessageBox.Icon.Information)
        if details:
            box.setDetailedText(details)
        box.exec()

    def show_help(self): #'도움말' 창 표시
        dialog = QDialog(self)
        dialog.setWindowTitle("사용 방법")
        dialog.resize(700, 590)
        layout = QVBoxLayout(dialog)
        text = QPlainTextEdit()
        text.setReadOnly(True)
        text.setPlainText(
            "1. 원본 Excel과 결과 저장 폴더를 선택한 뒤 [검사 실행]을 누릅니다.\n"
            "   측정데이터/기준표 시트를 읽어 판정하며, DB와 결과 Excel을 저장합니다.\n"
            "   같은 폴더의 기존 검사 결과는 새 결과로 교체됩니다.\n\n"
            "2. 검사 결과에서 제품을 선택하면 아래에 해당 제품의 항목별 상세가 나옵니다.\n"
            "   합격은 연두, 불합격은 분홍, 원인이 된 측정값은 회색입니다.\n"
            "   작은 창에서는 [선택 제품 상세 보기]로 상세 창을 엽니다.\n"
            "   컬럼 제목을 누르면 정렬됩니다. 긴 내용은 마우스를 올려 확인하세요.\n\n"
            "3. 검색어와 활성화된 판정/소재 필터를 함께 사용할 수 있습니다.\n"
            "   상단 개수는 전체 제품 집계이고, 각 표의 표시 개수만 필터에 따라 바뀝니다.\n\n"
            "4. [저장된 DB 열기]는 기존 데이터를 읽기 전용으로 조회합니다.\n"
            "   DB 조회만으로는 원본 Excel 경로나 결과 Excel 저장 여부를 알 수 없으므로\n"
            "   Excel 관련 버튼은 이 창에서 검사를 실행한 후 사용할 수 있습니다.\n\n"
            "5. SQL 탭에서 테이블 구조와 INSERT 문장을 확인합니다.\n"
            "   [SQL 내보내기]는 필터와 관계없이 현재 DB의 네 테이블 전체를 저장합니다.\n\n"
            "6. Excel 저장만 실패하면 결과는 유지됩니다. 열려 있는 파일을 닫거나\n"
            "   [Excel 저장 다시 시도]에서 다른 이름을 선택하세요.\n\n"
            "작업 중에는 중복 실행과 창 닫기를 잠시 막습니다. 완료된 뒤 닫아주세요.\n"
            "저장 위치를 위에서 바꿔도 현재 조회한 DB와 열기 버튼의 대상은 유지됩니다.\n\n"
            "PowerShell 실행: .\\.venv\\Scripts\\python.exe .\\gui.py\n"
            "터미널 검사: python .\\Lens.py\n"
            "터미널 조회: python .\\view_db.py --help\n"
            "터미널 SQL 내보내기: python .\\export_sql.py"
        )
        layout.addWidget(text)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        dialog.exec()

    def closeEvent(self, event): #저장 작업 중 창이 닫히는 것을 방지
        if self._thread is not None:
            self.status_label.setText("작업이 진행 중입니다. 저장이 끝난 뒤 창을 닫아주세요.")
            event.ignore()
            return
        event.accept()

# 창/버튼/입력창 등의 색상과 모양 지정
STYLE = """
QMainWindow { background: #f3f6fa; }
QWidget { color: #243247; font-size: 13px; }
QLabel#title { font-size: 24px; font-weight: 700; color: #163351; }
QLabel#muted { color: #607086; font-size: 12px; }
QLabel#productTitle { font-size: 14px; font-weight: 600; }
QFrame#card { background: white; border: 1px solid #dce3eb; border-radius: 8px; }
QPushButton { background: white; border: 1px solid #cbd5e1; border-radius: 5px; padding: 8px 14px; }
QPushButton:hover { background: #edf4fb; border-color: #8da7c1; }
QPushButton:disabled { color: #94a3b8; background: #f1f4f7; border-color: #dde3ea; }
QPushButton#runButton { background: #0f766e; color: white; border: 0; font-size: 17px; font-weight: 700; }
QPushButton#runButton:hover { background: #115e59; }
QPushButton#runButton:disabled { background: #8bb5af; }
QLineEdit, QComboBox { background: white; border: 1px solid #cbd5e1; border-radius: 4px; padding: 6px 8px; min-height: 21px; }
QLineEdit:focus, QComboBox:focus { border-color: #4082ac; }
QLineEdit:disabled, QComboBox:disabled { color: #94a3b8; background: #f1f4f7; }
QTabWidget::pane { border: 1px solid #dce3eb; background: white; border-radius: 4px; }
QTabBar::tab { padding: 10px 19px; background: #e9eef4; color: #607086; }
QTabBar::tab:selected { background: white; color: #0f766e; font-weight: 700; }
QTableView { background: white; gridline-color: #e3e8ee; border: 1px solid #dce3eb; selection-background-color: #dbeafe; selection-color: #132e4b; }
QHeaderView::section { background: #eef3f8; color: #34495e; padding: 7px 6px; border: 0; border-bottom: 1px solid #dce3eb; border-right: 1px solid #dce3eb; font-weight: 600; }
QPlainTextEdit { background: white; border: 1px solid #dce3eb; border-radius: 4px; padding: 8px; }
QProgressBar { border: 1px solid #dce3eb; border-radius: 4px; background: #e8edf3; max-height: 9px; }
QProgressBar::chunk { background: #0f766e; border-radius: 4px; }
"""


def configure_application(app):
    """운영체제 테마와 관계없이 판정 배경색과 글자를 읽을 수 있게 설정합니다."""
    app.setApplicationName("렌즈 품질 검사")
    app.setStyle("Fusion")
    app.setFont(QFont("Malgun Gothic", 10))
    palette = QPalette()
    for role, color in (
        (QPalette.ColorRole.Window, "#f3f6fa"), (QPalette.ColorRole.WindowText, "#243247"),
        (QPalette.ColorRole.Base, "#ffffff"), (QPalette.ColorRole.Text, "#243247"),
        (QPalette.ColorRole.Button, "#ffffff"), (QPalette.ColorRole.ButtonText, "#243247"),
        (QPalette.ColorRole.Highlight, "#dbeafe"), (QPalette.ColorRole.HighlightedText, "#132e4b"),
        (QPalette.ColorRole.ToolTipBase, "#ffffff"), (QPalette.ColorRole.ToolTipText, "#243247"),
    ):
        palette.setColor(role, QColor(color))
    app.setPalette(palette)
    app.setStyleSheet(STYLE)


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="렌즈 검사 GUI를 엽니다. 검사는 창의 [검사 실행] 버튼으로 시작합니다.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        add_help=False,
        epilog=rf"""실행:
  python .\gui.py
  .\.venv\Scripts\python.exe .\gui.py
  python .\gui.py --input ".\data\lens_qualityCheck.xlsx" --output ".\results"

--input / --output은 창에 표시할 초기 경로입니다.
창을 열거나 경로를 지정하는 것만으로 검사를 실행하지 않습니다.
[검사 실행]: 상단 원본 Excel로 새로 판정하고 DB·결과 Excel을 저장합니다.
[저장된 DB 열기]: 기존 판정 결과를 읽기 전용으로 조회합니다.
DB만 조회한 상태에서는 Excel 저장·열기 버튼을 사용할 수 없습니다.
[SQL 내보내기]: 화면 필터와 관계없이 현재 DB 전체를 .sql 파일로 저장합니다.
Excel 저장 실패 시에는 [Excel 저장 다시 시도]를 사용할 수 있습니다.
GUI 표에는 HTML 파일이 필요하지 않습니다.

기본 원본 Excel:
  {INPUT}
기본 결과 저장 폴더:
  {RESULTS}
GUI에서 변경한 경로는 view_db.py / export_sql.py의 기본 경로를 바꾸지 않습니다.
별도 폴더의 결과는 GUI에서 해당 DB를 열어 조회·내보내기하세요.

프로젝트 전체 도움말: python .\view_db.py --help
--help / -h는 GUI 창이나 파일을 열지 않고 도움말만 출력합니다.
실행 환경에는 pandas, openpyxl, PyQt6가 설치되어 있어야 합니다.
""",
    )
    parser.add_argument("-h", "--help", action="help", help="도움말을 표시하고 종료")
    parser.add_argument("--input", type=Path, default=INPUT, metavar="EXCEL", help="원본 Excel 초기 경로")
    parser.add_argument("--output", type=Path, default=RESULTS, metavar="FOLDER", help="결과 저장 폴더 초기 경로")
    args = parser.parse_args(argv)
    app = QApplication(sys.argv[:1]) #GUI전체의 실행과 사용자 입력 처리를 관리하는 객체 생성. QApplication 클래스 : 마우스 클릭, 키보드 입력, 글꼴/스타일등 관리, 화면이 입력에 반응하도록 이벤트 루프 실행
    configure_application(app) #글꼴/색상/스타일 적용
    window = LensWindow(args.input, args.output) #이 프로젝트의 메인 창 객체 생성
    screen = app.primaryScreen()
    if screen is not None:
        area = screen.availableGeometry()
        window.resize(min(1320, area.width() - 40), min(930, area.height() - 60))
    window.show() #창을 화면에 표시
    return app.exec() #app.exec() : 클릭이나 키보드 입력을 기다리면서 처리

#실행 후 가장 먼저 실행되고 gui 창을 만듦. main() 호출
if __name__ == "__main__":
    raise SystemExit(main())
