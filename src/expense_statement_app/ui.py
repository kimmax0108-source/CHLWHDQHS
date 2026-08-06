from __future__ import annotations

import json
from datetime import date
from pathlib import Path

from PySide6.QtCore import QDate, Qt, QTimer
from PySide6.QtGui import QColor, QIcon, QPainter, QPen
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSplitter,
    QSplitterHandle,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .models import (
    AllocationRow,
    ExpenseDocument,
    VendorAccount,
    format_money,
    format_quantity,
    normalize_quantity,
    parse_money,
)
from .preview import ExpensePreviewWidget
from .resource import resource_path
from .xlsx_engine import ExpenseXlsxEngine

APP_TITLE = "자재 문서 표준화"


class ElegantSplitterHandle(QSplitterHandle):
    def paintEvent(self, event) -> None:  # noqa: ANN001
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor("#D7E0EC"), 1))
        painter.setBrush(QColor("#AEBED3"))
        if self.orientation() == Qt.Orientation.Horizontal:
            x = self.width() // 2
            painter.drawLine(x, 8, x, max(8, self.height() - 8))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(x - 2, max(8, self.height() // 2 - 18), 4, 36, 2, 2)
        else:
            y = self.height() // 2
            painter.drawLine(10, y, max(10, self.width() - 10), y)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(max(10, self.width() // 2 - 18), y - 2, 36, 4, 2, 2)


class ElegantSplitter(QSplitter):
    def createHandle(self) -> QSplitterHandle:
        return ElegantSplitterHandle(self.orientation(), self)


def build_section_card(title: str) -> tuple[QFrame, QVBoxLayout]:
    frame = QFrame()
    frame.setObjectName("sectionCard")
    layout = QVBoxLayout(frame)
    layout.setContentsMargins(16, 14, 16, 14)
    layout.setSpacing(10)
    title_label = QLabel(title)
    title_label.setObjectName("sectionCardTitle")
    layout.addWidget(title_label)
    frame.setProperty("sectionTitle", title)
    return frame, layout


def item(text: object = "", right: bool = False, editable: bool = True) -> QTableWidgetItem:
    result = QTableWidgetItem(str(text))
    if right:
        result.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
    if not editable:
        result.setFlags(result.flags() & ~Qt.ItemFlag.ItemIsEditable)
    return result


class MainWindow(QMainWindow):
    def __init__(self, home_callback=None) -> None:
        super().__init__()
        self._home_callback = home_callback
        self.setWindowTitle(f"{APP_TITLE} v2.0.0 - 지출결의서")
        self.resize(1550, 980)
        self.setMinimumSize(1180, 760)
        icon_path = resource_path("assets/app.ico")
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        self.engine = ExpenseXlsxEngine()
        self.current_file: Path | None = None
        self._loading = False
        self.preview_timer = QTimer(self)
        self.preview_timer.setSingleShot(True)
        self.preview_timer.setInterval(100)
        self.preview_timer.timeout.connect(self.refresh_preview)

        self._build_toolbar()
        self._build_ui()
        self._connect_signals()
        self.new_document()
        self._apply_style()

    def _build_toolbar(self) -> None:
        # 통합 프로그램은 고정 헤더와 좌측 내비게이션을 사용한다.
        return

    def _icon(self, name: str) -> QIcon:
        path = resource_path(f"assets/{name}.svg")
        return QIcon(str(path)) if path.exists() else QIcon()

    def _build_sidebar(self) -> QFrame:
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(232)
        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(18, 22, 18, 18)
        layout.setSpacing(10)
        brand = QHBoxLayout()
        icon_label = QLabel()
        icon_label.setPixmap(self._icon("expense").pixmap(38, 38))
        names = QVBoxLayout()
        title = QLabel("자재 문서 표준화")
        title.setObjectName("brandTitle")
        subtitle = QLabel("지출결의서")
        subtitle.setObjectName("brandSubtitle")
        names.addWidget(title)
        names.addWidget(subtitle)
        brand.addWidget(icon_label)
        brand.addLayout(names, 1)
        layout.addLayout(brand)
        layout.addSpacing(12)
        nav = QPushButton(self._icon("expense"), "지출결의서 작성")
        nav.setObjectName("navButton")
        nav.setCheckable(True)
        nav.setChecked(True)
        layout.addWidget(nav)
        export_nav = QPushButton(self._icon("excel"), "지출결의서 Excel로 내보내기")
        export_nav.setObjectName("navActionButton")
        export_nav.clicked.connect(self.export_xlsx)
        layout.addWidget(export_nav)
        layout.addStretch(1)
        home = QPushButton(self._icon("home"), "홈으로")
        home.setObjectName("sidebarFooterButton")
        home.clicked.connect(self.return_home)
        layout.addWidget(home)
        version = QLabel("v2.0.0")
        version.setObjectName("versionLabel")
        layout.addWidget(version)
        return sidebar

    def _build_topbar(self) -> QFrame:
        bar = QFrame()
        bar.setObjectName("topbar")
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(22, 12, 22, 12)
        title = QLabel("지출결의서 작성")
        title.setObjectName("pageTitle")
        layout.addWidget(title)
        layout.addStretch(1)
        self.new_button = QPushButton(self._icon("new"), "새로 작성")
        self.load_button = QPushButton(self._icon("folder"), "불러오기")
        self.save_button = QPushButton(self._icon("save"), "저장하기")
        self.export_button = QPushButton(self._icon("excel"), "지출결의서 Excel로 내보내기")
        for button in (self.new_button, self.load_button, self.save_button):
            button.setObjectName("topActionButton")
        self.export_button.setObjectName("primaryButton")
        for button in (self.new_button, self.load_button, self.save_button, self.export_button):
            layout.addWidget(button)
        return bar

    def _build_ui(self) -> None:
        central = QWidget()
        central.setObjectName("appRoot")
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        self.setCentralWidget(central)
        root.addWidget(self._build_sidebar())

        workspace = QWidget()
        workspace.setObjectName("workspace")
        workspace_layout = QVBoxLayout(workspace)
        workspace_layout.setContentsMargins(0, 0, 0, 0)
        workspace_layout.setSpacing(0)
        workspace_layout.addWidget(self._build_topbar())

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(18, 16, 18, 10)
        content_layout.setSpacing(10)
        self.main_splitter = ElegantSplitter(Qt.Orientation.Horizontal)
        self.main_splitter.setObjectName("mainSplitter")
        self.main_splitter.setChildrenCollapsible(False)
        self.main_splitter.setHandleWidth(12)
        self.main_splitter.setOpaqueResize(True)
        input_panel = self._build_input_panel()
        input_panel.setMinimumWidth(590)
        preview_panel = self._build_preview_panel()
        preview_panel.setMinimumWidth(420)
        self.main_splitter.addWidget(input_panel)
        self.main_splitter.addWidget(preview_panel)
        self.main_splitter.setStretchFactor(0, 6)
        self.main_splitter.setStretchFactor(1, 5)
        self.main_splitter.setSizes([800, 650])
        content_layout.addWidget(self.main_splitter, 1)
        footer = QHBoxLayout()
        self.status_label = QLabel("준비 완료")
        self.file_label = QLabel("현재 파일: 새 문서")
        self.file_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        footer.addWidget(self.status_label)
        footer.addStretch(1)
        footer.addWidget(self.file_label)
        content_layout.addLayout(footer)
        workspace_layout.addWidget(content, 1)
        root.addWidget(workspace, 1)

    def _build_input_panel(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        body = QWidget()
        body.setObjectName("expenseInputBody")
        layout = QVBoxLayout(body)
        layout.setContentsMargins(4, 2, 6, 4)
        layout.setSpacing(8)

        layout.addWidget(self._build_base_group())
        layout.addWidget(self._build_allocation_group())
        layout.addWidget(self._build_description_group())
        layout.addWidget(self._build_account_group())

        layout.addStretch(1)
        scroll.setWidget(body)
        return scroll

    def _build_base_group(self) -> QFrame:
        group, container = build_section_card("1. 기본 정보")
        grid = QGridLayout()
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(8)
        self.month_combo = QComboBox()
        for month in range(1, 13):
            self.month_combo.addItem(f"{month:02d}월", month)
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("yyyy.MM.dd")
        self.writer_edit = QLineEdit()
        self.title_edit = QLineEdit()
        self.vat_check = QCheckBox("부가세 포함")
        grid.addWidget(QLabel("지급 월"), 0, 0)
        grid.addWidget(self.month_combo, 0, 1)
        grid.addWidget(QLabel("작성일"), 0, 2)
        grid.addWidget(self.date_edit, 0, 3)
        grid.addWidget(QLabel("작성인"), 0, 4)
        grid.addWidget(self.writer_edit, 0, 5)
        grid.addWidget(QLabel("결제 구분/제목"), 1, 0)
        grid.addWidget(self.title_edit, 1, 1, 1, 4)
        grid.addWidget(self.vat_check, 1, 5)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)
        grid.setColumnStretch(5, 1)
        container.addLayout(grid)
        return group

    def _build_allocation_group(self) -> QFrame:
        group, layout = build_section_card("2. 원가 안분 내역 (현장/업체별)")
        controls = QHBoxLayout()
        controls.setSpacing(8)
        self.add_allocation = QPushButton("＋ 행 추가")
        self.add_allocation.setObjectName("accentButton")
        self.delete_allocation = QPushButton("－ 선택 행 삭제")
        self.up_allocation = QPushButton("↑ 위로")
        self.down_allocation = QPushButton("↓ 아래로")
        for button in (self.add_allocation, self.delete_allocation, self.up_allocation, self.down_allocation):
            controls.addWidget(button)
        controls.addStretch(1)
        layout.addLayout(controls)

        self.allocation_table = QTableWidget(0, 5)
        self.allocation_table.setHorizontalHeaderLabels(["월", "현장명", "업체명", "수량", "금액(원)"])
        self.allocation_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.allocation_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.allocation_table.setAlternatingRowColors(True)
        self.allocation_table.verticalHeader().setDefaultSectionSize(31)
        self.allocation_table.setMinimumHeight(230)
        header = self.allocation_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.allocation_table)

        totals_card = QFrame()
        totals_card.setObjectName("totalsBar")
        totals = QHBoxLayout(totals_card)
        totals.setContentsMargins(10, 7, 10, 7)
        self.row_count_label = QLabel("합계 (0건)")
        self.qty_total_label = QLabel("수량계 0.000")
        self.amount_total_label = QLabel("금액계 0")
        self.row_count_label.setObjectName("totalCaption")
        self.qty_total_label.setObjectName("totalValue")
        self.amount_total_label.setObjectName("totalValue")
        totals.addWidget(self.row_count_label)
        totals.addStretch(1)
        totals.addWidget(self.qty_total_label)
        totals.addSpacing(18)
        totals.addWidget(self.amount_total_label)
        layout.addWidget(totals_card)
        return group

    def _build_description_group(self) -> QFrame:
        group, layout = build_section_card("3. 적요 / 구매품목 (자유 입력)")
        controls = QHBoxLayout()
        controls.setSpacing(8)
        self.add_description = QPushButton("＋ 행 추가")
        self.add_description.setObjectName("accentButton")
        self.delete_description = QPushButton("－ 선택 행 삭제")
        self.up_description = QPushButton("↑ 위로")
        self.down_description = QPushButton("↓ 아래로")
        for button in (self.add_description, self.delete_description, self.up_description, self.down_description):
            controls.addWidget(button)
        controls.addStretch(1)
        layout.addLayout(controls)
        self.description_table = QTableWidget(0, 1)
        self.description_table.setHorizontalHeaderLabels(["적요 / 구매품목"])
        self.description_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.description_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.description_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.description_table.verticalHeader().setDefaultSectionSize(30)
        self.description_table.setMinimumHeight(155)
        layout.addWidget(self.description_table)
        return group

    def _build_account_group(self) -> QFrame:
        group, layout = build_section_card("4. 업체 계좌 정보")
        hint = QLabel("업체명을 기준으로 자동 표시되며 필요한 값은 직접 수정할 수 있습니다.")
        hint.setObjectName("mutedText")
        layout.addWidget(hint)
        controls = QHBoxLayout()
        controls.setSpacing(8)
        self.sync_accounts_button = QPushButton("업체명에서 계좌행 동기화")
        self.sync_accounts_button.setObjectName("accentButton")
        self.add_account_button = QPushButton("＋ 계좌행 추가")
        self.delete_account_button = QPushButton("－ 선택 계좌행 삭제")
        controls.addWidget(self.sync_accounts_button)
        controls.addWidget(self.add_account_button)
        controls.addWidget(self.delete_account_button)
        controls.addStretch(1)
        layout.addLayout(controls)
        self.account_table = QTableWidget(0, 4)
        self.account_table.setHorizontalHeaderLabels(["업체명", "은행", "계좌번호", "예금주"])
        self.account_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.account_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.account_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.account_table.verticalHeader().setDefaultSectionSize(30)
        self.account_table.setMinimumHeight(145)
        layout.addWidget(self.account_table)
        return group

    def _build_preview_panel(self) -> QWidget:
        panel, layout = build_section_card("5. 미리보기")
        panel.setObjectName("previewCard")
        top = QHBoxLayout()
        top.setContentsMargins(0, 0, 0, 0)
        self.page_label = QLabel("1 / 1")
        self.page_label.setObjectName("pageBadge")
        self.zoom_combo = QComboBox()
        for percent in (65, 75, 85, 100, 115, 130, 150):
            self.zoom_combo.addItem(f"{percent}%", percent / 100)
        self.zoom_combo.setCurrentText("100%")
        self.fit_button = QPushButton("페이지 맞춤")
        top.addStretch(1)
        top.addWidget(self.page_label)
        top.addWidget(self.zoom_combo)
        top.addWidget(self.fit_button)
        layout.addLayout(top)
        self.preview_scroll = QScrollArea()
        self.preview_scroll.setWidgetResizable(False)
        self.preview_scroll.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_widget = ExpensePreviewWidget()
        self.preview_scroll.setWidget(self.preview_widget)
        layout.addWidget(self.preview_scroll, 1)
        return panel

    def _connect_signals(self) -> None:
        self.new_button.clicked.connect(self.new_document)
        self.load_button.clicked.connect(self.load_json)
        self.save_button.clicked.connect(self.save_json)
        self.export_button.clicked.connect(self.export_xlsx)
        self.add_allocation.clicked.connect(lambda: self._add_allocation_row())
        self.delete_allocation.clicked.connect(lambda: self._delete_selected(self.allocation_table))
        self.up_allocation.clicked.connect(lambda: self._move_selected(self.allocation_table, -1))
        self.down_allocation.clicked.connect(lambda: self._move_selected(self.allocation_table, 1))
        self.add_description.clicked.connect(lambda: self._add_description_row(""))
        self.delete_description.clicked.connect(lambda: self._delete_selected(self.description_table))
        self.up_description.clicked.connect(lambda: self._move_selected(self.description_table, -1))
        self.down_description.clicked.connect(lambda: self._move_selected(self.description_table, 1))
        self.sync_accounts_button.clicked.connect(self.sync_accounts)
        self.add_account_button.clicked.connect(lambda: self._add_account_row(VendorAccount()))
        self.delete_account_button.clicked.connect(lambda: self._delete_selected(self.account_table))
        self.allocation_table.itemChanged.connect(self._allocation_changed)
        self.zoom_combo.currentIndexChanged.connect(
            lambda: self.preview_widget.set_zoom(float(self.zoom_combo.currentData()))
        )
        self.fit_button.clicked.connect(self.fit_preview)

        for signal in (
            self.month_combo.currentIndexChanged,
            self.date_edit.dateChanged,
            self.writer_edit.textChanged,
            self.title_edit.textChanged,
            self.vat_check.toggled,
            self.description_table.itemChanged,
            self.account_table.itemChanged,
        ):
            signal.connect(self.schedule_refresh)


    def _allocation_changed(self, changed_item: QTableWidgetItem) -> None:
        if self._loading:
            return
        if changed_item.column() == 2:
            self.sync_accounts()
        self.schedule_refresh()

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            * { font-family: 'Pretendard', 'Malgun Gothic', sans-serif; font-size: 12px; color: #172033; }
            QMainWindow, QWidget#appRoot, QWidget#workspace, QWidget#expenseInputBody { background: #F6F8FC; }
            QFrame#sidebar { background: #FFFFFF; border-right: 1px solid #E6EBF3; }
            QLabel#brandTitle { font-size: 18px; font-weight: 800; color: #111827; }
            QLabel#brandSubtitle, QLabel#versionLabel { color: #7B879B; font-size: 11px; }
            QPushButton#navButton, QPushButton#navActionButton, QPushButton#sidebarFooterButton { min-height: 44px; border: 0; border-radius: 12px; padding: 0 14px; text-align: left; background: transparent; color: #334155; font-weight: 600; }
            QPushButton#navButton:hover, QPushButton#navActionButton:hover, QPushButton#sidebarFooterButton:hover { background: #F1F5FB; }
            QPushButton#navButton:checked { background: #E8F1FF; color: #1263E5; font-weight: 800; }
            QFrame#topbar { background: #FFFFFF; border-bottom: 1px solid #E7ECF4; }
            QLabel#pageTitle { font-size: 20px; font-weight: 800; color: #111827; }
            QPushButton { min-height: 34px; border: 1px solid #D8E1EE; border-radius: 10px; padding: 0 13px; background: #FFFFFF; color: #334155; font-weight: 600; }
            QPushButton:hover { background: #F3F7FD; border-color: #B8C8DF; }
            QPushButton#primaryButton { background: #16A34A; border-color: #16A34A; color: white; min-height: 38px; font-weight: 800; }
            QPushButton#primaryButton:hover { background: #15803D; }
            QPushButton#accentButton { background: #EAF7EE; border-color: #BFE6CB; color: #168344; font-weight: 800; }
            QPushButton#accentButton:hover { background: #DCF2E4; border-color: #9CD8AF; }
            QFrame#sectionCard, QFrame#previewCard { background: #FFFFFF; border: 1px solid #E4EAF3; border-radius: 14px; }
            QLabel#sectionCardTitle { font-size: 14px; font-weight: 800; color: #173B70; padding-bottom: 2px; }
            QLabel#mutedText { color: #6B778C; font-size: 11px; }
            QLabel#pageBadge { background: #EEF4FF; color: #1263E5; border: 1px solid #D8E6FF; border-radius: 9px; padding: 4px 9px; font-weight: 800; }
            QFrame#totalsBar { background: #F7FAFE; border: 1px solid #E4EAF3; border-radius: 9px; }
            QLabel#totalCaption { color: #475569; font-weight: 700; }
            QLabel#totalValue { color: #1263E5; font-weight: 800; }
            QLineEdit, QComboBox, QDateEdit { min-height: 34px; border: 1px solid #D9E2EF; border-radius: 9px; padding: 0 10px; background: #FFFFFF; selection-background-color: #DDEBFF; selection-color: #172033; }
            QLineEdit:focus, QComboBox:focus, QDateEdit:focus { border: 1px solid #2F7AF6; }
            QTableWidget { background: #FFFFFF; alternate-background-color: #FAFBFD; border: 1px solid #E1E7F0; border-radius: 9px; gridline-color: #E4E9F1; selection-background-color: #E7F0FF; selection-color: #172033; outline: 0; }
            QTableWidget::item { padding: 4px 7px; }
            QTableWidget::item:selected { background: #E7F0FF; color: #172033; border: 1px solid #2F7AF6; }
            QHeaderView::section { background: #F3F6FB; color: #20304A; padding: 7px; border: 0; border-right: 1px solid #E1E7F0; border-bottom: 1px solid #E1E7F0; font-weight: 800; }
            QSplitter#mainSplitter::handle { background: #F6F8FC; }
            QSplitter#mainSplitter::handle:hover { background: #EDF3FA; }
            QScrollArea { border: 0; background: transparent; }
            QScrollArea > QWidget > QWidget { border: 0; background: transparent; }
            """
        )

    def return_home(self) -> None:
        if callable(self._home_callback):
            self._home_callback(self)
        else:
            self.close()

    def new_document(self) -> None:
        self._loading = True
        today = date.today()
        self.month_combo.setCurrentIndex(max(0, today.month - 1))
        self.date_edit.setDate(QDate(today.year, today.month, today.day))
        self.writer_edit.setText("김윤재")
        self.title_edit.setText("정기결제")
        self.vat_check.setChecked(True)
        self.allocation_table.setRowCount(0)
        self._add_allocation_row(
            int(self.month_combo.currentData() or today.month),
            "",
            "",
            "0.000",
            "0",
            select_row=False,
        )
        self.description_table.setRowCount(0)
        self._add_description_row("철근구입비-정기결제")
        self._add_description_row("부가세 포함")
        self.account_table.setRowCount(0)
        self._clear_initial_selections()
        self.current_file = None
        self.file_label.setText("현재 파일: 새 문서")
        self._loading = False
        self.refresh_preview()


    def _clear_initial_selections(self) -> None:
        for table in (self.allocation_table, self.description_table, self.account_table):
            table.clearSelection()
            table.setCurrentCell(-1, -1)

    def _add_allocation_row(
        self,
        month: int | None = None,
        site: str = "",
        vendor: str = "",
        quantity: str = "0.000",
        amount: str = "0",
        *,
        select_row: bool = True,
    ) -> None:
        row = self.allocation_table.rowCount()
        self.allocation_table.insertRow(row)
        values = [month or int(self.month_combo.currentData() or 1), site, vendor, quantity, amount]
        for col, value in enumerate(values):
            self.allocation_table.setItem(row, col, item(value, right=col >= 3))
        if select_row and not self._loading:
            self.allocation_table.setCurrentCell(row, 0)
        self.schedule_refresh()

    def _add_description_row(self, text: str) -> None:
        row = self.description_table.rowCount()
        self.description_table.insertRow(row)
        self.description_table.setItem(row, 0, item(text))
        self.schedule_refresh()

    def _add_account_row(self, account: VendorAccount) -> None:
        row = self.account_table.rowCount()
        self.account_table.insertRow(row)
        for col, value in enumerate(
            (account.vendor_name, account.bank_name, account.account_number, account.account_holder)
        ):
            self.account_table.setItem(row, col, item(value))
        self.schedule_refresh()

    def _delete_selected(self, table: QTableWidget) -> None:
        rows = sorted({index.row() for index in table.selectedIndexes()}, reverse=True)
        for row in rows:
            table.removeRow(row)
        self.schedule_refresh()

    def _move_selected(self, table: QTableWidget, direction: int) -> None:
        row = table.currentRow()
        target = row + direction
        if row < 0 or target < 0 or target >= table.rowCount():
            return
        values = [table.takeItem(row, col) for col in range(table.columnCount())]
        target_values = [table.takeItem(target, col) for col in range(table.columnCount())]
        for col in range(table.columnCount()):
            table.setItem(row, col, target_values[col])
            table.setItem(target, col, values[col])
        current_col = max(0, table.currentColumn())
        table.setCurrentCell(target, current_col)
        self.schedule_refresh()

    def sync_accounts(self) -> None:
        existing: dict[str, VendorAccount] = {}
        for row in range(self.account_table.rowCount()):
            name = self._text(self.account_table, row, 0)
            if name:
                existing[name] = VendorAccount(
                    name,
                    self._text(self.account_table, row, 1),
                    self._text(self.account_table, row, 2),
                    self._text(self.account_table, row, 3),
                )
        vendor_names: list[str] = []
        for row in range(self.allocation_table.rowCount()):
            name = self._text(self.allocation_table, row, 2)
            if name and name not in vendor_names:
                vendor_names.append(name)
        self._loading = True
        self.account_table.setRowCount(0)
        for name in vendor_names:
            self._add_account_row(existing.get(name, VendorAccount(name, "", "", name)))
        self._loading = False
        self.refresh_preview()

    @staticmethod
    def _text(table: QTableWidget, row: int, column: int) -> str:
        table_item = table.item(row, column)
        return table_item.text().strip() if table_item else ""

    def collect_document(self) -> ExpenseDocument:
        qdate = self.date_edit.date()
        allocations: list[AllocationRow] = []
        for row in range(self.allocation_table.rowCount()):
            site = self._text(self.allocation_table, row, 1)
            vendor = self._text(self.allocation_table, row, 2)
            amount_text = self._text(self.allocation_table, row, 4)
            if not (site or vendor or amount_text):
                continue
            allocations.append(
                AllocationRow(
                    month=int(self._text(self.allocation_table, row, 0) or self.month_combo.currentData()),
                    site_name=site,
                    vendor_name=vendor,
                    quantity=normalize_quantity(self._text(self.allocation_table, row, 3)),
                    amount=parse_money(amount_text),
                )
            )
        descriptions = [
            self._text(self.description_table, row, 0)
            for row in range(self.description_table.rowCount())
            if self._text(self.description_table, row, 0)
        ]
        accounts = [
            VendorAccount(
                self._text(self.account_table, row, 0),
                self._text(self.account_table, row, 1),
                self._text(self.account_table, row, 2),
                self._text(self.account_table, row, 3),
            )
            for row in range(self.account_table.rowCount())
            if self._text(self.account_table, row, 0)
        ]
        return ExpenseDocument(
            payment_month=int(self.month_combo.currentData()),
            written_date=date(qdate.year(), qdate.month(), qdate.day()),
            writer=self.writer_edit.text().strip(),
            payment_title=self.title_edit.text().strip(),
            vat_included=self.vat_check.isChecked(),
            allocations=allocations,
            descriptions=descriptions,
            vendor_accounts=accounts,
        )

    def schedule_refresh(self, *args) -> None:  # noqa: ANN002
        if not self._loading:
            self.preview_timer.start()

    def refresh_preview(self) -> None:
        try:
            document = self.collect_document()
        except Exception as exc:
            self.status_label.setText(str(exc))
            return
        self.preview_widget.set_document(document)
        self.row_count_label.setText(f"합계 ({len(document.allocations)}건)")
        self.qty_total_label.setText(f"수량계 {format_quantity(document.total_quantity)}")
        self.amount_total_label.setText(f"금액계 {format_money(document.total_amount)}")
        try:
            document.validate()
            self.status_label.setText("준비 완료")
        except ValueError as exc:
            self.status_label.setText(str(exc))

    def fit_preview(self) -> None:
        viewport = self.preview_scroll.viewport().size()
        width_zoom = max(0.55, (viewport.width() - 35) / self.preview_widget.PAGE_W)
        height_zoom = max(0.55, (viewport.height() - 35) / self.preview_widget.page_h)
        zoom = min(width_zoom, height_zoom, 1.35)
        self.preview_widget.set_zoom(zoom)
        nearest = min(range(self.zoom_combo.count()), key=lambda i: abs(float(self.zoom_combo.itemData(i)) - zoom))
        self.zoom_combo.setCurrentIndex(nearest)

    def save_json(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "프로젝트 저장", "지출결의서.json", "JSON (*.json)")
        if not path:
            return
        document = self.collect_document()
        Path(path).write_text(json.dumps(document.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
        self.current_file = Path(path)
        self.file_label.setText(f"현재 파일: {self.current_file.name}")
        self.status_label.setText("프로젝트를 저장했습니다.")

    def load_json(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "프로젝트 불러오기", "", "JSON (*.json)")
        if not path:
            return
        try:
            document = ExpenseDocument.from_dict(json.loads(Path(path).read_text(encoding="utf-8")))
            self._load_document(document)
            self.current_file = Path(path)
            self.file_label.setText(f"현재 파일: {self.current_file.name}")
        except Exception as exc:
            QMessageBox.critical(self, "불러오기 실패", str(exc))

    def _load_document(self, document: ExpenseDocument) -> None:
        self._loading = True
        self.month_combo.setCurrentIndex(document.payment_month - 1)
        self.date_edit.setDate(QDate(document.written_date.year, document.written_date.month, document.written_date.day))
        self.writer_edit.setText(document.writer)
        self.title_edit.setText(document.payment_title)
        self.vat_check.setChecked(document.vat_included)
        self.allocation_table.setRowCount(0)
        for row in document.allocations:
            self._add_allocation_row(
                row.month,
                row.site_name,
                row.vendor_name,
                format_quantity(row.quantity),
                format_money(row.amount),
                select_row=False,
            )
        self.description_table.setRowCount(0)
        for text in document.descriptions:
            self._add_description_row(text)
        self.account_table.setRowCount(0)
        for account in document.vendor_accounts:
            self._add_account_row(account)
        self._clear_initial_selections()
        self._loading = False
        self.refresh_preview()

    def export_xlsx(self) -> None:
        try:
            document = self.collect_document()
            document.validate()
        except Exception as exc:
            QMessageBox.warning(self, "입력 확인", str(exc))
            return
        default_name = f"{document.written_date.year}-{document.payment_month:02d}월_지출결의서.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "지출결의서 Excel로 내보내기", default_name, "Excel (*.xlsx)")
        if not path:
            return
        try:
            self.engine.export(document, Path(path))
        except Exception as exc:
            QMessageBox.critical(self, "내보내기 실패", str(exc))
            return
        self.status_label.setText("지출결의서 Excel 파일을 생성했습니다.")
        QMessageBox.information(self, "완료", f"엑셀 파일을 생성했습니다.\n{path}")
