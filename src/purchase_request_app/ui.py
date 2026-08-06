from __future__ import annotations

import json
from datetime import date
from pathlib import Path
from typing import Any

from PySide6.QtCore import QDate, QStandardPaths, Qt, QTimer, QUrl
from PySide6.QtGui import QAction, QColor, QDesktopServices, QFont, QIcon, QPainter, QPen
from PySide6.QtWidgets import (
    QAbstractItemView,
    QAbstractSpinBox,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QStackedWidget,
    QSplitter,
    QSplitterHandle,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from .models import (
    ProjectData,
    QuoteItem,
    StatementItem,
    Vendor,
    build_classification,
    extract_phone_only,
    decimal_text,
    format_money,
    normalize_quote_payment,
    parse_decimal,
    parse_money,
    purchase_payment_text,
)
from .preset_store import PresetStore
from .preview import build_preview_html
from .resource import resource_path
from .xlsx_engine import XlsxTemplateEngine

APP_TITLE = "자재 문서 표준화"


class ElegantSplitterHandle(QSplitterHandle):
    """A quiet, clearly draggable splitter handle used across the editor."""

    def paintEvent(self, event) -> None:  # noqa: ANN001
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor("#D7E0EC"), 1))
        painter.setBrush(QColor("#AEBED3"))
        if self.orientation() == Qt.Orientation.Horizontal:
            x = self.width() // 2
            painter.drawLine(x, 8, x, max(8, self.height() - 8))
            y = max(8, self.height() // 2 - 18)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(x - 2, y, 4, 36, 2, 2)
        else:
            y = self.height() // 2
            painter.drawLine(10, y, max(10, self.width() - 10), y)
            x = max(10, self.width() // 2 - 18)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(x, y - 2, 36, 4, 2, 2)


class ElegantSplitter(QSplitter):
    def createHandle(self) -> QSplitterHandle:
        return ElegantSplitterHandle(self.orientation(), self)


def build_section_card(title: str) -> tuple[QFrame, QVBoxLayout]:
    """Create a white card whose title is fully inside the card boundary."""
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


class PresetNameDialog(QDialog):
    def __init__(self, title: str, initial: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("프리셋 이름"))
        self.edit = QLineEdit(initial)
        self.edit.selectAll()
        layout.addWidget(self.edit)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    @property
    def name(self) -> str:
        return self.edit.text().strip()


def text_item(text: Any = "", editable: bool = True, align_right: bool = False, bold: bool = False) -> QTableWidgetItem:
    value = "" if text is None else str(text)
    item = QTableWidgetItem(value)
    item.setToolTip(value)
    font = QFont(item.font())
    font.setPointSize(10)
    font.setBold(bold)
    item.setFont(font)
    if not editable:
        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
    if align_right:
        item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
    return item


def check_item(checked: bool) -> QTableWidgetItem:
    item = QTableWidgetItem()
    item.setFlags(
        Qt.ItemFlag.ItemIsEnabled
        | Qt.ItemFlag.ItemIsSelectable
        | Qt.ItemFlag.ItemIsUserCheckable
    )
    item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
    return item


def table_text(table: QTableWidget, row: int, column: int) -> str:
    item = table.item(row, column)
    return item.text().strip() if item else ""


class MainWindow(QMainWindow):
    def __init__(self, home_callback=None) -> None:
        super().__init__()
        self._home_callback = home_callback
        self.setWindowTitle(f"{APP_TITLE} v2.0.0 - 구매 품의서 양식")
        self.resize(1600, 1000)
        self.setMinimumSize(1240, 860)
        icon = resource_path("assets/app.ico")
        if icon.exists():
            self.setWindowIcon(QIcon(str(icon)))

        app_data = Path(
            QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppDataLocation)
        )
        self.store = PresetStore(app_data, resource_path("presets/default_presets.json"))
        self.engine = XlsxTemplateEngine()
        self.data = ProjectData()
        self.item_notes: list[str] = []
        self._loading = False
        self._project_path: Path | None = None
        self._last_export: Path | None = None
        self._body_template = ""
        self._common_delivery_place = ""
        self._statement_source_signature = ""
        self._purchase_overrides: set[str] = set()

        self._build_toolbar()
        self._build_ui()
        self.preview_timer = QTimer(self)
        self.preview_timer.setSingleShot(True)
        self.preview_timer.setInterval(120)
        self.preview_timer.timeout.connect(self.refresh_preview)
        self._preview_zoom_steps = 0
        self._connect_signals()
        self._load_presets()
        self._reset_with_defaults()
        self._apply_style()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------
    def _build_toolbar(self) -> None:
        # v2.0은 상단 툴바 대신 고정 헤더와 좌측 내비게이션을 사용한다.
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
        brand_icon = QLabel()
        brand_icon.setPixmap(self._icon("purchase").pixmap(38, 38))
        brand_text = QVBoxLayout()
        title = QLabel("자재 문서 표준화")
        title.setObjectName("brandTitle")
        subtitle = QLabel("구매 품의서 양식")
        subtitle.setObjectName("brandSubtitle")
        brand_text.addWidget(title)
        brand_text.addWidget(subtitle)
        brand.addWidget(brand_icon)
        brand.addLayout(brand_text, 1)
        layout.addLayout(brand)
        layout.addSpacing(10)

        self.nav_group = QButtonGroup(self)
        self.nav_group.setExclusive(True)
        self.nav_buttons: list[QPushButton] = []
        nav_items = [
            ("견적대비표 작성", "quote", 0),
            ("내역서 작성", "statement", 1),
            ("구매품의서 작성", "purchase", 2),
        ]
        for label, icon_name, index in nav_items:
            button = QPushButton(self._icon(icon_name), label)
            button.setObjectName("navButton")
            button.setCheckable(True)
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            button.clicked.connect(lambda _checked=False, i=index: self.set_workspace_page(i))
            self.nav_group.addButton(button, index)
            self.nav_buttons.append(button)
            layout.addWidget(button)

        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setObjectName("sidebarDivider")
        layout.addWidget(divider)

        self.data_nav_button = QPushButton(self._icon("info"), "프로그램 설정")
        self.data_nav_button.setObjectName("navButton")
        self.data_nav_button.setCheckable(True)
        self.data_nav_button.clicked.connect(lambda: self.set_workspace_page(3))
        self.nav_group.addButton(self.data_nav_button, 3)
        layout.addWidget(self.data_nav_button)
        layout.addStretch(1)

        self.home_nav_button = QPushButton(self._icon("home"), "홈으로")
        self.home_nav_button.setObjectName("sidebarFooterButton")
        self.home_nav_button.clicked.connect(self.return_home)
        layout.addWidget(self.home_nav_button)
        self.settings_nav_button = QPushButton("⚙  설정")
        self.settings_nav_button.setObjectName("sidebarFooterButton")
        self.settings_nav_button.clicked.connect(self.choose_template)
        layout.addWidget(self.settings_nav_button)
        version = QLabel("v2.0.0")
        version.setObjectName("versionLabel")
        layout.addWidget(version)
        return sidebar

    def _build_topbar(self) -> QFrame:
        bar = QFrame()
        bar.setObjectName("topbar")
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(22, 12, 22, 12)
        self.page_title_label = QLabel("견적대비표 작성")
        self.page_title_label.setObjectName("pageTitle")
        layout.addWidget(self.page_title_label)
        layout.addStretch(1)

        self.default_layout_button = QPushButton("기본 배치")
        self.default_layout_button.setObjectName("topActionButton")
        self.new_button = QPushButton(self._icon("new"), "새로 작성")
        self.load_button = QPushButton(self._icon("folder"), "불러오기")
        self.save_button = QPushButton(self._icon("save"), "저장하기")
        self.export_button = QPushButton(self._icon("excel"), "구매품의서 Excel로 내보내기")
        for button in (self.default_layout_button, self.new_button, self.load_button, self.save_button):
            button.setObjectName("topActionButton")
        self.export_button.setObjectName("primaryButton")
        self.new_button.clicked.connect(self.new_project)
        self.load_button.clicked.connect(self.load_project)
        self.save_button.clicked.connect(self.save_project)
        self.export_button.clicked.connect(self.export_excel)
        for button in (self.default_layout_button, self.new_button, self.load_button, self.save_button, self.export_button):
            layout.addWidget(button)
        return bar

    def _build_data_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(14)
        title = QLabel("현장·품목 데이터 관리")
        title.setObjectName("sectionPageTitle")
        layout.addWidget(title)
        layout.addWidget(self._build_preset_group())
        guide = QFrame()
        guide.setObjectName("card")
        guide_layout = QVBoxLayout(guide)
        guide_layout.addWidget(QLabel("사용 안내"))
        note = QLabel(
            "현장과 품목 프리셋은 공통으로 저장됩니다. 견적대비표 작성 화면에서 프리셋을 선택하면 "
            "현장명·약칭·품목명·기본 문구가 자동 입력됩니다."
        )
        note.setWordWrap(True)
        note.setObjectName("mutedText")
        guide_layout.addWidget(note)
        layout.addWidget(guide)
        layout.addStretch(1)
        return page

    def _build_overview_card(self) -> QFrame:
        """Compact document summary with an integrated step-navigation footer."""
        card = QFrame()
        card.setObjectName("overviewCard")
        card.setMinimumHeight(176)
        card.setMaximumHeight(190)
        card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)

        header = QHBoxLayout()
        self.overview_title_label = QLabel("문서 요약")
        self.overview_title_label.setObjectName("cardTitle")
        self.overview_step_label = QLabel("1 / 3 · 견적대비표")
        self.overview_step_label.setObjectName("stepBadge")
        header.addWidget(self.overview_title_label)
        header.addStretch(1)
        header.addWidget(self.overview_step_label)
        layout.addLayout(header)

        self.overview_vendor_count = QLabel("0 개")
        self.overview_group_count = QLabel("0 개")
        self.overview_item_count = QLabel("0 개")
        self.overview_average_count = QLabel("0 개")
        self.overview_winner = QLabel("-")
        self.overview_amount = QLabel("₩ 0")

        grid = QGridLayout()
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(5)

        name = QLabel("업체")
        name.setObjectName("summaryName")
        self.overview_vendor_count.setObjectName("summaryValue")
        grid.addWidget(name, 0, 0)
        grid.addWidget(self.overview_vendor_count, 0, 1)

        name = QLabel("대표/상세 품목")
        name.setObjectName("summaryName")
        combined = QWidget()
        combined_layout = QHBoxLayout(combined)
        combined_layout.setContentsMargins(0, 0, 0, 0)
        combined_layout.setSpacing(4)
        for widget in (self.overview_group_count, self.overview_item_count):
            widget.setObjectName("summaryValue")
            combined_layout.addWidget(widget)
        grid.addWidget(name, 0, 2)
        grid.addWidget(combined, 0, 3)

        name = QLabel("평균 반영")
        name.setObjectName("summaryName")
        self.overview_average_count.setObjectName("summaryValue")
        grid.addWidget(name, 1, 0)
        grid.addWidget(self.overview_average_count, 1, 1)

        name = QLabel("선정 업체")
        name.setObjectName("summaryName")
        self.overview_winner.setObjectName("summaryValue")
        self.overview_winner.setWordWrap(False)
        grid.addWidget(name, 1, 2)
        grid.addWidget(self.overview_winner, 1, 3)

        name = QLabel("계약금액")
        name.setObjectName("summaryName")
        self.overview_amount.setObjectName("summaryValue")
        self.overview_amount.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        grid.addWidget(name, 2, 0)
        grid.addWidget(self.overview_amount, 2, 1, 1, 3)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)
        layout.addLayout(grid)

        divider = QFrame()
        divider.setObjectName("overviewDivider")
        divider.setFrameShape(QFrame.Shape.HLine)
        layout.addWidget(divider)

        actions = QHBoxLayout()
        actions.setSpacing(8)
        self.previous_button = QPushButton("← 이전 단계")
        self.next_button = QPushButton("다음 단계 →")
        self.overview_export_button = QPushButton(self._icon("excel"), "Excel로 내보내기")
        self.previous_button.setObjectName("secondaryNavButton")
        self.next_button.setObjectName("primaryNavButton")
        self.overview_export_button.setObjectName("primaryButton")
        self.previous_button.clicked.connect(lambda: self.set_workspace_page(self.tabs.currentIndex() - 1))
        self.next_button.clicked.connect(lambda: self.set_workspace_page(self.tabs.currentIndex() + 1))
        self.overview_export_button.clicked.connect(self.export_excel)
        actions.addWidget(self.previous_button)
        actions.addStretch(1)
        actions.addWidget(self.next_button)
        actions.addWidget(self.overview_export_button)
        layout.addLayout(actions)
        return card

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

        self.tabs = QStackedWidget()
        self.tabs.setMinimumWidth(690)
        self.tabs.addWidget(self._build_quote_tab())
        self.tabs.addWidget(self._build_statement_tab())
        self.tabs.addWidget(self._build_purchase_tab())
        self.tabs.addWidget(self._build_data_page())

        self.preview_panel = self._build_preview_panel()
        self.right_column = QWidget()
        self.right_column.setMinimumWidth(350)
        right_layout = QVBoxLayout(self.right_column)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)
        self.right_splitter = QSplitter(Qt.Orientation.Vertical)
        self.right_splitter.setObjectName("rightSplitter")
        self.right_splitter.setChildrenCollapsible(False)
        self.right_splitter.setHandleWidth(0)
        self.right_splitter.addWidget(self.preview_panel)
        self.overview_card = self._build_overview_card()
        self.right_splitter.addWidget(self.overview_card)
        self.right_splitter.setStretchFactor(0, 1)
        self.right_splitter.setStretchFactor(1, 0)
        self.right_splitter.setSizes([700, 184])
        right_layout.addWidget(self.right_splitter, 1)

        self.main_splitter = ElegantSplitter(Qt.Orientation.Horizontal)
        self.main_splitter.setObjectName("mainSplitter")
        self.main_splitter.setChildrenCollapsible(False)
        self.main_splitter.setHandleWidth(12)
        self.main_splitter.setOpaqueResize(True)
        self.main_splitter.addWidget(self.tabs)
        self.main_splitter.addWidget(self.right_column)
        self.main_splitter.setStretchFactor(0, 5)
        self.main_splitter.setStretchFactor(1, 3)
        self.main_splitter.setSizes([900, 520])

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(18, 16, 18, 10)
        content_layout.setSpacing(10)
        content_layout.addWidget(self.main_splitter, 1)
        footer = QHBoxLayout()
        self.file_status = QLabel("현재 파일: 새 문서")
        self.calculation_status = QLabel("견적업체와 단가를 입력해 주세요.")
        self.calculation_status.setAlignment(Qt.AlignmentFlag.AlignRight)
        footer.addWidget(self.file_status, 1)
        footer.addWidget(self.calculation_status, 1)
        content_layout.addLayout(footer)
        workspace_layout.addWidget(content, 1)
        root.addWidget(workspace, 1)

        self.preview_action = QAction("미리보기 표시", self)
        self.preview_action.setCheckable(True)
        self.preview_action.setChecked(True)
        self.set_workspace_page(0)

    def set_workspace_page(self, index: int) -> None:
        if not hasattr(self, "tabs"):
            return
        index = max(0, min(index, self.tabs.count() - 1))
        self.tabs.setCurrentIndex(index)
        titles = ["견적대비표 작성", "내역서 작성", "구매품의서 작성", "데이터 관리"]
        self.page_title_label.setText(titles[index])
        button = self.nav_group.button(index)
        if button:
            button.setChecked(True)
        if hasattr(self, "preview_panel"):
            self.preview_panel.setVisible(index < 3)
            self.right_column.setVisible(index < 3)
        if hasattr(self, "previous_button"):
            step_names = ["견적대비표", "내역서", "구매품의서"]
            if index < 3:
                self.overview_step_label.setText(f"{index + 1} / 3 · {step_names[index]}")
            self.previous_button.setVisible(index in (1, 2))
            self.next_button.setVisible(index in (0, 1))
            self.overview_export_button.setVisible(index == 2)
        if hasattr(self, "default_layout_button"):
            self.default_layout_button.setVisible(index == 0)
        if hasattr(self, "export_button"):
            self.export_button.setVisible(index == 2)

    def _build_preset_group(self) -> QGroupBox:
        group = QGroupBox("현장·품목 프리셋")
        group.setObjectName("cardGroup")
        grid = QGridLayout(group)
        grid.setContentsMargins(16, 18, 16, 16)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)
        self.site_preset_combo = QComboBox()
        self.item_preset_combo = QComboBox()
        self.site_save_button = QPushButton("현장 저장")
        self.site_delete_button = QPushButton("삭제")
        self.item_save_button = QPushButton("품목 저장")
        self.item_delete_button = QPushButton("삭제")
        grid.addWidget(QLabel("현장 프리셋"), 0, 0)
        grid.addWidget(self.site_preset_combo, 0, 1)
        grid.addWidget(self.site_save_button, 0, 2)
        grid.addWidget(self.site_delete_button, 0, 3)
        grid.addWidget(QLabel("품목 프리셋"), 1, 0)
        grid.addWidget(self.item_preset_combo, 1, 1)
        grid.addWidget(self.item_save_button, 1, 2)
        grid.addWidget(self.item_delete_button, 1, 3)
        grid.setColumnStretch(1, 1)
        return group

    def _build_preview_panel(self) -> QFrame:
        panel = QFrame()
        panel.setObjectName("previewCard")
        panel.setMinimumWidth(390)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        header = QHBoxLayout()
        self.preview_title = QLabel("미리보기")
        self.preview_title.setObjectName("cardTitle")
        header.addWidget(self.preview_title)
        header.addStretch(1)
        layout.addLayout(header)

        self.preview_sheet_combo = QComboBox()
        self.preview_sheet_combo.addItem("견적대비표", "quote")
        self.preview_sheet_combo.addItem("내역서", "statement")
        self.preview_sheet_combo.addItem("구매품의서", "purchase")
        self.preview_sheet_combo.hide()
        self.preview_follow_check = QCheckBox("작성 탭 따라가기")
        self.preview_follow_check.setChecked(True)
        self.preview_follow_check.hide()

        tabs = QHBoxLayout()
        self.preview_button_group = QButtonGroup(self)
        self.preview_button_group.setExclusive(True)
        self.preview_kind_buttons: dict[str, QPushButton] = {}
        for label, kind in (("견적대비표", "quote"), ("내역서", "statement"), ("구매품의서", "purchase")):
            button = QPushButton(label)
            button.setObjectName("previewTabButton")
            button.setCheckable(True)
            button.clicked.connect(lambda _checked=False, k=kind: self.select_preview_kind(k))
            self.preview_button_group.addButton(button)
            self.preview_kind_buttons[kind] = button
            tabs.addWidget(button)
        self.preview_kind_buttons["quote"].setChecked(True)
        layout.addLayout(tabs)

        self.preview_browser = QTextBrowser()
        self.preview_browser.setOpenExternalLinks(False)
        self.preview_browser.setPlaceholderText("입력 내용이 여기에 실시간으로 표시됩니다.")
        layout.addWidget(self.preview_browser, 1)
        return panel

    def select_preview_kind(self, kind: str) -> None:
        index = self.preview_sheet_combo.findData(kind)
        if index >= 0:
            self.preview_sheet_combo.setCurrentIndex(index)
        button = self.preview_kind_buttons.get(kind)
        if button:
            button.setChecked(True)

    def _build_quote_tab(self) -> QWidget:
        page = QWidget()
        outer = QVBoxLayout(page)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(10)

        info, info_layout = build_section_card("견적 기본정보")
        info_grid = QGridLayout()
        info_grid.setContentsMargins(0, 0, 0, 0)
        info_grid.setHorizontalSpacing(10)
        info_grid.setVerticalSpacing(8)
        self.site_name_edit = QLineEdit()
        self.site_short_edit = QLineEdit()
        self.quote_title_edit = QLineEdit()
        self.item_label_edit = QLineEdit()
        self.author_edit = QLineEdit()
        self.quote_date_edit = QDateEdit(QDate.currentDate())
        self.quote_date_edit.setCalendarPopup(True)
        self.quote_date_edit.setDisplayFormat("yyyy.MM.dd")
        info_grid.addWidget(QLabel("현장명"), 0, 0)
        info_grid.addWidget(self.site_name_edit, 0, 1, 1, 5)
        info_grid.addWidget(QLabel("현장 약칭"), 1, 0)
        info_grid.addWidget(self.site_short_edit, 1, 1)
        info_grid.addWidget(QLabel("시트용 품목명"), 1, 2)
        info_grid.addWidget(self.item_label_edit, 1, 3)
        info_grid.addWidget(QLabel("작성자"), 1, 4)
        info_grid.addWidget(self.author_edit, 1, 5)
        info_grid.addWidget(QLabel("품명/견적 제목"), 2, 0)
        info_grid.addWidget(self.quote_title_edit, 2, 1, 1, 3)
        info_grid.addWidget(QLabel("작성일"), 2, 4)
        info_grid.addWidget(self.quote_date_edit, 2, 5)
        info_grid.setColumnStretch(1, 2)
        info_grid.setColumnStretch(3, 2)
        info_grid.setColumnStretch(5, 1)
        info_layout.addLayout(info_grid)
        outer.addWidget(info)

        vendor_group, vendor_layout = build_section_card("1. 견적업체")
        self.vendor_group = vendor_group
        vendor_group.setMinimumHeight(42)
        vendor_buttons = QHBoxLayout()
        vendor_buttons.setSpacing(8)
        self.add_vendor_button = QPushButton("＋ 업체 추가")
        self.add_vendor_button.setObjectName("accentButton")
        self.delete_vendor_button = QPushButton("선택 업체 삭제")
        self.vendor_left_button = QPushButton("← 왼쪽")
        self.vendor_right_button = QPushButton("오른쪽 →")
        for button in (self.add_vendor_button, self.delete_vendor_button, self.vendor_left_button, self.vendor_right_button):
            vendor_buttons.addWidget(button)
        vendor_buttons.addStretch(1)
        vendor_buttons.addWidget(QLabel("공통 납품장소"))
        self.common_delivery_place_edit = QLineEdit()
        self.common_delivery_place_edit.setMinimumWidth(150)
        self.apply_delivery_place_button = QPushButton("전체 업체에 적용")
        vendor_buttons.addWidget(self.common_delivery_place_edit)
        vendor_buttons.addWidget(self.apply_delivery_place_button)
        vendor_layout.addLayout(vendor_buttons)
        self.vendor_table = QTableWidget(0, 8)
        self.vendor_table.setHorizontalHeaderLabels(["업체명", "연락처", "담당자", "결제조건", "납품장소", "납품/설치일", "견적제출", "평균포함"])
        self._configure_edit_table(self.vendor_table)
        self.vendor_table.verticalHeader().setDefaultSectionSize(30)
        header = self.vendor_table.horizontalHeader()
        header.setSectionsMovable(True)
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.vendor_table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        for column, width in enumerate((210, 150, 110, 180, 210, 145, 95, 95)):
            self.vendor_table.setColumnWidth(column, width)
        vendor_layout.addWidget(self.vendor_table, 1)

        item_group, item_layout = build_section_card("2. 대표 품목명 및 상세 품목별 업체 단가")
        self.item_group = item_group
        item_group.setMinimumHeight(48)
        item_buttons = QHBoxLayout()
        item_buttons.setSpacing(8)
        self.add_item_button = QPushButton("＋ 상세 품목 추가")
        self.add_item_button.setObjectName("accentButton")
        self.add_group_item_button = QPushButton("＋ 새 대표 품목")
        self.delete_item_button = QPushButton("선택 행 삭제")
        self.item_up_button = QPushButton("↑ 위로")
        self.item_down_button = QPushButton("↓ 아래로")
        for button in (self.add_item_button, self.add_group_item_button, self.delete_item_button, self.item_up_button, self.item_down_button):
            item_buttons.addWidget(button)
        item_buttons.addStretch(1)
        item_layout.addLayout(item_buttons)
        self.item_table = QTableWidget(0, 6)
        self._configure_edit_table(self.item_table)
        self.item_table.verticalHeader().setDefaultSectionSize(32)
        self.item_table.horizontalHeader().setSectionsMovable(True)
        self.item_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.item_table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        item_layout.addWidget(self.item_table, 1)

        self.quote_splitter = ElegantSplitter(Qt.Orientation.Vertical)
        self.quote_splitter.setObjectName("quoteSplitter")
        self.quote_splitter.setChildrenCollapsible(False)
        self.quote_splitter.setHandleWidth(12)
        self.quote_splitter.setOpaqueResize(True)
        self.quote_splitter.addWidget(vendor_group)
        self.quote_splitter.addWidget(item_group)
        self.quote_splitter.setStretchFactor(0, 1)
        self.quote_splitter.setStretchFactor(1, 1)
        self.quote_splitter.setSizes([280, 300])
        outer.addWidget(self.quote_splitter, 1)

        summary_group, summary_container = build_section_card("3. 업체별 순위 및 최종선정")
        self.summary_group = summary_group
        summary_group.setMinimumHeight(248)
        summary_group.setMaximumHeight(248)
        summary_group.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        summary_layout = QHBoxLayout()
        summary_layout.setContentsMargins(0, 0, 0, 0)
        summary_layout.setSpacing(10)

        self.summary_table = QTableWidget(0, 6)
        self.summary_table.setHorizontalHeaderLabels(["업체명", "공급가액", "부가세", "부가세 포함금액", "순위", "평균반영"])
        self.summary_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.summary_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.summary_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.summary_table.verticalHeader().setDefaultSectionSize(28)
        self.summary_table.setMinimumWidth(280)
        self.summary_table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        summary_header = self.summary_table.horizontalHeader()
        summary_header.setSectionsMovable(False)
        summary_header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        for column, width in enumerate((118, 96, 84, 122, 58, 68)):
            self.summary_table.setColumnWidth(column, width)
        summary_layout.addWidget(self.summary_table, 1)

        selection_card = QFrame()
        selection_card.setObjectName("selectionCard")
        selection_card.setMinimumWidth(178)
        selection_card.setMaximumWidth(220)
        selection_layout = QVBoxLayout(selection_card)
        selection_layout.setContentsMargins(12, 11, 12, 10)
        selection_layout.setSpacing(7)
        selection_title = QLabel("최종 선정업체")
        selection_title.setObjectName("subCardTitle")
        selection_layout.addWidget(selection_title)
        self.selected_vendor_combo = QComboBox()
        self.selected_vendor_combo.setMinimumWidth(150)
        self.selected_vendor_combo.setFixedHeight(30)
        self.selected_vendor_combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.selected_vendor_combo.setMinimumContentsLength(12)
        selection_layout.addWidget(self.selected_vendor_combo)
        self.selected_vendor_rank_label = QLabel("-")
        self.selected_vendor_rank_label.setObjectName("rankBadge")
        self.selected_vendor_rank_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        selection_layout.addWidget(self.selected_vendor_rank_label, 0, Qt.AlignmentFlag.AlignLeft)
        selected_metrics = QGridLayout()
        selected_metrics.setHorizontalSpacing(7)
        selected_metrics.setVerticalSpacing(5)
        self.selected_vendor_supply_label = QLabel("₩ 0")
        self.selected_vendor_vat_label = QLabel("₩ 0")
        self.selected_vendor_total_label = QLabel("₩ 0")
        for row, (name, value) in enumerate((("공급가액", self.selected_vendor_supply_label), ("부가세", self.selected_vendor_vat_label), ("합계", self.selected_vendor_total_label))):
            name_label = QLabel(name)
            name_label.setObjectName("summaryName")
            value.setObjectName("fixedMetricValue")
            value.setWordWrap(False)
            value.setMinimumWidth(116)
            value.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            selected_metrics.addWidget(name_label, row, 0)
            selected_metrics.addWidget(value, row, 1)
        selected_metrics.setColumnStretch(1, 1)
        selection_layout.addLayout(selected_metrics)
        selection_layout.addStretch(1)
        summary_layout.addWidget(selection_card)

        budget_card = QFrame()
        budget_card.setObjectName("budgetCard")
        budget_card.setMinimumWidth(205)
        budget_card.setMaximumWidth(260)
        budget_grid = QGridLayout(budget_card)
        budget_grid.setContentsMargins(12, 11, 12, 10)
        budget_grid.setHorizontalSpacing(7)
        budget_grid.setVerticalSpacing(5)
        budget_title = QLabel("가실행 금액")
        budget_title.setObjectName("subCardTitle")
        budget_grid.addWidget(budget_title, 0, 0, 1, 2)
        self.budget_mode_combo = QComboBox()
        self.budget_mode_combo.addItem("업체 평균", "average")
        self.budget_mode_combo.addItem("공내역서 직접입력", "manual")
        self.budget_mode_combo.setFixedHeight(28)
        budget_grid.addWidget(QLabel("가실행 기준"), 1, 0)
        budget_grid.addWidget(self.budget_mode_combo, 1, 1)
        self.won_rounding_combo = QComboBox()
        self.won_rounding_combo.addItem("반올림", "round")
        self.won_rounding_combo.addItem("올림", "ceil")
        self.won_rounding_combo.addItem("버림", "floor")
        self.won_rounding_combo.addItem("소수점 유지", "keep")
        self.won_rounding_combo.setFixedHeight(28)
        budget_grid.addWidget(QLabel("원단위 처리"), 2, 0)
        budget_grid.addWidget(self.won_rounding_combo, 2, 1)
        self.manual_budget_supply_name = QLabel("직접입력 공급가")
        self.manual_budget_supply_edit = QLineEdit("0")
        self.manual_budget_supply_edit.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.manual_budget_supply_edit.setFixedHeight(28)
        budget_grid.addWidget(self.manual_budget_supply_name, 3, 0)
        budget_grid.addWidget(self.manual_budget_supply_edit, 3, 1)
        self.budget_label = QLabel("₩ 0")
        self.contract_label = QLabel("₩ 0")
        self.ratio_label = QLabel("-")
        for row, (name, value) in enumerate((("가실행", self.budget_label), ("계약", self.contract_label), ("계약금액 비율", self.ratio_label)), start=4):
            name_label = QLabel(name)
            name_label.setObjectName("summaryName")
            value.setObjectName("fixedMetricValue")
            value.setWordWrap(False)
            value.setMinimumWidth(116)
            value.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            budget_grid.addWidget(name_label, row, 0)
            budget_grid.addWidget(value, row, 1)
        budget_grid.setColumnStretch(1, 1)
        summary_layout.addWidget(budget_card)
        summary_container.addLayout(summary_layout)
        outer.addWidget(summary_group)
        return page

    @staticmethod
    def _configure_edit_table(table: QTableWidget) -> None:
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setAlternatingRowColors(True)
        table.setShowGrid(True)
        table.horizontalHeader().setFixedHeight(34)
        table.setMinimumHeight(54)
        table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

    def _build_statement_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        info, info_layout = build_section_card("1. 내역서 기본정보")
        info_grid = QGridLayout()
        info_grid.setContentsMargins(0, 0, 0, 0)
        info_grid.setHorizontalSpacing(10)
        info_grid.setVerticalSpacing(8)
        self.statement_title_edit = QLineEdit()
        self.statement_vendor_label = QLabel("선정업체 없음")
        self.statement_vendor_label.setObjectName("linkedValueBadge")
        info_grid.addWidget(QLabel("기본 대표 품목명"), 0, 0)
        info_grid.addWidget(self.statement_title_edit, 0, 1)
        info_grid.addWidget(QLabel("선정업체"), 0, 2)
        info_grid.addWidget(self.statement_vendor_label, 0, 3)
        info_grid.setColumnStretch(1, 2)
        info_grid.setColumnStretch(3, 1)
        info_layout.addLayout(info_grid)

        buttons = QHBoxLayout()
        buttons.setSpacing(8)
        self.statement_sync_button = QPushButton("견적대비표에서 불러오기")
        self.statement_sync_button.setObjectName("accentButton")
        self.statement_add_button = QPushButton("＋ 행 추가")
        self.statement_delete_button = QPushButton("선택 행 삭제")
        self.statement_up_button = QPushButton("↑ 위로")
        self.statement_down_button = QPushButton("↓ 아래로")
        self.statement_renumber_button = QPushButton("번호 다시 매기기")
        for button in (self.statement_sync_button, self.statement_add_button, self.statement_delete_button, self.statement_up_button, self.statement_down_button, self.statement_renumber_button):
            buttons.addWidget(button)
        buttons.addStretch(1)
        info_layout.addLayout(buttons)
        self.statement_status_label = QLabel("견적대비표 내용을 불러온 뒤 번호·품명·규격·단위·수량·단가·금액을 직접 편집할 수 있습니다.")
        self.statement_status_label.setObjectName("mutedText")
        self.statement_status_label.setWordWrap(True)
        info_layout.addWidget(self.statement_status_label)
        layout.addWidget(info)

        table_card, table_layout = build_section_card("2. 구매물품 내역")
        self.statement_table = QTableWidget(0, 9)
        self.statement_table.setHorizontalHeaderLabels(["대표 품목명", "번호", "품명", "규격", "단위", "수량", "단가", "금액", "비고"])
        self.statement_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.statement_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.statement_table.setAlternatingRowColors(True)
        self.statement_table.verticalHeader().setDefaultSectionSize(32)
        self.statement_table.horizontalHeader().setSectionsMovable(True)
        self.statement_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.statement_table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        for column, width in enumerate((190, 70, 220, 170, 75, 110, 125, 145, 220)):
            self.statement_table.setColumnWidth(column, width)
        table_layout.addWidget(self.statement_table, 1)
        layout.addWidget(table_card, 1)

        total_group, total_container = build_section_card("3. 내역서 합계")
        total_layout = QGridLayout()
        total_layout.setContentsMargins(0, 0, 0, 0)
        total_layout.setHorizontalSpacing(12)
        total_layout.setVerticalSpacing(8)
        self.statement_supply_label = QLabel("₩ 0")
        self.statement_vat_label = QLabel("₩ 0")
        self.statement_total_label = QLabel("₩ 0")
        self.statement_budget_label = QLabel("₩ 0")
        for label in (self.statement_supply_label, self.statement_vat_label, self.statement_total_label, self.statement_budget_label):
            label.setObjectName("fixedMetricValue")
            label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        total_layout.addWidget(QLabel("내역서 공급가"), 0, 0)
        total_layout.addWidget(self.statement_supply_label, 0, 1)
        total_layout.addWidget(QLabel("부가세"), 0, 2)
        total_layout.addWidget(self.statement_vat_label, 0, 3)
        total_layout.addWidget(QLabel("부가세 포함 합계"), 1, 0)
        total_layout.addWidget(self.statement_total_label, 1, 1)
        total_layout.addWidget(QLabel("가실행 합계"), 1, 2)
        total_layout.addWidget(self.statement_budget_label, 1, 3)
        total_layout.setColumnStretch(1, 1)
        total_layout.setColumnStretch(3, 1)
        total_container.addLayout(total_layout)
        layout.addWidget(total_group)
        return page

    def _build_purchase_tab(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setObjectName("purchaseScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        container = QWidget()
        container.setObjectName("purchasePage")
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 8, 14)
        layout.setSpacing(10)
        scroll.setWidget(container)

        header, header_layout = build_section_card("1. 문서 기본정보")
        header_grid = QGridLayout()
        header_grid.setContentsMargins(0, 0, 0, 0)
        header_grid.setHorizontalSpacing(10)
        header_grid.setVerticalSpacing(8)

        class_widget = QWidget()
        class_grid = QGridLayout(class_widget)
        class_grid.setContentsMargins(0, 0, 0, 0)
        class_grid.setHorizontalSpacing(8)
        self.year_spin = QSpinBox()
        self.year_spin.setRange(2020, 2100)
        self.year_spin.setValue(date.today().year)
        self.year_spin.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.PlusMinus)
        self.class_site_short_edit = QLineEdit()
        self.sequence_spin = QSpinBox()
        self.sequence_spin.setRange(1, 9999)
        self.sequence_spin.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.sequence_spin.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.sequence_minus_button = QPushButton("−")
        self.sequence_plus_button = QPushButton("+")
        for button in (self.sequence_minus_button, self.sequence_plus_button):
            button.setFixedSize(40, 34)
            button.setAutoRepeat(True)
            button.setAutoRepeatDelay(300)
            button.setAutoRepeatInterval(80)
        self.manual_class_check = QCheckBox("직접 입력")
        self.classification_edit = QLineEdit()
        self.classification_edit.setReadOnly(True)
        class_grid.addWidget(QLabel("연도"), 0, 0)
        class_grid.addWidget(self.year_spin, 0, 1)
        class_grid.addWidget(QLabel("현장 약칭"), 0, 2)
        class_grid.addWidget(self.class_site_short_edit, 0, 3)
        class_grid.addWidget(QLabel("호수"), 0, 4)
        class_grid.addWidget(self.sequence_minus_button, 0, 5)
        class_grid.addWidget(self.sequence_spin, 0, 6)
        class_grid.addWidget(self.sequence_plus_button, 0, 7)
        class_grid.addWidget(self.manual_class_check, 0, 8)
        class_grid.addWidget(self.classification_edit, 1, 0, 1, 9)
        class_grid.setColumnStretch(3, 1)

        self.department_edit = QLineEdit("자 재 부")
        self.draft_date_edit = QDateEdit(QDate.currentDate())
        self.draft_date_edit.setCalendarPopup(True)
        self.draft_date_edit.setDisplayFormat("yyyy.MM.dd")
        self.effective_date_edit = QLineEdit("결재후 즉시")
        self.drafter_edit = QLineEdit()
        self.approval_note_edit = QLineEdit("전결규정  제   조   항에 의한 전결사항임.")
        header_grid.addWidget(QLabel("분류번호"), 0, 0)
        header_grid.addWidget(class_widget, 0, 1, 1, 5)
        header_grid.addWidget(QLabel("기안부서"), 1, 0)
        header_grid.addWidget(self.department_edit, 1, 1)
        header_grid.addWidget(QLabel("기안일자"), 1, 2)
        header_grid.addWidget(self.draft_date_edit, 1, 3)
        header_grid.addWidget(QLabel("시행일자"), 1, 4)
        header_grid.addWidget(self.effective_date_edit, 1, 5)
        header_grid.addWidget(QLabel("기안자"), 2, 0)
        header_grid.addWidget(self.drafter_edit, 2, 1)
        header_grid.addWidget(QLabel("전결 문구"), 2, 2)
        header_grid.addWidget(self.approval_note_edit, 2, 3, 1, 3)
        header_grid.setColumnStretch(1, 1)
        header_grid.setColumnStretch(3, 1)
        header_grid.setColumnStretch(5, 2)
        header_layout.addLayout(header_grid)
        layout.addWidget(header)

        upper = QHBoxLayout()
        upper.setSpacing(10)
        content, content_layout = build_section_card("2. 품의 내용")
        content_grid = QGridLayout()
        content_grid.setContentsMargins(0, 0, 0, 0)
        content_grid.setHorizontalSpacing(10)
        content_grid.setVerticalSpacing(8)
        self.purchase_title_edit = QLineEdit()
        self.purchase_item_edit = QLineEdit()
        self.purchase_site_edit = QLineEdit()
        content_grid.addWidget(QLabel("제목"), 0, 0)
        content_grid.addWidget(self.purchase_title_edit, 0, 1)
        content_grid.addWidget(QLabel("현장명"), 1, 0)
        content_grid.addWidget(self.purchase_site_edit, 1, 1)
        content_grid.addWidget(QLabel("품명 및 규격"), 2, 0)
        content_grid.addWidget(self.purchase_item_edit, 2, 1)
        content_grid.setColumnStretch(1, 1)
        content_layout.addLayout(content_grid)
        upper.addWidget(content, 3)

        money, money_layout = build_section_card("3. 금액 정보")
        money_grid = QGridLayout()
        money_grid.setContentsMargins(0, 0, 0, 0)
        money_grid.setHorizontalSpacing(8)
        money_grid.setVerticalSpacing(7)
        self.budget_edit = QLineEdit("₩ 0")
        self.contract_edit = QLineEdit("₩ 0")
        self.purchase_ratio_edit = QLineEdit("-")
        for edit in (self.budget_edit, self.contract_edit, self.purchase_ratio_edit):
            edit.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            edit.setObjectName("moneyValueEdit")
        money_grid.addWidget(QLabel("가실행 금액(부가세 포함)"), 0, 0)
        money_grid.addWidget(self.budget_edit, 1, 0)
        money_grid.addWidget(QLabel("계약금액(부가세 포함)"), 2, 0)
        money_grid.addWidget(self.contract_edit, 3, 0)
        money_grid.addWidget(QLabel("계약금액 비율"), 4, 0)
        money_grid.addWidget(self.purchase_ratio_edit, 5, 0)
        money_grid.setRowStretch(6, 1)
        money_layout.addLayout(money_grid)
        upper.addWidget(money, 2)
        layout.addLayout(upper)

        contract, contract_layout = build_section_card("4. 기간·거래처·조건")
        contract_grid = QGridLayout()
        contract_grid.setContentsMargins(0, 0, 0, 0)
        contract_grid.setHorizontalSpacing(10)
        contract_grid.setVerticalSpacing(8)
        self.period_kind_combo = QComboBox()
        self.period_kind_combo.addItems(["계약기간", "임차기간", "납품일자"])
        self.period_edit = QLineEdit()
        period_widget = QWidget()
        period_layout = QHBoxLayout(period_widget)
        period_layout.setContentsMargins(0, 0, 0, 0)
        period_layout.setSpacing(8)
        period_layout.addWidget(self.period_kind_combo)
        period_layout.addWidget(self.period_edit, 1)
        self.vendor_edit = QLineEdit()
        self.phone_edit = QLineEdit()
        self.attachment_edit = QLineEdit("구매물품내역서")
        self.payment_edit = QLineEdit("기성결제 현금 (100%) 지급")
        contract_grid.addWidget(QLabel("기간"), 0, 0)
        contract_grid.addWidget(period_widget, 0, 1, 1, 3)
        contract_grid.addWidget(QLabel("거래처"), 1, 0)
        contract_grid.addWidget(self.vendor_edit, 1, 1)
        contract_grid.addWidget(QLabel("전화번호"), 1, 2)
        contract_grid.addWidget(self.phone_edit, 1, 3)
        contract_grid.addWidget(QLabel("첨부서류"), 2, 0)
        contract_grid.addWidget(self.attachment_edit, 2, 1)
        contract_grid.addWidget(QLabel("지불조건"), 2, 2)
        contract_grid.addWidget(self.payment_edit, 2, 3)
        contract_grid.setColumnStretch(1, 1)
        contract_grid.setColumnStretch(3, 1)
        contract_layout.addLayout(contract_grid)
        layout.addWidget(contract)

        body, body_layout = build_section_card("5. 본문 및 추가 문구")
        body_grid = QGridLayout()
        body_grid.setContentsMargins(0, 0, 0, 0)
        body_grid.setHorizontalSpacing(10)
        body_grid.setVerticalSpacing(8)
        self.body_edit = QPlainTextEdit()
        self.body_edit.setMinimumHeight(108)
        self.note_edit = QPlainTextEdit()
        self.note_edit.setMinimumHeight(70)
        self.note_edit.setPlaceholderText("입력한 내용이 품의 문구 바로 다음 행에 같은 서식으로 출력됩니다.")
        body_grid.addWidget(QLabel("품의 문구"), 0, 0)
        body_grid.addWidget(self.body_edit, 0, 1)
        body_grid.addWidget(QLabel("추가 문구"), 1, 0)
        body_grid.addWidget(self.note_edit, 1, 1)
        body_grid.setColumnStretch(1, 1)
        body_layout.addLayout(body_grid)
        layout.addWidget(body)
        layout.addStretch(1)
        return scroll

    def _connect_signals(self) -> None:
        self.site_preset_combo.currentTextChanged.connect(self.apply_site_preset)
        self.item_preset_combo.currentTextChanged.connect(self.apply_item_preset)
        self.site_save_button.clicked.connect(self.save_site_preset)
        self.site_delete_button.clicked.connect(self.delete_site_preset)
        self.item_save_button.clicked.connect(self.save_item_preset)
        self.item_delete_button.clicked.connect(self.delete_item_preset)

        self.add_vendor_button.clicked.connect(self.add_vendor)
        self.delete_vendor_button.clicked.connect(self.delete_vendor)
        self.vendor_left_button.clicked.connect(lambda: self.move_vendor(-1))
        self.vendor_right_button.clicked.connect(lambda: self.move_vendor(1))
        self.vendor_table.cellChanged.connect(self.on_vendor_changed)
        self.common_delivery_place_edit.editingFinished.connect(
            self.on_common_delivery_place_changed
        )
        self.apply_delivery_place_button.clicked.connect(self.apply_common_delivery_place_to_all)

        self.add_item_button.clicked.connect(self.add_item)
        self.add_group_item_button.clicked.connect(self.add_group_item)
        self.delete_item_button.clicked.connect(self.delete_item)
        self.item_up_button.clicked.connect(lambda: self.move_item(-1))
        self.item_down_button.clicked.connect(lambda: self.move_item(1))
        self.item_table.cellChanged.connect(self.on_item_changed)

        self.statement_sync_button.clicked.connect(self.sync_statement_from_quote)
        self.statement_add_button.clicked.connect(self.add_statement_item)
        self.statement_delete_button.clicked.connect(self.delete_statement_item)
        self.statement_up_button.clicked.connect(lambda: self.move_statement_item(-1))
        self.statement_down_button.clicked.connect(lambda: self.move_statement_item(1))
        self.statement_renumber_button.clicked.connect(self.renumber_statement_items)
        self.statement_table.cellChanged.connect(self.on_statement_changed)
        self.statement_title_edit.textChanged.connect(self.on_statement_title_changed)
        self.selected_vendor_combo.currentIndexChanged.connect(self.on_selected_vendor_changed)
        self.budget_mode_combo.currentIndexChanged.connect(self.on_budget_controls_changed)
        self.won_rounding_combo.currentIndexChanged.connect(self.on_budget_controls_changed)
        self.manual_budget_supply_edit.editingFinished.connect(self.on_budget_controls_changed)
        self.default_layout_button.clicked.connect(self.reset_default_layout)
        self.purchase_site_edit.textEdited.connect(lambda _text: self.mark_purchase_override("site"))
        self.vendor_edit.textEdited.connect(lambda _text: self.mark_purchase_override("vendor"))
        self.phone_edit.textEdited.connect(lambda _text: self.mark_purchase_override("phone"))
        self.budget_edit.textEdited.connect(lambda _text: self.mark_purchase_override("budget"))
        self.contract_edit.textEdited.connect(lambda _text: self.mark_purchase_override("contract"))
        self.purchase_ratio_edit.textEdited.connect(lambda _text: self.mark_purchase_override("ratio"))
        self.payment_edit.textEdited.connect(lambda _text: self.mark_purchase_override("payment"))
        self.tabs.currentChanged.connect(self.on_tab_changed)

        self.preview_sheet_combo.currentIndexChanged.connect(self.on_preview_kind_changed)
        self.preview_sheet_combo.currentIndexChanged.connect(self.schedule_preview_refresh)
        self.preview_follow_check.toggled.connect(self.on_preview_follow_changed)
        self.quote_splitter.splitterMoved.connect(self.save_layout_settings)
        self.main_splitter.splitterMoved.connect(self.save_layout_settings)
        self.right_splitter.splitterMoved.connect(self.save_layout_settings)

        for widget in (
            self.site_name_edit,
            self.site_short_edit,
            self.quote_title_edit,
            self.item_label_edit,
            self.author_edit,
        ):
            widget.textChanged.connect(self.refresh_calculations)
        self.quote_date_edit.dateChanged.connect(self.refresh_calculations)

        for widget in (
            self.classification_edit,
            self.department_edit,
            self.effective_date_edit,
            self.drafter_edit,
            self.approval_note_edit,
            self.purchase_title_edit,
            self.purchase_item_edit,
            self.period_edit,
            self.attachment_edit,
            self.payment_edit,
        ):
            widget.textChanged.connect(self.schedule_preview_refresh)
        self.draft_date_edit.dateChanged.connect(self.schedule_preview_refresh)
        self.period_kind_combo.currentTextChanged.connect(self.schedule_preview_refresh)
        self.body_edit.textChanged.connect(self.schedule_preview_refresh)
        self.note_edit.textChanged.connect(self.schedule_preview_refresh)

        self.year_spin.valueChanged.connect(self.update_classification)
        self.class_site_short_edit.textChanged.connect(self.update_classification)
        self.sequence_spin.valueChanged.connect(self.update_classification)
        self.sequence_minus_button.clicked.connect(lambda: self.sequence_spin.stepBy(-1))
        self.sequence_plus_button.clicked.connect(lambda: self.sequence_spin.stepBy(1))
        self.manual_class_check.toggled.connect(self.toggle_manual_classification)
        self.site_short_edit.textChanged.connect(self.sync_site_short_fields)
        self.class_site_short_edit.textChanged.connect(self.sync_site_short_fields_reverse)

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            * { font-family: 'Pretendard', 'Malgun Gothic', sans-serif; font-size: 12px; color: #172033; }
            QMainWindow, QWidget#appRoot, QWidget#workspace { background: #F6F8FC; }
            QFrame#sidebar { background: #FFFFFF; border-right: 1px solid #E6EBF3; }
            QLabel#brandTitle { font-size: 18px; font-weight: 800; color: #111827; }
            QLabel#brandSubtitle, QLabel#versionLabel { color: #7B879B; font-size: 11px; }
            QFrame#sidebarDivider { color: #E8EDF5; }
            QPushButton#navButton, QPushButton#navActionButton, QPushButton#sidebarFooterButton {
                min-height: 44px; border: 0; border-radius: 12px; padding: 0 14px; text-align: left; background: transparent; color: #334155; font-weight: 600;
            }
            QPushButton#navButton:hover, QPushButton#navActionButton:hover, QPushButton#sidebarFooterButton:hover { background: #F1F5FB; }
            QPushButton#navButton:checked { background: #E8F1FF; color: #1263E5; font-weight: 800; }
            QFrame#topbar { background: #FFFFFF; border-bottom: 1px solid #E7ECF4; }
            QLabel#pageTitle { font-size: 20px; font-weight: 800; color: #111827; }
            QPushButton#topActionButton, QPushButton {
                min-height: 34px; border: 1px solid #D8E1EE; border-radius: 10px; padding: 0 13px; background: #FFFFFF; color: #334155; font-weight: 600;
            }
            QPushButton:hover { background: #F3F7FD; border-color: #B8C8DF; }
            QPushButton#primaryButton { background: #1263E5; border-color: #1263E5; color: white; min-height: 38px; font-weight: 800; }
            QPushButton#primaryButton:hover { background: #0F55C7; }
            QPushButton#accentButton { background: #1263E5; border-color: #1263E5; color: white; }
            QPushButton#accentButton:hover { background: #0F55C7; }
            QFrame#card, QFrame#previewCard, QFrame#overviewCard, QFrame#selectionCard, QFrame#budgetCard, QFrame#sectionCard {
                background: #FFFFFF; border: 1px solid #E4EAF3; border-radius: 14px;
            }
            QGroupBox#cardGroup {
                background: #FFFFFF; border: 1px solid #E4EAF3; border-radius: 14px; margin-top: 12px; padding-top: 8px; font-size: 14px; font-weight: 800; color: #183B70;
            }
            QGroupBox#cardGroup::title { subcontrol-origin: margin; left: 14px; padding: 0 7px; background: #FFFFFF; }
            QLabel#sectionCardTitle { font-size: 14px; font-weight: 800; color: #173B70; padding: 0 0 2px 0; }
            QLabel#cardTitle, QLabel#sectionPageTitle { font-size: 16px; font-weight: 800; color: #173B70; }
            QLabel#subCardTitle { font-size: 13px; font-weight: 800; color: #173B70; }
            QLabel#sectionPageTitle { font-size: 20px; }
            QLabel#mutedText, QLabel#summaryName { color: #6B778C; }
            QLabel#summaryValue, QLabel#metricValue, QLabel#fixedMetricValue { color: #1263E5; font-weight: 800; }
            QLabel#fixedMetricValue { font-size: 12px; min-height: 24px; padding-left: 2px; }
            QLabel#rankBadge, QLabel#linkedValueBadge { background: #EAF7EE; color: #218B4F; border: 1px solid #BFE6CB; border-radius: 9px; padding: 5px 10px; font-weight: 800; }
            QLabel#stepBadge { background: #EEF4FF; color: #1263E5; border: 1px solid #D8E6FF; border-radius: 9px; padding: 4px 9px; font-weight: 800; }
            QFrame#overviewDivider { color: #E6ECF4; }
            QPushButton#secondaryNavButton { min-height: 34px; background: #FFFFFF; color: #334155; }
            QPushButton#primaryNavButton { min-height: 34px; background: #1263E5; border-color: #1263E5; color: #FFFFFF; font-weight: 800; }
            QPushButton#primaryNavButton:hover { background: #0F55C7; }
            QLabel#calculationNote { background: #EFFAF3; color: #2C8A53; border-radius: 9px; padding: 9px; font-size: 11px; }
            QLineEdit, QComboBox, QDateEdit, QSpinBox, QPlainTextEdit {
                min-height: 34px; border: 1px solid #D9E2EF; border-radius: 9px; padding: 0 10px; background: #FFFFFF; selection-background-color: #DDEBFF; selection-color: #172033;
            }
            QLineEdit:focus, QComboBox:focus, QDateEdit:focus, QSpinBox:focus, QPlainTextEdit:focus { border: 1px solid #2F7AF6; }
            QLineEdit#moneyValueEdit { color: #1263E5; font-weight: 800; background: #F8FBFF; }
            QScrollArea#purchaseScroll, QWidget#purchasePage { background: transparent; border: 0; }
            QTableWidget { font-size: 10pt; background: #FFFFFF; alternate-background-color: #FAFBFD; border: 1px solid #E1E7F0; border-radius: 9px; gridline-color: #E4E9F1; selection-background-color: #E7F0FF; selection-color: #172033; outline: 0; }
            QTableWidget::item { padding: 4px 7px; }
            QTableWidget::item:selected { background: #E7F0FF; color: #172033; border: 1px solid #2F7AF6; }
            QHeaderView::section { background: #F3F6FB; color: #20304A; padding: 7px; border: 0; border-right: 1px solid #E1E7F0; border-bottom: 1px solid #E1E7F0; font-weight: 800; }
            QSplitter#mainSplitter::handle, QSplitter#quoteSplitter::handle { background: #F6F8FC; }
            QSplitter#mainSplitter::handle:hover, QSplitter#quoteSplitter::handle:hover { background: #EDF3FA; }
            QTextBrowser { background: #EEF2F7; border: 1px solid #DDE4EE; border-radius: 10px; }
            QPushButton#previewTabButton { border: 0; border-bottom: 2px solid transparent; border-radius: 0; background: transparent; color: #66758C; min-height: 34px; }
            QPushButton#previewTabButton:checked { color: #1263E5; border-bottom: 2px solid #1263E5; font-weight: 800; }
            QScrollBar:vertical { width: 10px; background: transparent; }
            QScrollBar::handle:vertical { background: #CBD5E1; min-height: 30px; border-radius: 5px; }
            QLabel#previewHint, QLabel#splitterHint { color: #6B778C; font-size: 11px; }
            """
        )

    def _load_presets(self) -> None:
        previous_loading = self._loading
        self._loading = True
        try:
            self.site_preset_combo.clear()
            self.site_preset_combo.addItems(self.store.names("sites"))
            self.item_preset_combo.clear()
            self.item_preset_combo.addItems(self.store.names("items"))
        finally:
            self._loading = previous_loading

    def _reset_with_defaults(self) -> None:
        self._loading = True
        try:
            self.data = ProjectData()
            self.vendor_table.setRowCount(0)
            self.item_table.setRowCount(0)
            self.statement_table.setRowCount(0)
            self.item_notes = []
            self._common_delivery_place = ""
            self._statement_source_signature = ""
            self._purchase_overrides.clear()
            self.budget_mode_combo.setCurrentIndex(0)
            self.won_rounding_combo.setCurrentIndex(0)
            self.manual_budget_supply_edit.setText("0")
            self.common_delivery_place_edit.clear()
            last_site = self.store.setting("last_site", "")
            last_item = self.store.setting("last_item", "")
            if last_site and self.site_preset_combo.findText(last_site) >= 0:
                self.site_preset_combo.setCurrentText(last_site)
            elif self.site_preset_combo.count():
                self.site_preset_combo.setCurrentIndex(0)
            if last_item and self.item_preset_combo.findText(last_item) >= 0:
                self.item_preset_combo.setCurrentText(last_item)
            elif self.item_preset_combo.count():
                self.item_preset_combo.setCurrentIndex(0)
        finally:
            self._loading = False
        self.apply_site_preset(self.site_preset_combo.currentText())
        self.apply_item_preset(self.item_preset_combo.currentText())
        if self.vendor_table.rowCount() == 0:
            for index in range(3):
                self.add_vendor(
                    Vendor(
                        name=f"업체 {index + 1}",
                        payment="기성결제 현금",
                        delivery_place=self.common_delivery_place_edit.text().strip(),
                    ),
                    refresh=False,
                )
        if self.item_table.rowCount() == 0:
            self.rebuild_item_table(
                [
                    QuoteItem(
                        unit="EA",
                        quantity=1,
                        unit_prices=[0] * self.vendor_table.rowCount(),
                        group_title=self.statement_title_edit.text().strip(),
                        group_sequence="1",
                    )
                ]
            )
        self.update_classification()
        self.restore_layout_settings()
        self.tabs.setCurrentIndex(0)
        quote_preview_index = self.preview_sheet_combo.findData("quote")
        if quote_preview_index >= 0:
            self.preview_sheet_combo.setCurrentIndex(quote_preview_index)
        self.refresh_calculations()
        self.sync_statement_from_quote(confirm=False)
        for table in (self.vendor_table, self.item_table, self.summary_table, self.statement_table):
            table.clearSelection()
            table.setCurrentCell(-1, -1)

    def apply_site_preset(self, name: str) -> None:
        if self._loading or not name:
            return
        preset = self.store.find("sites", name)
        if not preset:
            return
        previous_loading = self._loading
        self._loading = True
        try:
            self.site_name_edit.setText(str(preset.get("full_name", "")))
            self.site_short_edit.setText(str(preset.get("short_name", "")))
            self.author_edit.setText(str(preset.get("drafter", "")))
            self.drafter_edit.setText(str(preset.get("drafter", "")))
            self.department_edit.setText(str(preset.get("department", "자 재 부")))
            self.effective_date_edit.setText(str(preset.get("effective_date", "결재후 즉시")))
            self.attachment_edit.setText(str(preset.get("attachment", "구매물품내역서")))
            base_payment = normalize_quote_payment(str(preset.get("payment", "기성결제 현금")))
            self.payment_edit.setText(purchase_payment_text(base_payment))
            delivery_place = str(preset.get("delivery_place", ""))
            self._common_delivery_place = delivery_place
            self.common_delivery_place_edit.setText(delivery_place)
            self.purchase_site_edit.setText(self.site_name_edit.text())
            if self._body_template:
                self.body_edit.setPlainText(self._body_template.format(site=self.site_name_edit.text()))
            self.store.set_setting("last_site", name)
        finally:
            self._loading = previous_loading
        self.update_classification()
        self.refresh_calculations()

    def apply_item_preset(self, name: str) -> None:
        if self._loading or not name:
            return
        preset = self.store.find("items", name)
        if not preset:
            return
        previous_loading = self._loading
        self._loading = True
        try:
            self.item_label_edit.setText(str(preset.get("sheet_label", name)))
            title = str(preset.get("title", name))
            statement_title = str(preset.get("statement_title", title))
            self.quote_title_edit.setText(title)
            self.purchase_title_edit.setText(title)
            self.purchase_item_edit.setText(str(preset.get("purchase_item_name", title)))
            self.statement_title_edit.setText(statement_title)
            self.period_kind_combo.setCurrentText(str(preset.get("period_kind", "임차기간")))
            self.period_edit.setText(str(preset.get("period", "")))
            self.attachment_edit.setText(str(preset.get("attachment", "구매물품내역서")))
            base_payment = normalize_quote_payment(str(preset.get("payment", "기성결제 현금")))
            self.payment_edit.setText(purchase_payment_text(base_payment))
            self._body_template = str(
                preset.get(
                    "body_template",
                    "{site} 현장의 업체를 선정하여 품의하오니 결재하여 주시기 바랍니다.",
                )
            )
            self.body_edit.setPlainText(self._body_template.format(site=self.site_name_edit.text()))
            self.note_edit.setPlainText(str(preset.get("note", "")))
            items = [QuoteItem.from_dict(value) for value in preset.get("items", [])]
            if items and not any(item.group_title.strip() for item in items):
                items[0].group_title = statement_title
                items[0].group_sequence = "1"
            self.item_notes = [item.note for item in items]
            self.rebuild_item_table(items)
            self.data.statement_items = []
            self._statement_source_signature = ""
            if hasattr(self, "statement_table"):
                self.rebuild_statement_table([])
            self.store.set_setting("last_item", name)
        finally:
            self._loading = previous_loading
        self.refresh_calculations()

    def save_site_preset(self) -> None:
        dialog = PresetNameDialog("현장 프리셋 저장", self.site_preset_combo.currentText(), self)
        if dialog.exec() != QDialog.DialogCode.Accepted or not dialog.name:
            return
        preset = {
            "name": dialog.name,
            "short_name": self.site_short_edit.text().strip(),
            "full_name": self.site_name_edit.text().strip(),
            "drafter": self.drafter_edit.text().strip(),
            "department": self.department_edit.text().strip(),
            "effective_date": self.effective_date_edit.text().strip(),
            "attachment": self.attachment_edit.text().strip(),
            "payment": normalize_quote_payment(self.payment_edit.text()),
            "delivery_place": self.common_delivery_place_edit.text().strip(),
        }
        self.store.upsert("sites", preset)
        self._load_presets()
        self.site_preset_combo.setCurrentText(dialog.name)

    def save_item_preset(self) -> None:
        dialog = PresetNameDialog("품목 프리셋 저장", self.item_preset_combo.currentText(), self)
        if dialog.exec() != QDialog.DialogCode.Accepted or not dialog.name:
            return
        items = [item.to_dict() for item in self.collect_items()]
        preset = {
            "name": dialog.name,
            "sheet_label": self.item_label_edit.text().strip(),
            "title": self.purchase_title_edit.text().strip(),
            "purchase_item_name": self.purchase_item_edit.text().strip(),
            "statement_title": self.statement_title_edit.text().strip(),
            "period_kind": self.period_kind_combo.currentText(),
            "period": self.period_edit.text().strip(),
            "attachment": self.attachment_edit.text().strip(),
            "payment": normalize_quote_payment(self.payment_edit.text()),
            "body_template": self._body_template or self.body_edit.toPlainText(),
            "note": self.note_edit.toPlainText(),
            "items": items,
        }
        self.store.upsert("items", preset)
        self._load_presets()
        self.item_preset_combo.setCurrentText(dialog.name)

    def delete_site_preset(self) -> None:
        self._delete_preset("sites", self.site_preset_combo.currentText())

    def delete_item_preset(self) -> None:
        self._delete_preset("items", self.item_preset_combo.currentText())

    def _delete_preset(self, kind: str, name: str) -> None:
        if not name:
            return
        if QMessageBox.question(self, "프리셋 삭제", f"'{name}' 프리셋을 삭제할까요?") != QMessageBox.StandardButton.Yes:
            return
        self.store.delete(kind, name)  # type: ignore[arg-type]
        self._load_presets()

    # ------------------------------------------------------------------
    # Vendor / item table handling
    # ------------------------------------------------------------------
    def add_vendor(self, vendor: Vendor | None = None, refresh: bool = True) -> None:
        common_place = self.common_delivery_place_edit.text().strip()
        vendor = vendor or Vendor(
            name=f"업체 {self.vendor_table.rowCount() + 1}",
            payment="기성결제 현금",
            delivery_place=common_place,
        )
        vendor.payment = normalize_quote_payment(vendor.payment)
        if not vendor.delivery_place:
            vendor.delivery_place = common_place
        items = self.collect_items()
        previous_loading = self._loading
        self._loading = True
        try:
            row = self.vendor_table.rowCount()
            self.vendor_table.insertRow(row)
            values = [
                vendor.name,
                vendor.phone,
                vendor.manager,
                vendor.payment,
                vendor.delivery_place,
                vendor.delivery_date,
            ]
            for column, value in enumerate(values):
                self.vendor_table.setItem(row, column, text_item(value))
            self.vendor_table.setItem(row, 6, check_item(vendor.submitted))
            self.vendor_table.setItem(row, 7, check_item(vendor.include_in_average))
        finally:
            self._loading = previous_loading
        for item in items:
            item.unit_prices.append(0)
        self.rebuild_item_table(items)
        if refresh:
            self.refresh_calculations()

    def delete_vendor(self) -> None:
        row = self.vendor_table.currentRow()
        if row < 0:
            return
        items = self.collect_items()
        for item in items:
            if row < len(item.unit_prices):
                item.unit_prices.pop(row)
        selected = self.data.selected_vendor_index
        if selected == row:
            self.data.selected_vendor_index = None
        elif selected is not None and selected > row:
            self.data.selected_vendor_index = selected - 1
        previous_loading = self._loading
        self._loading = True
        try:
            self.vendor_table.removeRow(row)
        finally:
            self._loading = previous_loading
        self.rebuild_item_table(items)
        self.refresh_calculations()

    def move_vendor(self, direction: int) -> None:
        row = self.vendor_table.currentRow()
        target = row + direction
        if row < 0 or target < 0 or target >= self.vendor_table.rowCount():
            return
        vendors = self.collect_vendors()
        items = self.collect_items()
        vendors[row], vendors[target] = vendors[target], vendors[row]
        for item in items:
            item.unit_prices[row], item.unit_prices[target] = item.unit_prices[target], item.unit_prices[row]
        selected = self.data.selected_vendor_index
        if selected == row:
            selected = target
        elif selected == target:
            selected = row
        self.data.selected_vendor_index = selected
        self.fill_vendor_table(vendors)
        self.rebuild_item_table(items)
        self.vendor_table.setCurrentCell(target, 0)
        self.refresh_calculations()

    def fill_vendor_table(self, vendors: list[Vendor]) -> None:
        previous_loading = self._loading
        self._loading = True
        try:
            self.vendor_table.setRowCount(0)
            for vendor in vendors:
                vendor.payment = normalize_quote_payment(vendor.payment)
                row = self.vendor_table.rowCount()
                self.vendor_table.insertRow(row)
                values = [
                    vendor.name,
                    vendor.phone,
                    vendor.manager,
                    vendor.payment,
                    vendor.delivery_place,
                    vendor.delivery_date,
                ]
                for column, value in enumerate(values):
                    self.vendor_table.setItem(row, column, text_item(value))
                self.vendor_table.setItem(row, 6, check_item(vendor.submitted))
                self.vendor_table.setItem(row, 7, check_item(vendor.include_in_average))
            common = self.data.common_delivery_place if self.data else ""
            if not common and vendors:
                common = vendors[0].delivery_place
            self._common_delivery_place = common
            self.common_delivery_place_edit.setText(common)
        finally:
            self._loading = previous_loading

    def collect_vendors(self) -> list[Vendor]:
        vendors: list[Vendor] = []
        for row in range(self.vendor_table.rowCount()):
            submitted_item = self.vendor_table.item(row, 6)
            average_item = self.vendor_table.item(row, 7)
            vendors.append(
                Vendor(
                    name=table_text(self.vendor_table, row, 0),
                    phone=table_text(self.vendor_table, row, 1),
                    manager=table_text(self.vendor_table, row, 2),
                    payment=normalize_quote_payment(table_text(self.vendor_table, row, 3)),
                    delivery_place=table_text(self.vendor_table, row, 4),
                    delivery_date=table_text(self.vendor_table, row, 5),
                    submitted=bool(submitted_item and submitted_item.checkState() == Qt.CheckState.Checked),
                    include_in_average=bool(average_item and average_item.checkState() == Qt.CheckState.Checked),
                )
            )
        return vendors

    def on_common_delivery_place_changed(self) -> None:
        new_value = self.common_delivery_place_edit.text().strip()
        old_value = self._common_delivery_place
        previous_loading = self._loading
        self._loading = True
        try:
            for row in range(self.vendor_table.rowCount()):
                current = table_text(self.vendor_table, row, 4)
                if not current or current == old_value:
                    self.vendor_table.setItem(row, 4, text_item(new_value))
        finally:
            self._loading = previous_loading
        self._common_delivery_place = new_value
        self.refresh_calculations()

    def apply_common_delivery_place_to_all(self) -> None:
        value = self.common_delivery_place_edit.text().strip()
        previous_loading = self._loading
        self._loading = True
        try:
            for row in range(self.vendor_table.rowCount()):
                self.vendor_table.setItem(row, 4, text_item(value))
        finally:
            self._loading = previous_loading
        self._common_delivery_place = value
        self.refresh_calculations()

    def add_item(self, item: QuoteItem | None = None) -> None:
        items = self.collect_items()
        items.append(
            item
            or QuoteItem(unit="EA", quantity=1, unit_prices=[0] * self.vendor_table.rowCount())
        )
        self.item_notes.append(item.note if item else "")
        self.rebuild_item_table(items)
        self.item_table.setCurrentCell(len(items) - 1, 2)
        self.refresh_calculations()

    def add_group_item(self) -> None:
        items = self.collect_items()
        next_sequence = 1 + sum(1 for item in items if item.group_title.strip())
        items.append(
            QuoteItem(
                unit="EA",
                quantity=1,
                unit_prices=[0] * self.vendor_table.rowCount(),
                group_title=self.statement_title_edit.text().strip() or self.quote_title_edit.text().strip(),
                group_sequence=str(next_sequence),
            )
        )
        self.item_notes.append("")
        self.rebuild_item_table(items)
        self.item_table.setCurrentCell(len(items) - 1, 2)
        self.refresh_calculations()

    def delete_item(self) -> None:
        row = self.item_table.currentRow()
        if row < 0:
            return
        items = self.collect_items()
        if row < len(items):
            items.pop(row)
        if row < len(self.item_notes):
            self.item_notes.pop(row)
        self.rebuild_item_table(items)
        self.refresh_calculations()

    def move_item(self, direction: int) -> None:
        row = self.item_table.currentRow()
        target = row + direction
        items = self.collect_items()
        if row < 0 or target < 0 or target >= len(items):
            return
        items[row], items[target] = items[target], items[row]
        if row < len(self.item_notes) and target < len(self.item_notes):
            self.item_notes[row], self.item_notes[target] = self.item_notes[target], self.item_notes[row]
        self.rebuild_item_table(items)
        self.item_table.setCurrentCell(target, 2)
        self.refresh_calculations()

    def rebuild_item_table(self, items: list[QuoteItem] | None = None) -> None:
        if items is None:
            items = self.collect_items()
        vendor_names = [vendor.name or f"업체 {i + 1}" for i, vendor in enumerate(self.collect_vendors())]
        previous_loading = self._loading
        self._loading = True
        try:
            self.item_table.setColumnCount(6 + len(vendor_names))
            self.item_table.setHorizontalHeaderLabels(
                ["순서", "대표 품목명", "품명", "규격", "단위", "수량"]
                + [f"{name}\n단가" for name in vendor_names]
            )
            self.item_table.setRowCount(len(items))
            for row, item in enumerate(items):
                item.normalize_prices(len(vendor_names))
                self.item_table.setItem(row, 0, text_item(item.group_sequence))
                self.item_table.setItem(row, 1, text_item(item.group_title, bold=bool(item.group_title.strip())))
                self.item_table.setItem(row, 2, text_item(item.name))
                self.item_table.setItem(row, 3, text_item(item.spec))
                self.item_table.setItem(row, 4, text_item(item.unit))
                self.item_table.setItem(row, 5, text_item(decimal_text(item.quantity), align_right=True))
                for vendor_index, price in enumerate(item.unit_prices):
                    self.item_table.setItem(
                        row, 6 + vendor_index, text_item(format_money(price), align_right=True)
                    )
            while len(self.item_notes) < len(items):
                self.item_notes.append("")
            self.item_notes = self.item_notes[: len(items)]
            header = self.item_table.horizontalHeader()
            header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
            header.setTextElideMode(Qt.TextElideMode.ElideNone)
            # 업체명 + "단가" 두 줄이 모두 보이도록 헤더 높이와 업체별 열 폭을
            # 실제 업체명 길이에 맞춰 확보한다. 화면이 좁을 때는 열을 줄이지 않고
            # 가로 스크롤을 사용해 업체명이 잘리지 않게 한다.
            header.setFixedHeight(52)
            base_widths = (70, 190, 220, 170, 75, 110)
            for column, width in enumerate(base_widths):
                self.item_table.setColumnWidth(column, width)
            metrics = header.fontMetrics()
            for vendor_index, name in enumerate(vendor_names):
                column = 6 + vendor_index
                full_header = f"{name}\n단가"
                header_item = self.item_table.horizontalHeaderItem(column)
                if header_item is not None:
                    header_item.setToolTip(full_header)
                vendor_width = max(220, min(420, metrics.horizontalAdvance(name) + 104))
                self.item_table.setColumnWidth(column, vendor_width)
        finally:
            self._loading = previous_loading

    def collect_items(self) -> list[QuoteItem]:
        items: list[QuoteItem] = []
        vendor_count = self.vendor_table.rowCount()
        for row in range(self.item_table.rowCount()):
            try:
                quantity = parse_decimal(table_text(self.item_table, row, 5))
            except Exception:
                quantity = parse_decimal(0)
            prices = [
                parse_money(table_text(self.item_table, row, 6 + index))
                for index in range(vendor_count)
            ]
            note = self.item_notes[row] if row < len(self.item_notes) else ""
            items.append(
                QuoteItem(
                    name=table_text(self.item_table, row, 2),
                    spec=table_text(self.item_table, row, 3),
                    unit=table_text(self.item_table, row, 4),
                    quantity=quantity,
                    unit_prices=prices,
                    note=note,
                    group_title=table_text(self.item_table, row, 1),
                    group_sequence=table_text(self.item_table, row, 0),
                )
            )
        return items

    def on_vendor_changed(self, row: int, column: int) -> None:
        if self._loading:
            return
        if column == 3:
            item = self.vendor_table.item(row, column)
            if item:
                normalized = normalize_quote_payment(item.text())
                previous_loading = self._loading
                self._loading = True
                try:
                    item.setText(normalized)
                finally:
                    self._loading = previous_loading
        if column == 4 and row == 0:
            new_value = table_text(self.vendor_table, 0, 4)
            old_value = self._common_delivery_place
            previous_loading = self._loading
            self._loading = True
            try:
                for other_row in range(1, self.vendor_table.rowCount()):
                    current = table_text(self.vendor_table, other_row, 4)
                    if not current or current == old_value:
                        self.vendor_table.setItem(other_row, 4, text_item(new_value))
                self.common_delivery_place_edit.setText(new_value)
            finally:
                self._loading = previous_loading
            self._common_delivery_place = new_value
        items = self.collect_items()
        if column == 0:
            self.rebuild_item_table(items)
        self.refresh_calculations()

    def on_item_changed(self, row: int, column: int) -> None:
        if self._loading:
            return
        if column >= 6:
            item = self.item_table.item(row, column)
            if item:
                previous_loading = self._loading
                self._loading = True
                try:
                    item.setText(format_money(item.text()))
                finally:
                    self._loading = previous_loading
        self.refresh_calculations()

    def on_budget_controls_changed(self, *_args: object) -> None:
        if self._loading:
            return
        manual = str(self.budget_mode_combo.currentData() or "average") == "manual"
        self.manual_budget_supply_name.setVisible(manual)
        self.manual_budget_supply_edit.setVisible(manual)
        self.manual_budget_supply_edit.setEnabled(manual)
        if manual and not self.manual_budget_supply_edit.hasFocus():
            self.manual_budget_supply_edit.setText(
                format_money(parse_money(self.manual_budget_supply_edit.text()))
            )
        self.refresh_calculations()

    def mark_purchase_override(self, key: str) -> None:
        if self._loading:
            return
        self._purchase_overrides.add(key)
        self.schedule_preview_refresh()

    def reset_default_layout(self) -> None:
        self.quote_splitter.setSizes([280, 300])
        self.main_splitter.setSizes([900, 520])
        self.right_splitter.setSizes([700, 184])
        self.save_layout_settings()

    # ------------------------------------------------------------------
    # Calculation and synchronization
    # ------------------------------------------------------------------
    def collect_project_data(self) -> ProjectData:
        current = self.data
        current.site_name = self.site_name_edit.text().strip()
        current.site_short = self.site_short_edit.text().strip()
        current.item_label = self.item_label_edit.text().strip()
        current.quote_title = self.quote_title_edit.text().strip()
        current.author = self.author_edit.text().strip()
        current.quote_date = self.quote_date_edit.date().toPython()
        current.common_delivery_place = self.common_delivery_place_edit.text().strip()
        current.vendors = self.collect_vendors()
        current.items = self.collect_items()
        current.budget_mode = str(self.budget_mode_combo.currentData() or "average")
        current.manual_budget_supply = parse_money(self.manual_budget_supply_edit.text())
        current.won_rounding = str(self.won_rounding_combo.currentData() or "round")
        if hasattr(self, "statement_table"):
            current.statement_items = self.collect_statement_items()
        current.classification = self.classification_edit.text().strip()
        current.department = self.department_edit.text().strip()
        current.draft_date = self.draft_date_edit.date().toPython()
        current.effective_date = self.effective_date_edit.text().strip()
        current.drafter = self.drafter_edit.text().strip()
        current.approval_note = self.approval_note_edit.text().strip()
        current.purchase_title = self.purchase_title_edit.text().strip()
        current.purchase_item_name = self.purchase_item_edit.text().strip()
        current.period_kind = self.period_kind_combo.currentText()
        current.period = self.period_edit.text().strip()
        current.attachment = self.attachment_edit.text().strip()
        current.payment = purchase_payment_text(self.payment_edit.text())
        current.body_text = self.body_edit.toPlainText().strip()
        current.note = self.note_edit.toPlainText().strip()
        current.statement_title = self.statement_title_edit.text().strip()
        current.purchase_override_fields = sorted(self._purchase_overrides)
        current.purchase_site_override = self.purchase_site_edit.text().strip()
        current.purchase_vendor_override = self.vendor_edit.text().strip()
        current.purchase_phone_override = self.phone_edit.text().strip()
        current.purchase_budget_override = parse_money(self.budget_edit.text())
        current.purchase_contract_override = parse_money(self.contract_edit.text())
        current.purchase_ratio_override = self.purchase_ratio_edit.text().strip()
        current.normalize()
        return current

    def _quote_signature(self, data: ProjectData) -> str:
        payload = {
            "winner": data.winner_index,
            "items": [item.to_dict() for item in data.items],
        }
        return json.dumps(payload, ensure_ascii=False, sort_keys=True)

    def refresh_calculations(self) -> None:
        if self._loading:
            return
        data = self.collect_project_data()
        manual_budget = data.budget_mode == "manual"
        self.manual_budget_supply_name.setVisible(manual_budget)
        self.manual_budget_supply_edit.setVisible(manual_budget)
        self.manual_budget_supply_edit.setEnabled(manual_budget)
        if not self.manual_budget_supply_edit.hasFocus():
            shown_supply = data.manual_budget_supply if manual_budget else data.budget_supply
            self.manual_budget_supply_edit.setText(format_money(shown_supply))
        previous_loading = self._loading
        self._loading = True
        try:
            self.summary_table.setRowCount(len(data.vendors))
            for index, vendor in enumerate(data.vendors):
                rank = data.rank_for(index)
                values = [
                    vendor.name,
                    format_money(data.vendor_supply_total(index)),
                    format_money(data.vendor_vat(index)),
                    format_money(data.vendor_total(index)),
                    "-" if rank is None else f"{rank}위" + (" (자동)" if rank == 1 else ""),
                    "포함" if index in data.eligible_vendor_indices() else "제외",
                ]
                for column, value in enumerate(values):
                    self.summary_table.setItem(
                        index,
                        column,
                        text_item(value, editable=False, align_right=column in {1, 2, 3}),
                    )
            current_selection = data.selected_vendor_index
            self.selected_vendor_combo.clear()
            self.selected_vendor_combo.addItem("자동 선정(최저가)", None)
            for index, vendor in enumerate(data.vendors):
                self.selected_vendor_combo.addItem(vendor.name or f"업체 {index + 1}", index)
            if current_selection is not None and 0 <= current_selection < len(data.vendors):
                self.selected_vendor_combo.setCurrentIndex(current_selection + 1)
            else:
                self.selected_vendor_combo.setCurrentIndex(0)
            self.selected_vendor_combo.setToolTip(self.selected_vendor_combo.currentText())

            winner_index = data.winner_index
            if winner_index is None:
                self.selected_vendor_rank_label.setText("-")
                selected_supply = selected_vat = selected_total = 0
            else:
                rank = data.rank_for(winner_index)
                self.selected_vendor_rank_label.setText(
                    f"{rank}순위" if rank is not None else "직접 선정"
                )
                selected_supply = data.vendor_supply_total(winner_index)
                selected_vat = data.vendor_vat(winner_index)
                selected_total = data.vendor_total(winner_index)
            self.selected_vendor_supply_label.setText(f"₩ {format_money(selected_supply)}")
            self.selected_vendor_vat_label.setText(f"₩ {format_money(selected_vat)}")
            self.selected_vendor_total_label.setText(f"₩ {format_money(selected_total)}")

            self.budget_label.setText(f"₩ {format_money(data.budget_amount)}")
            self.contract_label.setText(f"₩ {format_money(data.contract_amount)}")
            self.ratio_label.setText(data.ratio_text)
            for metric in (
                self.selected_vendor_supply_label,
                self.selected_vendor_vat_label,
                self.selected_vendor_total_label,
                self.budget_label,
                self.contract_label,
                self.ratio_label,
            ):
                metric.setToolTip(metric.text())
            self.calculation_status.setText(
                f"정상 견적업체 {len(data.eligible_vendor_indices())}곳 · "
                f"가실행 {format_money(data.budget_amount)}원 · 계약 {format_money(data.contract_amount)}원"
            )
            if hasattr(self, "overview_vendor_count"):
                group_count = sum(1 for item in data.items if item.group_title.strip())
                detail_count = sum(1 for item in data.items if item.name.strip())
                self.overview_vendor_count.setText(f"{len(data.vendors)} 개")
                self.overview_group_count.setText(f"{group_count} 개")
                self.overview_item_count.setText(f"{detail_count} 개")
                self.overview_average_count.setText(f"{len(data.eligible_vendor_indices())} 개")
                winner = data.selected_vendor
                self.overview_winner.setText(winner.name if winner else "-")
                self.overview_amount.setText(f"₩ {format_money(data.contract_amount)}")
        finally:
            self._loading = previous_loading
        if self._statement_source_signature and self._statement_source_signature != self._quote_signature(data):
            self.statement_status_label.setText(
                "견적대비표가 변경되었습니다. 기존 내역서 수정내용은 유지 중이며, 필요할 때 ‘견적대비표에서 다시 불러오기’를 눌러 주세요."
            )
        self.refresh_purchase_auto_fields()
        self.refresh_statement_totals()
        self.schedule_preview_refresh()

    def refresh_purchase_auto_fields(self) -> None:
        data = self.data
        vendor = data.selected_vendor
        values = {
            "site": data.site_name,
            "budget": format_money(data.budget_amount),
            "contract": format_money(data.contract_amount),
            "ratio": data.ratio_text,
            "vendor": vendor.name if vendor else "",
            "phone": extract_phone_only(vendor.phone if vendor else ""),
        }
        widgets = {
            "site": self.purchase_site_edit,
            "budget": self.budget_edit,
            "contract": self.contract_edit,
            "ratio": self.purchase_ratio_edit,
            "vendor": self.vendor_edit,
            "phone": self.phone_edit,
        }
        for key, widget in widgets.items():
            if key not in self._purchase_overrides and not widget.hasFocus():
                widget.setText(values[key])
        if vendor and "payment" not in self._purchase_overrides and not self.payment_edit.hasFocus():
            payment = purchase_payment_text(vendor.payment)
            self.payment_edit.setText(payment)
            self.data.payment = payment

    def rebuild_statement_table(self, items: list[StatementItem] | None = None) -> None:
        items = list(items if items is not None else self.data.statement_items)
        previous_loading = self._loading
        self._loading = True
        try:
            self.statement_table.setRowCount(len(items))
            for row, item in enumerate(items):
                values = [
                    item.group_title,
                    item.number,
                    item.name,
                    item.spec,
                    item.unit,
                    decimal_text(item.quantity),
                    format_money(item.unit_price),
                    format_money(item.amount),
                    item.note,
                ]
                for column, value in enumerate(values):
                    self.statement_table.setItem(
                        row,
                        column,
                        text_item(
                            value,
                            align_right=column in {1, 5, 6, 7},
                            bold=column == 0 and bool(str(value).strip()),
                        ),
                    )
        finally:
            self._loading = previous_loading

    def collect_statement_items(self) -> list[StatementItem]:
        items: list[StatementItem] = []
        for row in range(self.statement_table.rowCount()):
            try:
                quantity = parse_decimal(table_text(self.statement_table, row, 5))
            except Exception:
                quantity = parse_decimal(0)
            items.append(
                StatementItem(
                    group_title=table_text(self.statement_table, row, 0),
                    number=table_text(self.statement_table, row, 1),
                    name=table_text(self.statement_table, row, 2),
                    spec=table_text(self.statement_table, row, 3),
                    unit=table_text(self.statement_table, row, 4),
                    quantity=quantity,
                    unit_price=parse_money(table_text(self.statement_table, row, 6)),
                    amount=parse_money(table_text(self.statement_table, row, 7)),
                    note=table_text(self.statement_table, row, 8),
                )
            )
        return items

    def sync_statement_from_quote(self, _checked: bool = False, *, confirm: bool = True) -> None:
        data = self.collect_project_data()
        if confirm and data.statement_items:
            answer = QMessageBox.question(
                self,
                "내역서 다시 불러오기",
                "내역서에서 직접 수정한 내용이 견적대비표 기준으로 교체됩니다. 계속할까요?",
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        data.sync_statement_from_quote()
        self.statement_title_edit.setText(
            next((item.group_title for item in data.statement_items if item.group_title), data.statement_title)
        )
        self.rebuild_statement_table(data.statement_items)
        self._statement_source_signature = self._quote_signature(data)
        self.statement_status_label.setText(
            "견적대비표에서 가져왔습니다. 이제 내역서의 모든 칸을 자유롭게 수정할 수 있습니다."
        )
        self.refresh_statement_totals()
        self.schedule_preview_refresh()

    def add_statement_item(self) -> None:
        items = self.collect_statement_items()
        items.append(StatementItem(number=str(len(items) + 1), unit="EA", quantity=1))
        self.data.statement_items = items
        self.rebuild_statement_table(items)
        self.statement_table.setCurrentCell(len(items) - 1, 2)
        self.refresh_statement_totals()

    def delete_statement_item(self) -> None:
        row = self.statement_table.currentRow()
        if row < 0:
            return
        items = self.collect_statement_items()
        items.pop(row)
        self.data.statement_items = items
        self.rebuild_statement_table(items)
        self.refresh_statement_totals()

    def move_statement_item(self, direction: int) -> None:
        row = self.statement_table.currentRow()
        items = self.collect_statement_items()
        target = row + direction
        if row < 0 or target < 0 or target >= len(items):
            return
        items[row], items[target] = items[target], items[row]
        self.data.statement_items = items
        self.rebuild_statement_table(items)
        self.statement_table.setCurrentCell(target, 2)
        self.refresh_statement_totals()

    def renumber_statement_items(self) -> None:
        previous_loading = self._loading
        self._loading = True
        try:
            for row in range(self.statement_table.rowCount()):
                self.statement_table.setItem(row, 1, text_item(str(row + 1), align_right=True))
        finally:
            self._loading = previous_loading
        self.data.statement_items = self.collect_statement_items()
        self.refresh_statement_totals()

    def refresh_statement_totals(self) -> None:
        data = self.data
        data.statement_items = self.collect_statement_items()
        vendor = data.selected_vendor
        self.statement_vendor_label.setText(vendor.name if vendor else "선정업체 없음")
        self.statement_supply_label.setText(f"₩ {format_money(data.statement_supply_total)}")
        self.statement_vat_label.setText(f"₩ {format_money(data.statement_vat)}")
        self.statement_total_label.setText(f"₩ {format_money(data.statement_total)}")
        self.statement_budget_label.setText(f"₩ {format_money(data.budget_amount)}")

    def on_statement_changed(self, row: int, column: int) -> None:
        if self._loading:
            return
        previous_loading = self._loading
        self._loading = True
        try:
            if column in {6, 7}:
                item = self.statement_table.item(row, column)
                if item:
                    item.setText(format_money(item.text()))
            if column in {5, 6}:
                try:
                    quantity = parse_decimal(table_text(self.statement_table, row, 5))
                except Exception:
                    quantity = parse_decimal(0)
                unit_price = parse_money(table_text(self.statement_table, row, 6))
                amount = quantity * unit_price
                self.statement_table.setItem(row, 7, text_item(format_money(amount), align_right=True))
        finally:
            self._loading = previous_loading
        self.data.statement_items = self.collect_statement_items()
        self.refresh_statement_totals()
        self.schedule_preview_refresh()

    def on_statement_title_changed(self, text: str) -> None:
        if self._loading:
            return
        self.data.statement_title = text
        if self.statement_table.rowCount() and not table_text(self.statement_table, 0, 0):
            previous_loading = self._loading
            self._loading = True
            try:
                self.statement_table.setItem(0, 0, text_item(text))
            finally:
                self._loading = previous_loading
            self.data.statement_items = self.collect_statement_items()
        self.schedule_preview_refresh()

    def on_selected_vendor_changed(self, index: int) -> None:
        if self._loading:
            return
        selected = self.selected_vendor_combo.itemData(index)
        self.data.selected_vendor_index = None if selected is None else int(selected)
        self.refresh_calculations()

    def on_tab_changed(self, index: int) -> None:
        self.refresh_calculations()
        if index == 1:
            if not self.data.statement_items:
                self.sync_statement_from_quote(confirm=False)
            else:
                self.refresh_statement_totals()
        elif index == 2:
            self.refresh_purchase_auto_fields()
        if self.preview_follow_check.isChecked():
            kind_for_tab = {0: "quote", 1: "statement", 2: "purchase"}.get(index, "quote")
            combo_index = self.preview_sheet_combo.findData(kind_for_tab)
            if combo_index >= 0:
                self.preview_sheet_combo.setCurrentIndex(combo_index)
        self.schedule_preview_refresh()

    # ------------------------------------------------------------------
    # Live preview / adjustable layout
    # ------------------------------------------------------------------
    def schedule_preview_refresh(self, *_args: object) -> None:
        if self._loading or not hasattr(self, "preview_timer"):
            return
        self.preview_timer.start()

    def refresh_preview(self) -> None:
        if self._loading or not self.preview_panel.isVisible():
            return
        data = self.collect_project_data()
        kind = str(self.preview_sheet_combo.currentData() or "quote")
        self.preview_browser.setHtml(build_preview_html(data, kind))

    def change_preview_zoom(self, delta: int) -> None:
        next_steps = max(-4, min(6, self._preview_zoom_steps + delta))
        actual_delta = next_steps - self._preview_zoom_steps
        if actual_delta > 0:
            self.preview_browser.zoomIn(actual_delta)
        elif actual_delta < 0:
            self.preview_browser.zoomOut(-actual_delta)
        self._preview_zoom_steps = next_steps
        self.preview_zoom_label.setText(f"{100 + self._preview_zoom_steps * 10}%")

    def toggle_preview(self, visible: bool) -> None:
        self.preview_panel.setVisible(visible)
        self.preview_action.setText("미리보기 숨기기" if visible else "미리보기 표시")
        if visible:
            saved = self.store.setting("main_splitter_sizes", [980, 560])
            if isinstance(saved, list) and len(saved) == 2:
                self.main_splitter.setSizes([int(saved[0]), max(360, int(saved[1]))])
            self.schedule_preview_refresh()

    def on_preview_kind_changed(self, index: int) -> None:
        kind = str(self.preview_sheet_combo.itemData(index) or "quote")
        button = self.preview_kind_buttons.get(kind) if hasattr(self, "preview_kind_buttons") else None
        if button:
            button.setChecked(True)

    def on_preview_follow_changed(self, checked: bool) -> None:
        if checked:
            self.on_tab_changed(self.tabs.currentIndex())
        else:
            self.schedule_preview_refresh()

    def set_quote_splitter_mode(self, mode: str) -> None:
        # 하단 순위/선정/가실행 카드는 분할 대상이 아니며 위 두 표만 조절한다.
        flexible_total = max(260, sum(self.quote_splitter.sizes()))
        if mode == "vendor":
            vendor_size = int(flexible_total * 0.58)
        elif mode == "item":
            vendor_size = int(flexible_total * 0.42)
        else:
            vendor_size = int(flexible_total * 0.50)
        item_size = flexible_total - vendor_size
        self.quote_splitter.setSizes([max(90, vendor_size), max(90, item_size)])
        self.save_layout_settings()

    def save_layout_settings(self, *_args: object) -> None:
        if self._loading:
            return
        self.store.set_setting("quote_splitter_sizes_v203", self.quote_splitter.sizes())
        self.store.set_setting("main_splitter_sizes", self.main_splitter.sizes())
        self.store.set_setting("right_splitter_sizes_v201", self.right_splitter.sizes())

    def restore_layout_settings(self) -> None:
        # 이전 버전의 작은 높이는 사용하지 않는다. 기본 화면에서 업체 5행이
        # 보이도록 v1.3.2 전용 설정 키와 새 기본 배치를 사용한다.
        quote_sizes = self.store.setting("quote_splitter_sizes_v203", [280, 300])
        if isinstance(quote_sizes, list) and len(quote_sizes) == 2:
            self.quote_splitter.setSizes([max(90, int(quote_sizes[0])), max(90, int(quote_sizes[1]))])
        main_sizes = self.store.setting("main_splitter_sizes", [900, 520])
        if isinstance(main_sizes, list) and len(main_sizes) == 2:
            self.main_splitter.setSizes([max(690, int(main_sizes[0])), max(350, int(main_sizes[1]))])
        right_sizes = self.store.setting("right_splitter_sizes_v201", [700, 184])
        if isinstance(right_sizes, list) and len(right_sizes) == 2:
            self.right_splitter.setSizes([max(180, int(right_sizes[0])), 184])
        self.preview_action.setText("미리보기 숨기기")
        self.schedule_preview_refresh()

    # ------------------------------------------------------------------
    # Classification / common sync
    # ------------------------------------------------------------------
    def update_classification(self) -> None:
        if self.manual_class_check.isChecked():
            return
        self.classification_edit.setText(
            build_classification(
                self.year_spin.value(),
                self.class_site_short_edit.text().strip(),
                self.sequence_spin.value(),
            )
        )

    def toggle_manual_classification(self, checked: bool) -> None:
        self.classification_edit.setReadOnly(not checked)
        self.year_spin.setEnabled(not checked)
        self.class_site_short_edit.setEnabled(not checked)
        self.sequence_spin.setEnabled(not checked)
        self.sequence_minus_button.setEnabled(not checked)
        self.sequence_plus_button.setEnabled(not checked)
        if not checked:
            self.update_classification()

    def sync_site_short_fields(self, text: str) -> None:
        if not self._loading and self.class_site_short_edit.text() != text:
            self.class_site_short_edit.setText(text)

    def sync_site_short_fields_reverse(self, text: str) -> None:
        if not self._loading and self.site_short_edit.text() != text:
            self.site_short_edit.setText(text)

    def return_home(self) -> None:
        if callable(self._home_callback):
            self._home_callback(self)
        else:
            self.close()

    # ------------------------------------------------------------------
    # Project file / export
    # ------------------------------------------------------------------
    def new_project(self) -> None:
        if QMessageBox.question(self, "신규 작성", "현재 입력을 초기화할까요?") != QMessageBox.StandardButton.Yes:
            return
        self.data = ProjectData()
        self._project_path = None
        self.file_status.setText("현재 파일: 새 문서")
        self._reset_with_defaults()

    def save_project(self) -> None:
        self.collect_project_data()
        initial = str(self._project_path or Path.home() / "purchase_request_project.json")
        path_text, _ = QFileDialog.getSaveFileName(
            self, "프로젝트 저장", initial, "구매품의서 프로젝트 (*.json)"
        )
        if not path_text:
            return
        path = Path(path_text)
        if path.suffix.lower() != ".json":
            path = path.with_suffix(".json")
        path.write_text(json.dumps(self.data.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
        self._project_path = path
        self.file_status.setText(f"현재 파일: {path.name}")

    def load_project(self) -> None:
        path_text, _ = QFileDialog.getOpenFileName(
            self, "프로젝트 불러오기", str(Path.home()), "구매품의서 프로젝트 (*.json)"
        )
        if not path_text:
            return
        path = Path(path_text)
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
            data = ProjectData.from_dict(payload)
        except (OSError, json.JSONDecodeError, TypeError, ValueError) as exc:
            QMessageBox.critical(self, "불러오기 실패", str(exc))
            return
        self.apply_data_to_ui(data)
        self._project_path = path
        self.file_status.setText(f"현재 파일: {path.name}")

    def apply_data_to_ui(self, data: ProjectData) -> None:
        previous_loading = self._loading
        self._loading = True
        try:
            self.data = data
            self.site_name_edit.setText(data.site_name)
            self.site_short_edit.setText(data.site_short)
            self.class_site_short_edit.setText(data.site_short)
            self.item_label_edit.setText(data.item_label)
            self.quote_title_edit.setText(data.quote_title)
            self.author_edit.setText(data.author)
            self.quote_date_edit.setDate(self._qdate_from_text(data.quote_date_text))
            self._common_delivery_place = data.common_delivery_place
            self.common_delivery_place_edit.setText(data.common_delivery_place)
            self.fill_vendor_table(data.vendors)
            self.item_notes = [item.note for item in data.items]
            self.rebuild_item_table(data.items)
            self.budget_mode_combo.setCurrentIndex(max(0, self.budget_mode_combo.findData(data.budget_mode)))
            self.manual_budget_supply_edit.setText(format_money(data.manual_budget_supply))
            self.won_rounding_combo.setCurrentIndex(max(0, self.won_rounding_combo.findData(data.won_rounding)))
            self.rebuild_statement_table(data.statement_items)
            self._statement_source_signature = self._quote_signature(data)
            self.classification_edit.setText(data.classification)
            self.department_edit.setText(data.department)
            self.draft_date_edit.setDate(self._qdate_from_text(data.draft_date_text))
            self.effective_date_edit.setText(data.effective_date)
            self.drafter_edit.setText(data.drafter)
            self.approval_note_edit.setText(data.approval_note)
            self.purchase_title_edit.setText(data.purchase_title)
            self.purchase_item_edit.setText(data.purchase_item_name)
            self.period_kind_combo.setCurrentText(data.period_kind)
            self.period_edit.setText(data.period)
            self.attachment_edit.setText(data.attachment)
            self.payment_edit.setText(purchase_payment_text(data.payment))
            self.body_edit.setPlainText(data.body_text)
            self.note_edit.setPlainText(data.note)
            self.statement_title_edit.setText(data.statement_title)
            self._purchase_overrides = set(data.purchase_override_fields)
            self.purchase_site_edit.setText(data.purchase_site_effective)
            self.vendor_edit.setText(data.purchase_vendor_effective)
            self.phone_edit.setText(data.purchase_phone_effective)
            self.budget_edit.setText(f"₩ {format_money(data.purchase_budget_effective)}")
            self.contract_edit.setText(f"₩ {format_money(data.purchase_contract_effective)}")
            self.purchase_ratio_edit.setText(data.purchase_ratio_text)
        finally:
            self._loading = previous_loading
        self.refresh_calculations()
        self.refresh_statement_totals()

    @staticmethod
    def _qdate_from_text(text: str) -> QDate:
        for fmt in ("yyyy.MM.dd", "yyyy-MM-dd"):
            parsed = QDate.fromString(text, fmt)
            if parsed.isValid():
                return parsed
        return QDate.currentDate()

    def export_excel(self) -> None:
        data = self.collect_project_data()
        errors = data.validate()
        if errors:
            QMessageBox.warning(self, "입력 확인", "\n".join(f"• {error}" for error in errors))
            return
        output_dir = Path(self.store.setting("output_dir", "") or Path.home() / "Documents")
        output_dir.mkdir(parents=True, exist_ok=True)
        initial = output_dir / data.suggested_filename()
        path_text, _ = QFileDialog.getSaveFileName(
            self, "구매품의서 Excel로 내보내기", str(initial), "Excel 통합문서 (*.xlsx)"
        )
        if not path_text:
            return
        output = Path(path_text)
        template_setting = str(self.store.setting("template_path", "") or "").strip()
        template = Path(template_setting) if template_setting else resource_path(
            "templates/purchase_request_3sheet_template.xlsx"
        )
        try:
            result = self.engine.export(data, template, output)
        except Exception as exc:  # 사용자에게 전체 오류 표시
            QMessageBox.critical(self, "내보내기 실패", str(exc))
            return
        self._last_export = result
        self.store.set_setting("output_dir", str(result.parent))
        self.file_status.setText(f"내보낸 파일: {result.name}")
        QMessageBox.information(
            self,
            "완료",
            "구매품의서 Excel 파일을 생성했습니다.\n\n"
            f"시트 순서: 구매품의서 → {data.statement_sheet_name} → {data.quote_sheet_name}\n"
            f"파일: {result}",
        )

    def choose_template(self) -> None:
        path_text, _ = QFileDialog.getOpenFileName(
            self,
            "3개 시트 양식 선택",
            str(Path.home()),
            "Excel 통합문서 (*.xlsx)",
        )
        if path_text:
            self.store.set_setting("template_path", path_text)
            QMessageBox.information(self, "양식 설정", "새 양식 파일을 저장했습니다.")

    def open_export_folder(self) -> None:
        folder = self._last_export.parent if self._last_export else Path(
            self.store.setting("output_dir", "") or Path.home() / "Documents"
        )
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))

    def show_about(self) -> None:
        QMessageBox.information(
            self,
            "프로그램 정보",
            f"{APP_TITLE} v2.0.0 - 구매 품의서 양식\n\n"
            "작성 순서\n"
            "1. 견적대비표 작성\n2. 내역서 확인·비고 작성\n3. 구매품의서 작성\n\n"
            "Excel 시트 순서\n"
            "구매품의서 → 내역서 → 견적대비표\n\n"
            "화면 기능\n"
            "업체표·품목표 높이 조절 / 우측 실시간 미리보기",
        )
