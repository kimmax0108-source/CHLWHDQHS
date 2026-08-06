from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, Optional

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, Signal
from PySide6.QtGui import QColor, QDesktopServices, QIcon
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QTabWidget,
    QTableView,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtCore import QUrl

from material_document_app.resource import resource_path

from .excel_io import list_sheet_names, load_ledgers, preferred_ledger_sheets
from .exporter import export_rows
from .models import LedgerData, LedgerRow, MoneySummary
from .services import (
    cancel_rollover,
    claim_breakdown_period,
    effective_claim_month,
    effective_trade,
    exclude_rows,
    grouped_summary,
    is_excluded,
    money_summary,
    processing_status,
    reset_classification,
    restore_excluded_rows,
    rollover_status,
    row_info,
    set_claim_month,
    set_classification,
    shift_month,
)
from .storage import ClaimOverrideStore, PresetStore

APP_TITLE = "자재 입고 청구관리"
APP_VERSION = "2.0.0"


def format_money(value: float | int | None, *, signed: bool = False) -> str:
    if value is None:
        return "-"
    number = int(round(value))
    if signed and number > 0:
        return f"+{number:,}"
    return f"{number:,}"


def format_quantity(value: float | int | None) -> str:
    if value is None:
        return "-"
    text = f"{float(value):,.3f}"
    return text.rstrip("0").rstrip(".")


def _period_from_controls(year_combo: QComboBox, month_combo: QComboBox) -> tuple[Optional[int], Optional[int]]:
    year = year_combo.currentData()
    month = month_combo.currentData()
    return (int(year) if year is not None else None, int(month) if month is not None else None)


class MultiSelectPanel(QFrame):
    changed = Signal()

    def __init__(self, title: str, placeholder: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.title = title
        self.setObjectName("filterSubCard")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(7)

        top = QHBoxLayout()
        label = QLabel(title)
        label.setObjectName("subCardTitle")
        self.count_label = QLabel("전체")
        self.count_label.setObjectName("selectionCount")
        top.addWidget(label)
        top.addStretch(1)
        top.addWidget(self.count_label)
        layout.addLayout(top)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText(placeholder)
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.textChanged.connect(self._apply_search)
        layout.addWidget(self.search_edit)

        self.list_widget = QListWidget()
        self.list_widget.setMinimumHeight(118)
        self.list_widget.setMaximumHeight(150)
        self.list_widget.itemChanged.connect(self._item_changed)
        layout.addWidget(self.list_widget, 1)

        buttons = QHBoxLayout()
        all_button = QPushButton("전체 선택")
        clear_button = QPushButton("전체 해제")
        all_button.setObjectName("smallButton")
        clear_button.setObjectName("smallButton")
        all_button.clicked.connect(self.select_all_visible)
        clear_button.clicked.connect(self.clear)
        buttons.addWidget(all_button)
        buttons.addWidget(clear_button)
        buttons.addStretch(1)
        layout.addLayout(buttons)

    def set_options(self, values: Iterable[str], *, select_all: bool = False, preserve: bool = True) -> None:
        current = self.selected_values() if preserve else set()
        options = sorted({str(value).strip() for value in values if str(value).strip()}, key=str.casefold)
        self.list_widget.blockSignals(True)
        self.list_widget.clear()
        for value in options:
            item = QListWidgetItem(value)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            checked = select_all or value in current
            item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
            self.list_widget.addItem(item)
        self.list_widget.blockSignals(False)
        self._apply_search(self.search_edit.text())
        self._update_count()

    def selected_values(self) -> set[str]:
        values: set[str] = set()
        for index in range(self.list_widget.count()):
            item = self.list_widget.item(index)
            if item.checkState() == Qt.CheckState.Checked:
                values.add(item.text())
        return values

    def set_selected(self, values: Iterable[str]) -> None:
        selected = {str(value) for value in values}
        self.list_widget.blockSignals(True)
        for index in range(self.list_widget.count()):
            item = self.list_widget.item(index)
            item.setCheckState(
                Qt.CheckState.Checked if item.text() in selected else Qt.CheckState.Unchecked
            )
        self.list_widget.blockSignals(False)
        self._update_count()
        self.changed.emit()

    def clear(self) -> None:
        self.set_selected(set())

    def select_all_visible(self) -> None:
        self.list_widget.blockSignals(True)
        for index in range(self.list_widget.count()):
            item = self.list_widget.item(index)
            if not item.isHidden():
                item.setCheckState(Qt.CheckState.Checked)
        self.list_widget.blockSignals(False)
        self._update_count()
        self.changed.emit()

    def _apply_search(self, text: str) -> None:
        needle = text.strip().casefold()
        for index in range(self.list_widget.count()):
            item = self.list_widget.item(index)
            item.setHidden(bool(needle and needle not in item.text().casefold()))

    def _item_changed(self, _item: QListWidgetItem) -> None:
        self._update_count()
        self.changed.emit()

    def _update_count(self) -> None:
        count = len(self.selected_values())
        self.count_label.setText(f"{count}개 선택" if count else "전체")


class SummaryCard(QFrame):
    def __init__(self, title: str, subtitle: str, accent: str, *, strong: bool = False) -> None:
        super().__init__()
        self.setProperty("strong", strong)
        self.setObjectName("summaryCardStrong" if strong else "summaryCard")
        self.setMinimumHeight(126)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 11, 15, 13)
        layout.setSpacing(5)

        accent_bar = QFrame()
        accent_bar.setFixedHeight(4)
        accent_bar.setStyleSheet(f"background: {accent}; border: 0; border-radius: 2px;")
        layout.addWidget(accent_bar)

        title_label = QLabel(title)
        title_label.setObjectName("summaryTitleStrong" if strong else "summaryTitle")
        title_label.setStyleSheet(f"font-weight: 900; color: {accent};")
        subtitle_label = QLabel(subtitle)
        subtitle_label.setObjectName("summarySubtitle")
        layout.addWidget(title_label)
        layout.addWidget(subtitle_label)
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setObjectName("summaryLine")
        layout.addWidget(line)
        grid = QGridLayout()
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setVerticalSpacing(3)
        self.value_labels: list[QLabel] = []
        for row, label_text in enumerate(("공급가액", "부가세", "최종금액")):
            grid.addWidget(QLabel(label_text), row, 0)
            value_label = QLabel("0 원")
            value_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            value_label.setStyleSheet(f"font-weight: 850; color: {accent if row == 2 else '#172033'};")
            grid.addWidget(value_label, row, 1)
            self.value_labels.append(value_label)
        grid.setColumnStretch(1, 1)
        layout.addLayout(grid)

    def set_values(self, summary: MoneySummary, *, signed: bool = False) -> None:
        for label, value in zip(self.value_labels, (summary.supply, summary.vat, summary.total)):
            label.setText(f"{format_money(value, signed=signed)} 원")


@dataclass(frozen=True)
class ColumnSpec:
    title: str
    width: int
    numeric: bool = False


class LedgerTableModel(QAbstractTableModel):
    COLUMNS = [
        ColumnSpec("원본 시트", 90),
        ColumnSpec("입고일", 82),
        ColumnSpec("원본 공종", 84),
        ColumnSpec("청구분류", 88),
        ColumnSpec("품명", 160),
        ColumnSpec("규격", 145),
        ColumnSpec("단위", 55),
        ColumnSpec("수량", 86, True),
        ColumnSpec("단가", 105, True),
        ColumnSpec("공급가액(원본)", 125, True),
        ColumnSpec("계산금액", 115, True),
        ColumnSpec("차이금액", 100, True),
        ColumnSpec("검토상태", 76),
        ColumnSpec("입고월", 78),
        ColumnSpec("적용 청구월", 88),
        ColumnSpec("이월상태", 80),
        ColumnSpec("구입처", 145),
        ColumnSpec("비고/용도", 210),
        ColumnSpec("처리상태", 84),
        ColumnSpec("관리메모", 150),
    ]

    def __init__(self, store_getter: Callable[[], Optional[ClaimOverrideStore]]) -> None:
        super().__init__()
        self.rows: list[LedgerRow] = []
        self._store_getter = store_getter

    def set_rows(self, rows: Iterable[LedgerRow]) -> None:
        self.beginResetModel()
        self.rows = list(rows)
        self.endResetModel()

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        return 0 if parent.isValid() else len(self.COLUMNS)

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole):  # noqa: N802
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal and 0 <= section < len(self.COLUMNS):
            return self.COLUMNS[section].title
        if orientation == Qt.Orientation.Vertical:
            return section + 1
        return None

    def _raw_value(self, row: LedgerRow, column: int):
        store = self._store_getter()
        if store is None:
            return ""
        values = [
            row.source_sheet or "단일대장",
            row.intake_date.strftime("%y.%m.%d"),
            row.trade,
            effective_trade(row, store),
            row.item,
            row.spec,
            row.unit,
            row.quantity,
            row.unit_price if row.unit_price_entered else None,
            row.amount,
            row.calculated_amount,
            row.amount_difference,
            row.amount_review_status,
            row.intake_month,
            effective_claim_month(row, store),
            rollover_status(row, store),
            row.vendor,
            row.usage or row.note,
            processing_status(row, store),
            store.management_note(row.fingerprint),
        ]
        return values[column]

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or not (0 <= index.row() < len(self.rows)):
            return None
        row = self.rows[index.row()]
        column = index.column()
        store = self._store_getter()
        if store is None:
            return None
        raw = self._raw_value(row, column)

        if role == Qt.ItemDataRole.DisplayRole:
            if column == 7:
                return format_quantity(raw)
            if column in (8, 9, 10, 11):
                return format_money(raw, signed=column == 11)
            return "" if raw is None else str(raw)
        if role == Qt.ItemDataRole.TextAlignmentRole:
            if self.COLUMNS[column].numeric:
                return int(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            return int(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        if role == Qt.ItemDataRole.ToolTipRole:
            return "" if raw is None else str(raw)
        if role == Qt.ItemDataRole.BackgroundRole:
            if is_excluded(row, store):
                return QColor("#FDE7EA")
            if row.amount_review_status == "불일치":
                return QColor("#FDE8E7")
            if row.amount_review_status == "검토불가":
                return QColor("#FFF4D6")
            claim_month = effective_claim_month(row, store)
            if claim_month > row.intake_month:
                return QColor("#FFF0DD")
            if claim_month < row.intake_month:
                return QColor("#E8F1FF")
        if role == Qt.ItemDataRole.ForegroundRole:
            if column == 12:
                return QColor("#B42318" if row.amount_review_status == "불일치" else "#18794E")
            if column == 18 and is_excluded(row, store):
                return QColor("#B42318")
        return None

    def sort(self, column: int, order: Qt.SortOrder = Qt.SortOrder.AscendingOrder) -> None:
        store = self._store_getter()
        if store is None:
            return
        self.layoutAboutToBeChanged.emit()

        def key(row: LedgerRow):
            value = self._raw_value(row, column)
            if value is None:
                return (1, "")
            if isinstance(value, (int, float)):
                return (0, float(value))
            return (0, str(value).casefold())

        self.rows.sort(key=key, reverse=order == Qt.SortOrder.DescendingOrder)
        self.layoutChanged.emit()


class SummaryTableModel(QAbstractTableModel):
    HEADERS = ["구분", "건수", "수량", "공급가액", "부가세", "최종금액"]

    def __init__(self) -> None:
        super().__init__()
        self.rows: list[tuple[str, MoneySummary]] = []

    def set_rows(self, rows: Iterable[tuple[str, MoneySummary]]) -> None:
        self.beginResetModel()
        self.rows = sorted(list(rows), key=lambda value: (-value[1].supply, value[0].casefold()))
        self.endResetModel()

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        return 0 if parent.isValid() else len(self.HEADERS)

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole):  # noqa: N802
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return self.HEADERS[section]
            return section + 1
        return None

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        name, summary = self.rows[index.row()]
        raw = [name, summary.count, summary.quantity, summary.supply, summary.vat, summary.total][index.column()]
        if role == Qt.ItemDataRole.DisplayRole:
            if index.column() == 2:
                return format_quantity(raw)
            if index.column() >= 3:
                return format_money(raw)
            return str(raw)
        if role == Qt.ItemDataRole.TextAlignmentRole:
            return int(
                (Qt.AlignmentFlag.AlignRight if index.column() >= 1 else Qt.AlignmentFlag.AlignLeft)
                | Qt.AlignmentFlag.AlignVCenter
            )
        return None


class RowSelectionDialog(QDialog):
    def __init__(self, title: str, rows: list[LedgerRow], store: ClaimOverrideStore, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(980, 600)
        layout = QVBoxLayout(self)
        guide = QLabel("가져올 행을 선택하십시오. Ctrl/Shift로 여러 행을 선택할 수 있습니다.")
        guide.setObjectName("dialogGuide")
        layout.addWidget(guide)
        self.table = QTableWidget(len(rows), 7)
        self.table.setHorizontalHeaderLabels(["원본 시트", "입고일", "품명", "규격", "구입처", "공급가액", "현재 청구월"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self._rows = rows
        for r, row in enumerate(rows):
            values = [
                row.source_sheet or "단일대장",
                row.intake_date.strftime("%y.%m.%d"),
                row.item,
                row.spec,
                row.vendor,
                format_money(row.amount),
                effective_claim_month(row, store),
            ]
            for c, value in enumerate(values):
                self.table.setItem(r, c, QTableWidgetItem(str(value)))
        layout.addWidget(self.table, 1)
        select_all = QPushButton("전체 선택")
        select_all.clicked.connect(self.table.selectAll)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        bottom = QHBoxLayout()
        bottom.addWidget(select_all)
        bottom.addStretch(1)
        bottom.addWidget(buttons)
        layout.addLayout(bottom)

    def selected_rows(self) -> list[LedgerRow]:
        indices = sorted({index.row() for index in self.table.selectionModel().selectedRows()})
        return [self._rows[index] for index in indices]


class SheetSelectionDialog(QDialog):
    def __init__(self, sheets: list[str], selected: set[str], parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("검토 시트 선택")
        self.resize(420, 500)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("불러올 자재입출고대장 시트를 선택하십시오."))
        self.list_widget = QListWidget()
        for sheet in sheets:
            item = QListWidgetItem(sheet)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked if sheet in selected else Qt.CheckState.Unchecked)
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def selected_sheets(self) -> list[str]:
        return [
            self.list_widget.item(index).text()
            for index in range(self.list_widget.count())
            if self.list_widget.item(index).checkState() == Qt.CheckState.Checked
        ]


class HistoryDialog(QDialog):
    ACTION_NAMES = {
        "claim_month": "청구월 변경",
        "claim_month_reset": "이월 취소",
        "classification": "분류 변경",
        "classification_reset": "분류 원복",
        "exclude": "청구 제외",
        "exclude_reset": "제외 취소",
        "management_note": "관리메모",
    }

    def __init__(self, history: list[dict], parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("관리 변경 이력")
        self.resize(1150, 650)
        layout = QVBoxLayout(self)
        table = QTableWidget(len(history), 9)
        table.setHorizontalHeaderLabels(
            ["처리일시", "처리유형", "원본 시트", "품명", "구입처", "기존값", "변경값", "사유", "원본행"]
        )
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)
        for r, entry in enumerate(reversed(history)):
            info = entry.get("row") or {}
            values = [
                entry.get("at", ""),
                self.ACTION_NAMES.get(str(entry.get("action", "")), str(entry.get("action", ""))),
                info.get("source_sheet", ""),
                info.get("item", ""),
                info.get("vendor", ""),
                entry.get("old", ""),
                entry.get("new", ""),
                entry.get("reason", ""),
                info.get("source_row", ""),
            ]
            for c, value in enumerate(values):
                table.setItem(r, c, QTableWidgetItem(str(value)))
        layout.addWidget(table)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        buttons.clicked.connect(self.accept)
        layout.addWidget(buttons)


class MaterialClaimWindow(QMainWindow):
    def __init__(self, home_callback=None) -> None:
        super().__init__()
        self.home_callback = home_callback
        self.data: Optional[LedgerData] = None
        self.store: Optional[ClaimOverrideStore] = None
        self.preset_store = PresetStore()
        self.current_breakdown: dict[str, list[LedgerRow]] = {}
        self._loading = False
        self._workbook_hash_before = ""

        self.setWindowTitle(f"{APP_TITLE} v{APP_VERSION}")
        self.resize(1640, 950)
        self.setMinimumSize(1280, 760)
        icon_path = resource_path("assets/app.ico")
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        self._build_ui()
        self._apply_style()
        self._refresh_presets()

    def _icon(self, name: str) -> QIcon:
        path = resource_path(f"assets/{name}.svg")
        return QIcon(str(path)) if path.exists() else QIcon()

    def _build_ui(self) -> None:
        central = QWidget()
        central.setObjectName("claimRoot")
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        self.setCentralWidget(central)
        root.addWidget(self._build_sidebar())

        workspace = QWidget()
        workspace.setObjectName("claimWorkspace")
        body = QVBoxLayout(workspace)
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(0)
        body.addWidget(self._build_topbar())

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(18, 16, 18, 10)
        content_layout.setSpacing(12)

        upper_splitter = QSplitter(Qt.Orientation.Horizontal)
        upper_splitter.setObjectName("upperSplitter")
        upper_splitter.setChildrenCollapsible(False)
        upper_splitter.addWidget(self._build_filter_card())
        upper_splitter.addWidget(self._build_summary_area())
        upper_splitter.setHandleWidth(12)
        upper_splitter.setStretchFactor(0, 58)
        upper_splitter.setStretchFactor(1, 42)
        upper_splitter.setSizes([900, 660])
        upper_splitter.setMinimumHeight(398)
        upper_splitter.setMaximumHeight(420)
        content_layout.addWidget(upper_splitter)

        content_layout.addLayout(self._build_action_bar())
        self.tabs = QTabWidget()
        self.tabs.setObjectName("claimTabs")
        self.detail_models: dict[str, LedgerTableModel] = {}
        self.detail_views: dict[str, QTableView] = {}
        detail_tabs = (
            ("current_intake", "당월 입고분", "#2563EB"),
            ("brought_in", "전월 이월분", "#0F8F70"),
            ("moved_out", "다음 달 이월분", "#D97706"),
            ("claim_target", "실제 청구대상", "#7C3AED"),
        )
        for key, title, accent in detail_tabs:
            view, model = self._make_detail_table()
            self.detail_views[key] = view
            self.detail_models[key] = model
            tab_index = self.tabs.addTab(view, title)
            self.tabs.tabBar().setTabTextColor(tab_index, QColor(accent))

        self.summary_models: dict[str, SummaryTableModel] = {}
        for key, title in (("item", "품명별 집계"), ("vendor", "구입처별 집계"), ("sheet", "시트별 집계")):
            view, model = self._make_summary_table()
            self.summary_models[key] = model
            self.tabs.addTab(view, title)
        self.tabs.setMinimumHeight(350)
        content_layout.addWidget(self.tabs, 1)

        footer = QHBoxLayout()
        self.status_label = QLabel("엑셀 파일을 불러와 주세요.")
        self.count_label = QLabel("데이터 건수: 0건")
        self.count_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        footer.addWidget(self.status_label, 1)
        footer.addWidget(self.count_label)
        content_layout.addLayout(footer)

        body.addWidget(content, 1)
        root.addWidget(workspace, 1)

    def _build_sidebar(self) -> QFrame:
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(238)
        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(18, 22, 18, 18)
        layout.setSpacing(9)
        brand = QHBoxLayout()
        icon = QLabel()
        icon.setPixmap(self._icon("database").pixmap(38, 38))
        texts = QVBoxLayout()
        title = QLabel("자재 문서 표준화")
        title.setObjectName("brandTitle")
        subtitle = QLabel("자재 입고 청구관리")
        subtitle.setObjectName("brandSubtitle")
        texts.addWidget(title)
        texts.addWidget(subtitle)
        brand.addWidget(icon)
        brand.addLayout(texts, 1)
        layout.addLayout(brand)
        layout.addSpacing(12)

        section = QLabel("조회 및 집계")
        section.setObjectName("sidebarSection")
        layout.addWidget(section)
        for label, tab_index in (("상세내역", 3), ("품명별 집계", 4), ("구입처별 집계", 5), ("시트별 집계", 6)):
            button = QPushButton(label)
            button.setObjectName("navButton")
            button.clicked.connect(lambda _checked=False, i=tab_index: self.tabs.setCurrentIndex(i))
            layout.addWidget(button)

        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setObjectName("sidebarDivider")
        layout.addWidget(divider)
        section2 = QLabel("관리 데이터")
        section2.setObjectName("sidebarSection")
        layout.addWidget(section2)
        history = QPushButton("이월·변경 이력")
        backup = QPushButton("관리 JSON 백업")
        restore = QPushButton("관리 JSON 복원")
        for button in (history, backup, restore):
            button.setObjectName("navButton")
            layout.addWidget(button)
        history.clicked.connect(self.show_history)
        backup.clicked.connect(self.manual_backup)
        restore.clicked.connect(self.restore_backup)
        layout.addStretch(1)
        home = QPushButton(self._icon("home"), "홈으로")
        home.setObjectName("sidebarFooterButton")
        home.clicked.connect(self.return_home)
        layout.addWidget(home)
        version = QLabel(f"v{APP_VERSION}")
        version.setObjectName("versionLabel")
        layout.addWidget(version)
        return sidebar

    def _build_topbar(self) -> QFrame:
        bar = QFrame()
        bar.setObjectName("topbar")
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(20, 12, 20, 12)
        title = QLabel("자재 입고 청구관리")
        title.setObjectName("pageTitle")
        layout.addWidget(title)
        self.file_label = QLabel("파일 미선택")
        self.file_label.setObjectName("filePathLabel")
        self.file_label.setMaximumWidth(360)
        layout.addWidget(self.file_label)
        layout.addStretch(1)

        self.preset_combo = QComboBox()
        self.preset_combo.setMinimumWidth(170)
        self.preset_combo.setPlaceholderText("조회 프리셋")
        layout.addWidget(self.preset_combo)
        for text, slot in (("적용", self.apply_preset), ("새로 저장", self.save_preset), ("수정", self.update_preset), ("삭제", self.delete_preset)):
            button = QPushButton(text)
            button.setObjectName("topActionButton")
            button.clicked.connect(slot)
            layout.addWidget(button)
        self.load_button = QPushButton(self._icon("folder"), "엑셀 파일 불러오기")
        self.export_button = QPushButton(self._icon("excel"), "조회결과 Excel로 내보내기")
        self.load_button.setObjectName("primaryButton")
        self.export_button.setObjectName("purpleButton")
        self.load_button.clicked.connect(self.choose_file)
        self.export_button.clicked.connect(self.export_current)
        layout.addWidget(self.load_button)
        layout.addWidget(self.export_button)
        return bar

    def _build_filter_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        card.setMinimumHeight(398)
        card.setMaximumHeight(420)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(15, 14, 15, 14)
        layout.setSpacing(9)
        heading = QHBoxLayout()
        title = QLabel("1. 조회 조건 — 입고월 기준")
        title.setObjectName("cardTitle")
        heading.addWidget(title)
        heading.addStretch(1)
        self.sheet_select_button = QPushButton("검토 시트 선택")
        self.sheet_select_button.setObjectName("smallButton")
        self.sheet_select_button.clicked.connect(self.select_source_sheets)
        heading.addWidget(self.sheet_select_button)
        layout.addLayout(heading)

        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(7)
        self.year_combo = QComboBox()
        self.month_combo = QComboBox()
        self.trade_combo = QComboBox()
        self.keyword_edit = QLineEdit()
        self.keyword_edit.setPlaceholderText("품명·규격·구입처·용도·메모 통합검색")
        self.keyword_edit.setClearButtonEnabled(True)
        self.year_combo.addItem("전체", None)
        self.month_combo.addItem("전체", None)
        for month in range(1, 13):
            self.month_combo.addItem(f"{month}월", month)
        self.trade_combo.addItem("전체")
        form.addWidget(QLabel("연도"), 0, 0)
        form.addWidget(self.year_combo, 0, 1)
        form.addWidget(QLabel("월"), 0, 2)
        form.addWidget(self.month_combo, 0, 3)
        form.addWidget(QLabel("공종"), 0, 4)
        form.addWidget(self.trade_combo, 0, 5)
        form.addWidget(QLabel("통합검색"), 1, 0)
        form.addWidget(self.keyword_edit, 1, 1, 1, 5)
        form.setColumnStretch(1, 1)
        form.setColumnStretch(3, 1)
        form.setColumnStretch(5, 2)
        layout.addLayout(form)

        panels = QHBoxLayout()
        self.source_panel = MultiSelectPanel("원본 시트", "시트 검색")
        self.vendor_panel = MultiSelectPanel("구입처", "구입처 검색")
        self.item_panel = MultiSelectPanel("품명", "품명 검색")
        self.source_panel.setMaximumWidth(220)
        panels.addWidget(self.source_panel)
        panels.addWidget(self.vendor_panel, 1)
        panels.addWidget(self.item_panel, 1)
        layout.addLayout(panels)

        controls = QHBoxLayout()
        self.mismatch_check = QCheckBox("금액 불일치·검토불가만 보기")
        self.query_button = QPushButton("조회")
        self.reset_button = QPushButton("조건 초기화")
        self.query_button.setObjectName("primaryButton")
        self.reset_button.setObjectName("secondaryButton")
        self.query_button.clicked.connect(self.apply_filters)
        self.reset_button.clicked.connect(self.reset_filters)
        controls.addWidget(self.mismatch_check)
        controls.addStretch(1)
        controls.addWidget(self.reset_button)
        controls.addWidget(self.query_button)
        layout.addLayout(controls)
        return card

    def _build_summary_area(self) -> QFrame:
        wrapper = QFrame()
        wrapper.setObjectName("card")
        wrapper.setMinimumHeight(398)
        wrapper.setMaximumHeight(420)
        layout = QVBoxLayout(wrapper)
        layout.setContentsMargins(15, 14, 15, 14)
        layout.setSpacing(9)
        heading = QLabel("2. 청구 현황 요약")
        heading.setObjectName("cardTitle")
        layout.addWidget(heading)
        grid = QGridLayout()
        grid.setSpacing(10)
        self.summary_cards = {
            "current_intake": SummaryCard("당월 입고분", "이번 달 입고 · 이번 달 청구", "#2563EB"),
            "brought_in": SummaryCard("전월 이월분", "이전 달 입고 · 이번 달 청구", "#0F8F70"),
            "moved_out": SummaryCard("다음 달 이월분", "이번 달 입고 · 다음 달 청구", "#D97706"),
            "claim_target": SummaryCard("실제 청구대상", "당월 + 전월 이월 − 다음 달 이월", "#7C3AED", strong=True),
        }
        for index, key in enumerate(("current_intake", "brought_in", "moved_out", "claim_target")):
            grid.addWidget(self.summary_cards[key], index // 2, index % 2)
        layout.addLayout(grid)
        info = QFrame()
        info.setObjectName("noticeCard")
        info_layout = QVBoxLayout(info)
        info_layout.setContentsMargins(12, 10, 12, 10)
        self.site_label = QLabel("현장명: -")
        self.loaded_sheet_label = QLabel("불러온 시트: -")
        self.validation_label = QLabel("금액 검토: -")
        for label in (self.site_label, self.loaded_sheet_label, self.validation_label):
            info_layout.addWidget(label)
        layout.addWidget(info)
        return wrapper

    def _build_action_bar(self) -> QHBoxLayout:
        bar = QHBoxLayout()
        bar.setSpacing(7)
        actions = [
            ("이전 달에서 가져오기", self.import_previous, "blueAction"),
            ("다음 달로 이월", self.move_selected_next, "orangeAction"),
            ("이월 취소", self.cancel_selected_rollover, "secondaryButton"),
            ("청구분류 변경", self.change_selected_classification, "secondaryButton"),
            ("분류 원복", self.reset_selected_classification, "secondaryButton"),
            ("청구대상 제외", self.exclude_selected, "dangerButton"),
            ("제외 취소", self.restore_selected_excluded, "secondaryButton"),
            ("관리메모", self.edit_management_note, "secondaryButton"),
            ("이월·변경 이력", self.show_history, "secondaryButton"),
        ]
        for text, slot, object_name in actions:
            button = QPushButton(text)
            button.setObjectName(object_name)
            button.clicked.connect(slot)
            bar.addWidget(button)
        bar.addStretch(1)
        self.open_original_button = QPushButton("원본 Excel 열기")
        self.open_original_button.setObjectName("secondaryButton")
        self.open_original_button.clicked.connect(self.open_original)
        bar.addWidget(self.open_original_button)
        return bar

    def _make_detail_table(self) -> tuple[QTableView, LedgerTableModel]:
        model = LedgerTableModel(lambda: self.store)
        view = QTableView()
        view.setModel(model)
        view.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        view.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        view.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        view.setAlternatingRowColors(True)
        view.setSortingEnabled(True)
        view.setWordWrap(False)
        view.verticalHeader().setDefaultSectionSize(27)
        view.horizontalHeader().setMinimumSectionSize(52)
        view.horizontalHeader().setSectionsMovable(True)
        for index, column in enumerate(LedgerTableModel.COLUMNS):
            view.setColumnWidth(index, column.width)
        return view, model

    def _make_summary_table(self) -> tuple[QTableView, SummaryTableModel]:
        model = SummaryTableModel()
        view = QTableView()
        view.setModel(model)
        view.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        view.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        view.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for column in range(1, 6):
            view.horizontalHeader().setSectionResizeMode(column, QHeaderView.ResizeMode.ResizeToContents)
        return view, model

    # ------------------------------------------------------------------
    # File loading and query
    # ------------------------------------------------------------------
    def choose_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "자재입출고대장 불러오기",
            str(Path.home()),
            "Excel 파일 (*.xlsx *.xlsm *.xls)",
        )
        if path:
            self.load_file(path)

    def load_file(self, path: str, selected_sheets: Optional[list[str]] = None) -> None:
        try:
            sheets = list_sheet_names(path)
            selected = selected_sheets or preferred_ledger_sheets(sheets)
            if not selected:
                selected = sheets[:1]
            data = load_ledgers(path, sheet_names=selected)
            store = ClaimOverrideStore(path)
        except Exception as exc:
            QMessageBox.critical(self, "불러오기 실패", str(exc))
            return
        self.data = data
        self.store = store
        self.file_label.setText(Path(path).name)
        self.file_label.setToolTip(path)
        self.site_label.setText(f"현장명: {data.site_name or '-'}")
        self.loaded_sheet_label.setText(f"불러온 시트: {', '.join(data.loaded_sheet_names)}")
        self._populate_filters()
        self.apply_filters()
        self.status_label.setText(f"불러오기 완료: {Path(path).name}")
        self.count_label.setText(f"데이터 건수: {len(data.rows):,}건")

    def _populate_filters(self) -> None:
        if not self.data or not self.store:
            return
        self._loading = True
        try:
            years = sorted({row.intake_date.year for row in self.data.rows})
            self.year_combo.clear()
            self.year_combo.addItem("전체", None)
            for year in years:
                self.year_combo.addItem(str(year), year)
            if years:
                self.year_combo.setCurrentIndex(self.year_combo.findData(max(years)))
            trades = sorted({effective_trade(row, self.store) for row in self.data.rows if effective_trade(row, self.store)}, key=str.casefold)
            self.trade_combo.clear()
            self.trade_combo.addItem("전체")
            self.trade_combo.addItems(trades)
            self.source_panel.set_options(self.data.loaded_sheet_names or ["단일대장"], select_all=True, preserve=False)
            self.vendor_panel.set_options((row.vendor for row in self.data.rows), preserve=False)
            self.item_panel.set_options((row.item for row in self.data.rows), preserve=False)
        finally:
            self._loading = False

    def select_source_sheets(self) -> None:
        if not self.data:
            QMessageBox.information(self, "안내", "먼저 Excel 파일을 불러와 주세요.")
            return
        dialog = SheetSelectionDialog(
            self.data.sheet_names,
            set(self.data.loaded_sheet_names),
            self,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        selected = dialog.selected_sheets()
        if not selected:
            QMessageBox.warning(self, "시트 선택", "한 개 이상의 시트를 선택해야 합니다.")
            return
        self.load_file(self.data.path, selected)

    def reset_filters(self) -> None:
        if not self.data:
            return
        self._loading = True
        try:
            self.year_combo.setCurrentIndex(0)
            self.month_combo.setCurrentIndex(0)
            self.trade_combo.setCurrentIndex(0)
            self.keyword_edit.clear()
            self.vendor_panel.clear()
            self.item_panel.clear()
            self.source_panel.set_selected(self.data.loaded_sheet_names or ["단일대장"])
            self.mismatch_check.setChecked(False)
        finally:
            self._loading = False
        self.apply_filters()

    def apply_filters(self) -> None:
        if self._loading or not self.data or not self.store:
            return
        year, month = _period_from_controls(self.year_combo, self.month_combo)
        source_sheets = self.source_panel.selected_values()
        vendors = self.vendor_panel.selected_values()
        items = self.item_panel.selected_values()
        trade = self.trade_combo.currentText() or "전체"
        breakdown = claim_breakdown_period(
            self.data.rows,
            self.store,
            year=year,
            month=month,
            trade=trade,
            vendors=vendors or None,
            items=items or None,
            source_sheets=source_sheets or None,
            keyword=self.keyword_edit.text(),
        )
        self.current_breakdown = breakdown
        mismatch_only = self.mismatch_check.isChecked()

        display_map = {
            "current_intake": breakdown["display_current_intake"],
            "brought_in": breakdown["display_brought_in"],
            "moved_out": breakdown["display_moved_out"],
            "claim_target": breakdown["display_claim_rows"],
        }
        for key, values in display_map.items():
            rows = [row for row in values if row.amount_review_status != "정상"] if mismatch_only else values
            self.detail_models[key].set_rows(rows)
            index = list(self.detail_models).index(key)
            self.tabs.setTabText(index, f"{self._detail_title(key)} ({len(rows):,})")

        for key in ("current_intake", "brought_in", "moved_out", "claim_target"):
            self.summary_cards[key].set_values(money_summary(breakdown[key]), signed=key == "moved_out")

        claim_rows = breakdown["claim_target"]
        self.summary_models["item"].set_rows(grouped_summary(claim_rows, lambda row: row.item))
        self.summary_models["vendor"].set_rows(grouped_summary(claim_rows, lambda row: row.vendor))
        self.summary_models["sheet"].set_rows(grouped_summary(claim_rows, lambda row: row.source_sheet or "단일대장"))
        mismatch = sum(1 for row in self.data.rows if row.amount_review_status == "불일치")
        unavailable = sum(1 for row in self.data.rows if row.amount_review_status == "검토불가")
        self.validation_label.setText(f"금액 검토: 불일치 {mismatch:,}건 · 검토불가 {unavailable:,}건")
        self.status_label.setText(f"조회 완료: 실제 청구대상 {len(claim_rows):,}건")

    @staticmethod
    def _detail_title(key: str) -> str:
        return {
            "current_intake": "당월 입고분",
            "brought_in": "전월 이월분",
            "moved_out": "다음 달 이월분",
            "claim_target": "실제 청구대상",
        }[key]

    # ------------------------------------------------------------------
    # Selection and bulk operations
    # ------------------------------------------------------------------
    def _selected_rows(self) -> list[LedgerRow]:
        current = self.tabs.currentIndex()
        if current < 0 or current >= 4:
            return []
        key = list(self.detail_views)[current]
        view = self.detail_views[key]
        model = self.detail_models[key]
        indices = sorted({index.row() for index in view.selectionModel().selectedRows()})
        return [model.rows[index] for index in indices if 0 <= index < len(model.rows)]

    def _require_selected(self, message: str) -> list[LedgerRow]:
        rows = self._selected_rows()
        if not rows:
            QMessageBox.information(self, "선택 필요", message)
        return rows

    def _current_month_key(self) -> Optional[str]:
        year, month = _period_from_controls(self.year_combo, self.month_combo)
        if year is None or month is None:
            QMessageBox.information(self, "기간 선택", "이월 작업 전 연도와 월을 각각 선택해 주세요.")
            return None
        return f"{year:04d}-{month:02d}"

    def _confirm_rollover(self, title: str, rows: list[LedgerRow], old_month: str, new_month: str) -> bool:
        summary = money_summary(rows)
        text = (
            f"선택 건수: {summary.count:,}건\n"
            f"공급가액: {format_money(summary.supply)} 원\n"
            f"부가세: {format_money(summary.vat)} 원\n"
            f"합계: {format_money(summary.total)} 원\n\n"
            f"변경 전: {old_month}\n변경 후: {new_month}\n\n계속하시겠습니까?"
        )
        return QMessageBox.question(self, title, text) == QMessageBox.StandardButton.Yes

    def move_selected_next(self) -> None:
        if not self.store:
            return
        period = self._current_month_key()
        if not period:
            return
        rows = self._require_selected("다음 달로 이월할 상세내역 행을 선택해 주세요.")
        if not rows:
            return
        eligible = [
            row for row in rows
            if effective_claim_month(row, self.store) == period
            and not is_excluded(row, self.store)
        ]
        skipped = len(rows) - len(eligible)
        if not eligible:
            QMessageBox.information(self, "이월 불가", "선택 행은 현재 조회월의 청구대상이 아니거나 이미 이월되었거나 청구 제외 상태입니다.")
            return
        target = shift_month(period, 1)
        if not self._confirm_rollover("다음 달 이월 확인", eligible, period, target):
            return
        set_claim_month(eligible, self.store, target, reason="다음 달로 이월")
        self.apply_filters()
        QMessageBox.information(self, "이월 완료", f"처리 {len(eligible):,}건 · 제외 {skipped:,}건")

    def import_previous(self) -> None:
        if not self.data or not self.store:
            return
        period = self._current_month_key()
        if not period:
            return
        previous = shift_month(period, -1)
        sources = self.source_panel.selected_values()
        vendors = self.vendor_panel.selected_values()
        items = self.item_panel.selected_values()
        trade = self.trade_combo.currentText() or "전체"
        candidates: list[LedgerRow] = []
        for row in self.data.rows:
            if row.intake_month != previous:
                continue
            if effective_claim_month(row, self.store) != previous:
                continue
            if is_excluded(row, self.store):
                continue
            if sources and row.source_sheet not in sources:
                continue
            if vendors and row.vendor not in vendors:
                continue
            if items and row.item not in items:
                continue
            if trade not in ("", "전체") and effective_trade(row, self.store) != trade:
                continue
            candidates.append(row)
        if not candidates:
            QMessageBox.information(self, "가져오기", f"{previous}에서 가져올 수 있는 미이월 행이 없습니다.")
            return
        dialog = RowSelectionDialog(f"{previous} → {period} 가져오기", candidates, self.store, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        selected = dialog.selected_rows()
        if not selected:
            QMessageBox.information(self, "선택 필요", "가져올 행을 선택해 주세요.")
            return
        if not self._confirm_rollover("이전 달 가져오기 확인", selected, previous, period):
            return
        set_claim_month(selected, self.store, period, reason="이전 달에서 가져오기")
        self.apply_filters()
        QMessageBox.information(self, "가져오기 완료", f"{len(selected):,}건을 {period} 청구대상으로 가져왔습니다.")

    def cancel_selected_rollover(self) -> None:
        if not self.store:
            return
        rows = self._require_selected("이월을 취소할 상세내역 행을 선택해 주세요.")
        if not rows:
            return
        eligible = [row for row in rows if effective_claim_month(row, self.store) != row.intake_month]
        if not eligible:
            QMessageBox.information(self, "이월 취소", "선택 행에 이월된 내역이 없습니다.")
            return
        summary = money_summary(eligible)
        if QMessageBox.question(
            self,
            "이월 취소 확인",
            f"선택 {summary.count:,}건, 공급가액 {format_money(summary.supply)}원을 원래 입고월로 복구할까요?",
        ) != QMessageBox.StandardButton.Yes:
            return
        cancel_rollover(eligible, self.store, reason="이월 취소")
        self.apply_filters()

    def change_selected_classification(self) -> None:
        if not self.store:
            return
        rows = self._require_selected("청구분류를 변경할 행을 선택해 주세요.")
        if not rows:
            return
        value, ok = QInputDialog.getText(self, "청구분류 변경", "변경할 청구분류")
        if ok and value.strip():
            set_classification(rows, self.store, value.strip(), reason="수동 분류 변경")
            self._populate_trade_options_preserve()
            self.apply_filters()

    def _populate_trade_options_preserve(self) -> None:
        if not self.data or not self.store:
            return
        current = self.trade_combo.currentText()
        trades = sorted({effective_trade(row, self.store) for row in self.data.rows if effective_trade(row, self.store)}, key=str.casefold)
        self.trade_combo.clear()
        self.trade_combo.addItem("전체")
        self.trade_combo.addItems(trades)
        index = self.trade_combo.findText(current)
        self.trade_combo.setCurrentIndex(index if index >= 0 else 0)

    def reset_selected_classification(self) -> None:
        if not self.store:
            return
        rows = self._require_selected("분류를 원복할 행을 선택해 주세요.")
        if rows:
            reset_classification(rows, self.store, reason="분류 원복")
            self._populate_trade_options_preserve()
            self.apply_filters()

    def exclude_selected(self) -> None:
        if not self.store:
            return
        rows = self._require_selected("청구대상에서 제외할 행을 선택해 주세요.")
        if not rows:
            return
        reason, ok = QInputDialog.getText(self, "청구대상 제외", "제외 사유")
        if ok:
            exclude_rows(rows, self.store, reason=reason)
            self.apply_filters()

    def restore_selected_excluded(self) -> None:
        if not self.store:
            return
        rows = self._require_selected("제외를 취소할 행을 선택해 주세요.")
        if rows:
            restore_excluded_rows(rows, self.store, reason="제외 취소")
            self.apply_filters()

    def edit_management_note(self) -> None:
        if not self.store:
            return
        rows = self._require_selected("관리메모를 입력할 행을 선택해 주세요.")
        if not rows:
            return
        initial = self.store.management_note(rows[0].fingerprint) if len(rows) == 1 else ""
        note, ok = QInputDialog.getMultiLineText(self, "관리메모", f"선택 {len(rows):,}건에 적용할 메모", initial)
        if not ok:
            return
        for row in rows:
            self.store.set_management_note(
                row.fingerprint,
                note,
                row_info=row_info(row),
                autosave=False,
            )
        self.store.save()
        self.apply_filters()

    # ------------------------------------------------------------------
    # Export, presets, backups
    # ------------------------------------------------------------------
    def export_current(self) -> None:
        if not self.data or not self.store:
            QMessageBox.information(self, "안내", "먼저 Excel 파일을 불러와 주세요.")
            return
        rows = self.current_breakdown.get("claim_target", [])
        if not rows:
            QMessageBox.information(self, "내보내기", "현재 조건의 실제 청구대상이 없습니다.")
            return
        year, month = _period_from_controls(self.year_combo, self.month_combo)
        suffix = f"_{year:04d}-{month:02d}" if year and month else "_전체"
        initial = Path(self.data.path).with_name(f"{Path(self.data.path).stem}_청구대상{suffix}.xlsx")
        output, _ = QFileDialog.getSaveFileName(self, "조회결과 Excel로 내보내기", str(initial), "Excel 통합문서 (*.xlsx)")
        if not output:
            return
        try:
            export_rows(output, rows, self.store, site_name=self.data.site_name)
        except Exception as exc:
            QMessageBox.critical(self, "내보내기 실패", str(exc))
            return
        QMessageBox.information(self, "저장 완료", f"Excel 파일을 저장했습니다.\n{output}")

    def _preset_payload(self) -> dict[str, object]:
        year, month = _period_from_controls(self.year_combo, self.month_combo)
        return {
            "year": year,
            "month": month,
            "trade": self.trade_combo.currentText(),
            "source_sheets": sorted(self.source_panel.selected_values()),
            "vendors": sorted(self.vendor_panel.selected_values()),
            "items": sorted(self.item_panel.selected_values()),
            "keyword": self.keyword_edit.text(),
            "mismatch_only": self.mismatch_check.isChecked(),
        }

    def _refresh_presets(self) -> None:
        current = self.preset_combo.currentText()
        self.preset_combo.clear()
        self.preset_combo.addItems(self.preset_store.names())
        index = self.preset_combo.findText(current)
        if index >= 0:
            self.preset_combo.setCurrentIndex(index)

    def save_preset(self) -> None:
        name, ok = QInputDialog.getText(self, "조회 프리셋 저장", "프리셋 이름")
        name = name.strip()
        if not ok or not name:
            return
        if self.preset_store.get(name) and QMessageBox.question(self, "덮어쓰기", f"'{name}' 프리셋을 덮어쓸까요?") != QMessageBox.StandardButton.Yes:
            return
        self.preset_store.set(name, self._preset_payload())
        self._refresh_presets()
        self.preset_combo.setCurrentText(name)

    def update_preset(self) -> None:
        name = self.preset_combo.currentText()
        if not name:
            QMessageBox.information(self, "프리셋", "수정할 프리셋을 선택해 주세요.")
            return
        self.preset_store.set(name, self._preset_payload())
        QMessageBox.information(self, "프리셋", f"'{name}' 프리셋을 수정했습니다.")

    def delete_preset(self) -> None:
        name = self.preset_combo.currentText()
        if name and QMessageBox.question(self, "프리셋 삭제", f"'{name}' 프리셋을 삭제할까요?") == QMessageBox.StandardButton.Yes:
            self.preset_store.delete(name)
            self._refresh_presets()

    def apply_preset(self) -> None:
        name = self.preset_combo.currentText()
        payload = self.preset_store.get(name) if name else None
        if not payload:
            QMessageBox.information(self, "프리셋", "적용할 프리셋을 선택해 주세요.")
            return
        if payload.get("year") is not None:
            index = self.year_combo.findData(int(payload["year"]))
            self.year_combo.setCurrentIndex(index if index >= 0 else 0)
        else:
            self.year_combo.setCurrentIndex(0)
        if payload.get("month") is not None:
            index = self.month_combo.findData(int(payload["month"]))
            self.month_combo.setCurrentIndex(index if index >= 0 else 0)
        else:
            self.month_combo.setCurrentIndex(0)
        trade = str(payload.get("trade", "전체"))
        index = self.trade_combo.findText(trade)
        self.trade_combo.setCurrentIndex(index if index >= 0 else 0)
        self.source_panel.set_selected(payload.get("source_sheets", []))
        self.vendor_panel.set_selected(payload.get("vendors", []))
        self.item_panel.set_selected(payload.get("items", []))
        self.keyword_edit.setText(str(payload.get("keyword", "")))
        self.mismatch_check.setChecked(bool(payload.get("mismatch_only", False)))
        self.apply_filters()

    def manual_backup(self) -> None:
        if not self.store:
            QMessageBox.information(self, "백업", "먼저 Excel 파일을 불러와 주세요.")
            return
        initial = self.store.source_path.with_name(f"{self.store.source_path.stem}_청구관리_백업.json")
        output, _ = QFileDialog.getSaveFileName(self, "관리 JSON 백업", str(initial), "JSON 파일 (*.json)")
        if output:
            try:
                result = self.store.create_manual_backup(output)
                QMessageBox.information(self, "백업 완료", str(result))
            except Exception as exc:
                QMessageBox.critical(self, "백업 실패", str(exc))

    def restore_backup(self) -> None:
        if not self.store:
            QMessageBox.information(self, "복원", "먼저 원본 Excel 파일을 불러와 주세요.")
            return
        source, _ = QFileDialog.getOpenFileName(self, "관리 JSON 복원", str(self.store.source_path.parent), "JSON 파일 (*.json)")
        if not source:
            return
        if QMessageBox.question(self, "복원 확인", "현재 관리값을 백업한 뒤 선택한 JSON으로 복원합니다. 계속할까요?") != QMessageBox.StandardButton.Yes:
            return
        try:
            self.store.restore_backup(source)
            self._populate_trade_options_preserve()
            self.apply_filters()
            QMessageBox.information(self, "복원 완료", "관리 데이터를 복원했습니다.")
        except Exception as exc:
            QMessageBox.critical(self, "복원 실패", str(exc))

    def show_history(self) -> None:
        if not self.store:
            QMessageBox.information(self, "이력", "먼저 Excel 파일을 불러와 주세요.")
            return
        HistoryDialog(self.store.history, self).exec()

    def open_original(self) -> None:
        if not self.data:
            QMessageBox.information(self, "원본", "먼저 Excel 파일을 불러와 주세요.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(self.data.path))

    def return_home(self) -> None:
        if self.home_callback:
            self.home_callback(self)
        else:
            self.close()

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            * { font-family: 'Pretendard', 'Malgun Gothic', sans-serif; font-size: 12px; color: #172033; }
            QWidget#claimRoot, QWidget#claimWorkspace { background: #FFFFFF; }
            QFrame#sidebar { background: #122B48; }
            QLabel#brandTitle { color: white; font-size: 16px; font-weight: 800; }
            QLabel#brandSubtitle { color: #B7CBE2; font-size: 11px; }
            QLabel#sidebarSection { color: #8FA9C6; font-size: 11px; font-weight: 800; padding-top: 8px; }
            QPushButton#navButton, QPushButton#sidebarFooterButton { min-height: 38px; border: 0; border-radius: 8px; padding: 0 10px; text-align: left; background: transparent; color: #E8F0FA; font-weight: 600; }
            QPushButton#navButton:hover, QPushButton#sidebarFooterButton:hover { background: #1C4775; }
            QFrame#sidebarDivider { color: #31506F; }
            QLabel#versionLabel { color: #AFC4DD; font-size: 11px; }
            QFrame#topbar { background: white; border-bottom: 1px solid #DDE4EE; }
            QLabel#pageTitle { font-size: 20px; font-weight: 900; }
            QLabel#filePathLabel { color: #65738A; padding-left: 8px; }
            QFrame#card, QFrame#summaryCard, QFrame#summaryCardStrong, QFrame#noticeCard, QFrame#filterSubCard { background: #FFFFFF; border: 1px solid #E1E7F0; border-radius: 12px; }
            QFrame#summaryCardStrong { border: 2px solid #9A7BE8; }
            QLabel#cardTitle { font-size: 15px; font-weight: 850; color: #1F2A3D; }
            QLabel#subCardTitle, QLabel#summaryTitle, QLabel#summaryTitleStrong { font-weight: 800; }
            QLabel#summaryTitleStrong { color: #7C3AED; }
            QLabel#summarySubtitle, QLabel#selectionCount { color: #758399; font-size: 10px; }
            QLineEdit, QComboBox, QListWidget, QTableView, QTableWidget { background: white; border: 1px solid #CDD6E3; border-radius: 7px; padding: 5px; selection-background-color: #DDEAFF; selection-color: #172033; }
            QLineEdit:focus, QComboBox:focus, QListWidget:focus, QTableView:focus { border: 1px solid #4A78C2; }
            QTableView, QTableWidget { border-radius: 8px; gridline-color: #E7EBF1; alternate-background-color: #FAFBFD; }
            QHeaderView::section { background: #EEF2F7; border: 0; border-right: 1px solid #DDE4EE; border-bottom: 1px solid #DDE4EE; padding: 7px; font-weight: 800; }
            QTabWidget::pane { border: 1px solid #DCE3ED; background: white; border-radius: 9px; }
            QTabBar::tab { background: #E9EEF5; padding: 9px 14px; margin-right: 2px; border-top-left-radius: 7px; border-top-right-radius: 7px; }
            QTabBar::tab:selected { background: white; font-weight: 900; border-bottom: 3px solid #4F7CC4; }
            QSplitter#upperSplitter::handle { background: transparent; }
            QPushButton { min-height: 34px; border: 1px solid #C8D2E0; border-radius: 8px; padding: 0 12px; background: white; font-weight: 650; }
            QPushButton:hover { background: #F2F5FA; }
            QPushButton#primaryButton { background: #2867B2; color: white; border: 0; }
            QPushButton#primaryButton:hover { background: #215995; }
            QPushButton#purpleButton { background: #6D4BC3; color: white; border: 0; }
            QPushButton#purpleButton:hover { background: #5A3FA2; }
            QPushButton#blueAction { background: #E8F1FF; color: #245FAE; border-color: #A9C8F2; }
            QPushButton#orangeAction { background: #FFF0DD; color: #A65300; border-color: #F0BE7D; }
            QPushButton#dangerButton { background: #FFF0F1; color: #B42318; border-color: #F0B5B9; }
            QPushButton#secondaryButton, QPushButton#topActionButton, QPushButton#smallButton { background: white; }
            QPushButton#smallButton { min-height: 28px; padding: 0 8px; font-size: 11px; }
            QCheckBox { spacing: 6px; }
            """
        )


def run() -> None:
    app = QApplication.instance() or QApplication([])
    window = MaterialClaimWindow()
    window.show()
    app.exec()
