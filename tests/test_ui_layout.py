from __future__ import annotations

import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
PySide6 = pytest.importorskip("PySide6")
from PySide6.QtWidgets import QAbstractItemView, QApplication, QLabel

from purchase_request_app.ui import MainWindow
from expense_statement_app.ui import MainWindow as ExpenseMainWindow


def test_v2_workspace_order_and_flexible_tables() -> None:
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    try:
        assert window.tabs.count() == 4
        assert window.tabs.currentIndex() == 0
        assert [button.text() for button in window.nav_buttons] == [
            "견적대비표 작성",
            "내역서 작성",
            "구매품의서 작성",
        ]
        assert window.preview_sheet_combo.currentData() == "quote"

        # Only the two editable tables are resizable. The ranking/selection area is fixed below.
        assert not window.quote_splitter.childrenCollapsible()
        assert window.quote_splitter.count() == 2
        assert window.quote_splitter.indexOf(window.summary_group) == -1
        assert window.quote_splitter.handleWidth() >= 10
        assert window.vendor_group.minimumHeight() <= 50
        assert window.item_group.minimumHeight() <= 60
        assert window.summary_group.minimumHeight() == 248
        assert window.summary_group.maximumHeight() == 248

        # Preview width must be visibly and actually draggable.
        assert window.main_splitter.count() == 2
        assert not window.main_splitter.childrenCollapsible()
        assert window.main_splitter.handleWidth() >= 10
        assert window.main_splitter.widget(1).minimumWidth() >= 350

        assert window.default_layout_button.text() == "기본 배치"
        assert window.export_button.text() == "구매품의서 Excel로 내보내기"
        assert not hasattr(window, "output_nav_button")
        assert window.right_splitter.orientation().name == "Vertical"
        assert not window.right_splitter.childrenCollapsible()
        assert window.overview_card.maximumHeight() <= 196
        assert window.overview_title_label.text() == "문서 요약"
        assert window.overview_step_label.text().startswith("1 / 3")

        # Section titles must be inside their white cards, not floating on the border.
        card_titles = [label.text() for label in window.findChildren(QLabel, "sectionCardTitle")]
        for expected in (
            "견적 기본정보",
            "1. 견적업체",
            "2. 대표 품목명 및 상세 품목별 업체 단가",
            "3. 업체별 순위 및 최종선정",
            "1. 내역서 기본정보",
            "2. 구매물품 내역",
            "1. 문서 기본정보",
            "5. 본문 및 추가 문구",
        ):
            assert expected in card_titles

        assert window.selected_vendor_supply_label.minimumWidth() >= 100
        assert window.budget_label.minimumWidth() >= 110
        assert window.budget_mode_combo.count() == 2
        assert window.won_rounding_combo.count() == 4
        assert not window.purchase_site_edit.isReadOnly()
        assert not window.vendor_edit.isReadOnly()
        assert not window.phone_edit.isReadOnly()
        assert window.vendor_table.selectionBehavior() == QAbstractItemView.SelectionBehavior.SelectItems
        assert window.item_table.selectionBehavior() == QAbstractItemView.SelectionBehavior.SelectItems
        assert window.statement_table.selectionBehavior() == QAbstractItemView.SelectionBehavior.SelectItems
        assert window.vendor_table.horizontalHeader().sectionsMovable()
        assert window.item_table.horizontalHeader().sectionsMovable()
        assert window.statement_table.horizontalHeader().sectionsMovable()
        assert window.minimumWidth() >= 1240
        assert window.minimumHeight() >= 860
    finally:
        window.close()
        app.processEvents()


def test_expense_sections_are_aligned_cards_and_preview_is_resizable() -> None:
    app = QApplication.instance() or QApplication([])
    window = ExpenseMainWindow()
    try:
        assert window.main_splitter.count() == 2
        assert not window.main_splitter.childrenCollapsible()
        assert window.main_splitter.handleWidth() >= 10
        assert window.main_splitter.widget(0).minimumWidth() >= 590
        assert window.main_splitter.widget(1).minimumWidth() >= 420
        titles = {label.text() for label in window.findChildren(QLabel, "sectionCardTitle")}
        assert titles == {
            "1. 기본 정보",
            "2. 원가 안분 내역 (현장/업체별)",
            "3. 적요 / 구매품목 (자유 입력)",
            "4. 업체 계좌 정보",
            "5. 미리보기",
        }
        assert window.allocation_table.minimumHeight() >= 230
        assert window.account_table.minimumHeight() >= 140
    finally:
        window.close()
        app.processEvents()
