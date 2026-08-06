from __future__ import annotations

import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
pytest.importorskip("PySide6")
from PySide6.QtWidgets import QApplication

from material_claim_manager.ui import MaterialClaimWindow


def test_claim_window_builds_all_tabs_and_actions() -> None:
    app = QApplication.instance() or QApplication([])
    window = MaterialClaimWindow()
    try:
        assert window.windowTitle().endswith("v2.0.0")
        assert window.tabs.count() == 7
        assert set(window.detail_models) == {"current_intake", "brought_in", "moved_out", "claim_target"}
        assert set(window.summary_models) == {"item", "vendor", "sheet"}
        assert window.load_button.text() == "엑셀 파일 불러오기"
        assert window.export_button.text() == "조회결과 Excel로 내보내기"
        assert window.minimumWidth() >= 1280
        assert window.minimumHeight() >= 760
    finally:
        window.close()
        app.processEvents()
