from __future__ import annotations

import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
pytest.importorskip("PySide6")
from PySide6.QtWidgets import QApplication

from material_document_app.launcher import LauncherWindow


def test_launcher_contains_three_independent_workflows() -> None:
    app = QApplication.instance() or QApplication([])
    window = LauncherWindow()
    try:
        assert window.windowTitle().startswith("자재 문서 표준화")
        assert window.minimumWidth() >= 900
        assert window.minimumHeight() >= 600
        assert hasattr(window, "open_purchase")
        assert hasattr(window, "open_expense")
        assert hasattr(window, "open_claim")
    finally:
        window.close()
        app.processEvents()
