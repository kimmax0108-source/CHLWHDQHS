from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from .ui import MainWindow


def run() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("MaterialDocumentStandardizationExpense")
    app.setOrganizationName("Yangwoo")
    window = MainWindow()
    window.show()
    return app.exec()
