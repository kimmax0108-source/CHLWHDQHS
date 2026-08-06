from __future__ import annotations

import sys

from PySide6.QtCore import QCoreApplication, QLocale
from PySide6.QtWidgets import QApplication

from . import __version__
from .ui import APP_TITLE, MainWindow


def main() -> int:
    QCoreApplication.setOrganizationName("YangwooTools")
    QCoreApplication.setApplicationName("MaterialDocumentStandardizationPurchase")
    QCoreApplication.setApplicationVersion(__version__)
    QLocale.setDefault(QLocale(QLocale.Language.Korean, QLocale.Country.SouthKorea))

    app = QApplication(sys.argv)
    app.setApplicationDisplayName(APP_TITLE)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
