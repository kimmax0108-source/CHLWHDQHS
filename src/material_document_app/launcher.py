from __future__ import annotations

from typing import Callable

from PySide6.QtCore import Qt
from PySide6.QtGui import QCloseEvent, QIcon, QMouseEvent
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from expense_statement_app.ui import MainWindow as ExpenseWindow
from material_claim_manager.ui import MaterialClaimWindow
from purchase_request_app.ui import MainWindow as PurchaseWindow

from . import __version__
from .resource import resource_path

APP_TITLE = "자재 문서 표준화"


class ClickableCard(QFrame):
    def __init__(self, callback: Callable[[], None], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._callback = callback
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setObjectName("startCard")

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:  # noqa: N802
        if event.button() == Qt.MouseButton.LeftButton and self.rect().contains(event.position().toPoint()):
            self._callback()
        super().mouseReleaseEvent(event)


class LauncherWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.child_window: QMainWindow | None = None
        self.setWindowTitle(f"{APP_TITLE} v{__version__}")
        self.resize(1380, 820)
        self.setMinimumSize(1080, 690)
        icon_path = resource_path("assets/app.ico")
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        self._build_ui()
        self._apply_style()

    def _icon(self, name: str) -> QIcon:
        path = resource_path(f"assets/{name}.svg")
        return QIcon(str(path)) if path.exists() else QIcon()

    def _nav_button(self, label: str, icon: str, callback: Callable[[], None]) -> QPushButton:
        button = QPushButton(self._icon(icon), label)
        button.setObjectName("launcherNav")
        button.clicked.connect(callback)
        return button

    def _build_ui(self) -> None:
        central = QWidget()
        central.setObjectName("launcherRoot")
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        self.setCentralWidget(central)

        sidebar = QFrame()
        sidebar.setObjectName("launcherSidebar")
        sidebar.setFixedWidth(260)
        side = QVBoxLayout(sidebar)
        side.setContentsMargins(20, 25, 20, 20)
        side.setSpacing(8)

        brand = QHBoxLayout()
        logo = QLabel()
        logo.setPixmap(self._icon("appmark").pixmap(44, 44))
        brand_text = QVBoxLayout()
        title = QLabel("자재 문서 표준화")
        title.setObjectName("launcherBrandTitle")
        version = QLabel(f"v{__version__}")
        version.setObjectName("launcherVersion")
        brand_text.addWidget(title)
        brand_text.addWidget(version)
        brand.addWidget(logo)
        brand.addLayout(brand_text, 1)
        side.addLayout(brand)
        side.addSpacing(15)

        section1 = QLabel("구매 품의서 양식")
        section1.setObjectName("sidebarSection")
        side.addWidget(section1)
        side.addWidget(self._nav_button("견적대비표 작성", "quote", lambda: self.open_purchase(0)))
        side.addWidget(self._nav_button("내역서 작성", "statement", lambda: self.open_purchase(1)))
        side.addWidget(self._nav_button("구매품의서 작성", "purchase", lambda: self.open_purchase(2)))

        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setObjectName("launcherDivider")
        side.addWidget(divider)

        section2 = QLabel("지출결의서")
        section2.setObjectName("sidebarSection")
        side.addWidget(section2)
        side.addWidget(self._nav_button("지출결의서 작성", "expense", self.open_expense))

        section3 = QLabel("자재 입고 청구관리")
        section3.setObjectName("sidebarSection")
        side.addWidget(section3)
        side.addWidget(self._nav_button("자재 입고 청구관리", "database", self.open_claim))

        side.addStretch(1)
        settings = QPushButton("⚙  프로그램 설정")
        settings.setObjectName("launcherFooterNav")
        settings.clicked.connect(self.show_settings)
        about = QPushButton("ⓘ  프로그램 정보")
        about.setObjectName("launcherFooterNav")
        about.clicked.connect(self.show_about)
        side.addWidget(settings)
        side.addWidget(about)
        root.addWidget(sidebar)

        content = QWidget()
        content.setObjectName("launcherContent")
        body = QVBoxLayout(content)
        body.setContentsMargins(46, 44, 46, 36)
        body.setSpacing(18)

        heading = QLabel("자재 문서 표준화 프로그램")
        heading.setObjectName("launcherHeading")
        subheading = QLabel("구매 문서 작성, 지출결의, 자재 청구 검토를 하나의 프로그램에서 관리합니다.")
        subheading.setObjectName("launcherSubheading")
        body.addWidget(heading)
        body.addWidget(subheading)
        body.addSpacing(8)

        cards = QHBoxLayout()
        cards.setSpacing(18)
        cards.addWidget(
            self._build_card(
                title="구매 품의서 양식",
                lines=("견적대비표 → 내역서 → 구매품의서", "3개 문서를 작성하고 하나의 Excel 파일로 저장합니다."),
                icon_name="purchase_set",
                button_text="작성 시작",
                accent="blue",
                callback=lambda: self.open_purchase(0),
            ),
            1,
        )
        cards.addWidget(
            self._build_card(
                title="지출결의서 작성",
                lines=("지출결의서를 Excel로 저장합니다.",),
                icon_name="expense_large",
                button_text="작성 시작",
                accent="green",
                callback=self.open_expense,
            ),
            1,
        )
        cards.addWidget(
            self._build_card(
                title="자재 입고 청구관리",
                lines=("자재 입고대장을 불러와", "청구금액·분류·제외·이월내역을 검토합니다."),
                icon_name="database",
                button_text="검토 시작",
                accent="purple",
                callback=self.open_claim,
            ),
            1,
        )
        body.addLayout(cards, 1)

        lower = QHBoxLayout()
        program_info = QFrame()
        program_info.setObjectName("programInfo")
        info_layout = QVBoxLayout(program_info)
        info_layout.setContentsMargins(20, 15, 20, 15)
        info_title = QLabel("프로그램 정보")
        info_title.setObjectName("programInfoTitle")
        info_layout.addWidget(info_title)
        info_layout.addWidget(QLabel(f"• 통합 버전 : {__version__}"))
        info_layout.addWidget(QLabel("• 원본 Excel 읽기 전용 및 별도 관리 JSON 저장"))
        info_layout.addWidget(QLabel("• GitHub Actions 테스트·Windows EXE 빌드 지원"))
        lower.addWidget(program_info, 1)

        recent = QFrame()
        recent.setObjectName("programInfo")
        recent_layout = QVBoxLayout(recent)
        recent_layout.setContentsMargins(20, 15, 20, 15)
        recent_title = QLabel("업무 구성")
        recent_title.setObjectName("programInfoTitle")
        recent_layout.addWidget(recent_title)
        recent_layout.addWidget(QLabel("구매 품의서 3개 시트 · 지출결의서 · 자재 청구관리"))
        recent_layout.addWidget(QLabel("각 업무는 독립 저장되며 계산 기준과 문서 양식을 유지합니다."))
        lower.addWidget(recent, 1)
        body.addLayout(lower)
        root.addWidget(content, 1)

    def _build_card(
        self,
        *,
        title: str,
        lines: tuple[str, ...],
        icon_name: str,
        button_text: str,
        accent: str,
        callback: Callable[[], None],
    ) -> ClickableCard:
        card = ClickableCard(callback)
        card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(24, 23, 24, 23)
        layout.setSpacing(9)
        title_label = QLabel(title)
        title_label.setObjectName("startCardTitle")
        layout.addWidget(title_label)
        for line in lines:
            text = QLabel(line)
            text.setWordWrap(True)
            text.setObjectName("startCardText")
            layout.addWidget(text)
        layout.addStretch(1)
        art = QLabel()
        art.setAlignment(Qt.AlignmentFlag.AlignCenter)
        art.setPixmap(self._icon(icon_name).pixmap(140, 110))
        layout.addWidget(art)
        button = QPushButton(button_text)
        button.setObjectName(
            {"blue": "startBlueButton", "green": "startGreenButton", "purple": "startPurpleButton"}[accent]
        )
        button.clicked.connect(callback)
        button.setCursor(Qt.CursorShape.PointingHandCursor)
        layout.addWidget(button, alignment=Qt.AlignmentFlag.AlignHCenter)
        return card

    def _show_child(self, child: QMainWindow) -> None:
        child.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        child.destroyed.connect(self._clear_child)
        self.child_window = child
        self.hide()
        child.show()
        child.raise_()
        child.activateWindow()

    def open_purchase(self, page_index: int = 0) -> None:
        if isinstance(self.child_window, PurchaseWindow):
            self.child_window.set_workspace_page(page_index)
            self.child_window.raise_()
            return
        if self.child_window is not None:
            self.child_window.raise_()
            return
        child = PurchaseWindow(home_callback=self.return_home)
        child.set_workspace_page(page_index)
        self._show_child(child)

    def open_expense(self) -> None:
        if isinstance(self.child_window, ExpenseWindow):
            self.child_window.raise_()
            return
        if self.child_window is not None:
            self.child_window.raise_()
            return
        self._show_child(ExpenseWindow(home_callback=self.return_home))

    def open_claim(self) -> None:
        if isinstance(self.child_window, MaterialClaimWindow):
            self.child_window.raise_()
            return
        if self.child_window is not None:
            self.child_window.raise_()
            return
        self._show_child(MaterialClaimWindow(home_callback=self.return_home))

    def return_home(self, child: QMainWindow) -> None:
        child.close()
        self.child_window = None
        self.show()
        self.raise_()
        self.activateWindow()

    def _clear_child(self, *_args: object) -> None:
        self.child_window = None
        if not self.isVisible():
            self.show()

    def show_settings(self) -> None:
        QMessageBox.information(
            self,
            "프로그램 설정",
            "구매 품의서의 Excel 양식은 구매 품의서 화면의 설정에서 변경할 수 있습니다.\n"
            "자재 청구관리 프리셋과 백업은 해당 화면에서 관리합니다.",
        )

    def show_about(self) -> None:
        QMessageBox.information(
            self,
            "프로그램 정보",
            f"{APP_TITLE} v{__version__}\n\n"
            "구매 품의서 양식, 지출결의서, 자재 입고 청구관리 통합 프로그램\n"
            "원본 문서 보존 · 정밀 계산 · Excel 내보내기 · 관리 JSON 백업",
        )

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802
        if self.child_window is not None:
            self.child_window.close()
            self.child_window = None
        event.accept()
        QApplication.quit()

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            * { font-family: 'Pretendard', 'Malgun Gothic', sans-serif; font-size: 12px; color: #172033; }
            QWidget#launcherRoot, QWidget#launcherContent { background: #F8FAFD; }
            QFrame#launcherSidebar { background: #122B48; }
            QLabel#launcherBrandTitle { color: white; font-size: 17px; font-weight: 800; }
            QLabel#launcherVersion { color: #AFC4DD; font-size: 11px; }
            QLabel#sidebarSection { color: #8FA9C6; font-size: 11px; font-weight: 800; padding-top: 6px; }
            QPushButton#launcherNav, QPushButton#launcherFooterNav { min-height: 38px; border: 0; border-radius: 8px; padding: 0 10px; text-align: left; background: transparent; color: #E8F0FA; font-weight: 600; }
            QPushButton#launcherNav:hover, QPushButton#launcherFooterNav:hover { background: #1C4775; }
            QFrame#launcherDivider { color: #31506F; }
            QLabel#launcherHeading { font-size: 27px; font-weight: 900; color: #111827; }
            QLabel#launcherSubheading { font-size: 13px; color: #637083; }
            QFrame#startCard { background: white; border: 1px solid #E2E8F2; border-radius: 18px; }
            QFrame#startCard:hover { border: 1px solid #9FBCE5; background: #FCFDFF; }
            QLabel#startCardTitle { font-size: 19px; font-weight: 850; color: #15243A; }
            QLabel#startCardText { color: #5D6B7D; }
            QPushButton#startBlueButton, QPushButton#startGreenButton, QPushButton#startPurpleButton { min-width: 120px; min-height: 39px; border: 0; border-radius: 10px; color: white; font-weight: 800; }
            QPushButton#startBlueButton { background: #2867B2; }
            QPushButton#startGreenButton { background: #198754; }
            QPushButton#startPurpleButton { background: #6D4BC3; }
            QPushButton#startBlueButton:hover { background: #215995; }
            QPushButton#startGreenButton:hover { background: #146C43; }
            QPushButton#startPurpleButton:hover { background: #5A3FA2; }
            QFrame#programInfo { background: white; border: 1px solid #E2E8F2; border-radius: 14px; }
            QLabel#programInfoTitle { font-weight: 850; font-size: 14px; }
            """
        )
