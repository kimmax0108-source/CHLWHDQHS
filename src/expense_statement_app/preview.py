from __future__ import annotations

from decimal import Decimal

from PySide6.QtCore import QRectF, QSize, Qt
from PySide6.QtGui import QColor, QFont, QPainter, QPainterPath, QPen
from PySide6.QtWidgets import QSizePolicy, QWidget

from .calculations import financial_amount_text, quantity_parts
from .models import ExpenseDocument, format_money


class ExpensePreviewWidget(QWidget):
    PAGE_W = 520
    BASE_PAGE_H = 735
    ROW_H = 39

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.document = ExpenseDocument()
        self.zoom = 1.0
        self.page_h = self.BASE_PAGE_H
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self._update_size()

    def set_document(self, document: ExpenseDocument) -> None:
        self.document = document
        body_rows = max(7, len(document.export_body_lines()))
        self.page_h = self.BASE_PAGE_H + max(0, body_rows - 7) * self.ROW_H
        self._update_size()
        self.update()

    def set_zoom(self, zoom: float) -> None:
        self.zoom = max(0.55, min(1.8, zoom))
        self._update_size()
        self.update()

    def _update_size(self) -> None:
        self.setFixedSize(int(self.PAGE_W * self.zoom + 24), int(self.page_h * self.zoom + 24))

    def sizeHint(self) -> QSize:
        return QSize(int(self.PAGE_W * self.zoom + 24), int(self.page_h * self.zoom + 24))

    def paintEvent(self, event) -> None:  # noqa: ANN001
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.fillRect(self.rect(), QColor("#eef2f7"))
        painter.save()
        painter.translate(12, 12)
        painter.scale(self.zoom, self.zoom)
        self._draw_page(painter)
        painter.restore()

    def _font(self, size: int, bold: bool = False, family: str = "맑은 고딕") -> QFont:
        font = QFont(family, size)
        font.setBold(bold)
        return font

    def _line(self, painter: QPainter, x1: float, y1: float, x2: float, y2: float, width=1.0) -> None:
        painter.setPen(QPen(Qt.GlobalColor.black, width))
        painter.drawLine(QRectF(x1, y1, x2 - x1, y2 - y1).topLeft(), QRectF(x1, y1, x2 - x1, y2 - y1).bottomRight())

    def _text(
        self,
        painter: QPainter,
        rect: QRectF,
        text: str,
        size: int = 9,
        bold: bool = False,
        align: Qt.AlignmentFlag = Qt.AlignmentFlag.AlignCenter,
        family: str = "맑은 고딕",
    ) -> None:
        painter.setPen(Qt.GlobalColor.black)
        painter.setFont(self._font(size, bold, family))
        painter.drawText(rect, int(align | Qt.AlignmentFlag.AlignVCenter), text)

    def _draw_rich_quantity(self, painter: QPainter, rect: QRectF, value: Decimal) -> None:
        main, fraction = quantity_parts(value)
        main_font = self._font(8, False, "굴림")
        small_font = self._font(6, False, "굴림")
        painter.setFont(main_font)
        main_width = painter.fontMetrics().horizontalAdvance(main)
        painter.setFont(small_font)
        frac_width = painter.fontMetrics().horizontalAdvance(fraction)
        x = rect.right() - main_width - frac_width - 2
        baseline = rect.center().y() + 3
        painter.setPen(Qt.GlobalColor.black)
        painter.setFont(main_font)
        painter.drawText(QRectF(x, rect.top(), main_width, rect.height()), int(Qt.AlignmentFlag.AlignVCenter), main)
        painter.setFont(small_font)
        painter.drawText(QRectF(x + main_width, rect.top() - 2, frac_width, rect.height() * 0.72), int(Qt.AlignmentFlag.AlignVCenter), fraction)

    def _draw_digits(self, painter: QPainter, rect: QRectF, amount: int) -> None:
        digit_w = rect.width() / 10
        digits = str(amount).rjust(10, "0")[-10:]
        started = False
        for idx, digit in enumerate(digits):
            cell = QRectF(rect.left() + idx * digit_w, rect.top(), digit_w, rect.height())
            painter.drawRect(cell)
            if digit != "0" or started or idx == 9:
                started = True
                self._text(painter, cell, digit, 9)

    def _draw_page(self, p: QPainter) -> None:
        p.fillRect(QRectF(0, 0, self.PAGE_W, self.page_h), Qt.GlobalColor.white)
        p.setPen(QPen(Qt.GlobalColor.black, 1.3))
        p.drawRect(QRectF(4, 4, self.PAGE_W - 8, self.page_h - 8))

        left, right = 22.0, self.PAGE_W - 22.0
        top = 18.0
        title_box = QRectF(left, top, 118, 50)
        p.setPen(QPen(Qt.GlobalColor.black, 1.6))
        p.drawRect(title_box)
        self._text(p, title_box, "원가안분", 21, True)
        self._text(p, QRectF(148, top, 250, 50), "支 出 決 議 書", 20, False, family="굴림")
        self._line(p, 151, 61, 390, 61, 1.3)

        approval_top = 75
        approval_h = 72
        p.drawRect(QRectF(left, approval_top, right - left, approval_h))
        decision_w = 24
        self._text(p, QRectF(left, approval_top, decision_w, approval_h), "決\n裁", 8)
        self._line(p, left + decision_w, approval_top, left + decision_w, approval_top + approval_h)
        labels = ["係", "代 理", "課 長", "次 長", "部 長", "理 事", "常 務", "專 務", "副社長", "社 長"]
        col_w = (right - left - decision_w) / len(labels)
        for idx, label in enumerate(labels):
            x = left + decision_w + idx * col_w
            self._line(p, x, approval_top, x, approval_top + approval_h)
            self._text(p, QRectF(x, approval_top, col_w, 22), label, 7)
            self._line(p, x, approval_top + 22, x + col_w, approval_top + 22)

        amount_top = approval_top + approval_h
        audit_w = 22
        trailing_blank_w = 68
        amount_right = right - audit_w - trailing_blank_w
        audit_left = amount_right
        blank_left = audit_left + audit_w
        p.drawRect(QRectF(left, amount_top, right - left, 66))
        self._line(p, audit_left, amount_top, audit_left, amount_top + 66)
        self._line(p, blank_left, amount_top, blank_left, amount_top + 66)
        self._text(p, QRectF(audit_left, amount_top, audit_w, 33), "監", 10, False, family="굴림")
        self._text(p, QRectF(audit_left, amount_top + 33, audit_w, 33), "事", 10, False, family="굴림")
        amount_text = financial_amount_text(self.document.total_amount)
        self._text(
            p,
            QRectF(left + 12, amount_top + 4, amount_right - left - 24, 30),
            amount_text,
            11,
            False,
            Qt.AlignmentFlag.AlignRight,
            "굴림",
        )
        self._text(
            p,
            QRectF(left + 12, amount_top + 32, amount_right - left - 24, 28),
            f"( ₩ {format_money(self.document.total_amount)} )",
            11,
            False,
            Qt.AlignmentFlag.AlignRight,
        )

        body_top = amount_top + 66
        body_left_w = 310
        amount_left = left + body_left_w
        header_h = 28
        p.drawRect(QRectF(left, body_top, right - left, header_h))
        self._text(p, QRectF(left, body_top, body_left_w, header_h), "摘        要", 10, False, family="굴림")
        self._text(p, QRectF(amount_left, body_top, right - amount_left, header_h), "金        額", 10, False, family="굴림")
        self._line(p, amount_left, body_top, amount_left, body_top + header_h)

        planned_lines = self.document.export_body_lines()
        body_rows = max(7, len(planned_lines))
        row_h = self.ROW_H
        content_top = body_top + header_h
        for row_idx in range(body_rows):
            y = content_top + row_idx * row_h
            p.drawRect(QRectF(left, y, right - left, row_h))
            self._line(p, amount_left, y, amount_left, y + row_h)

        for idx, line in enumerate(planned_lines):
            y = content_top + idx * row_h
            if line.kind == "allocation" and line.allocation is not None:
                allocation = line.allocation
                self._text(
                    p,
                    QRectF(left + 5, y, 190, row_h),
                    allocation.display_site,
                    9,
                    False,
                    Qt.AlignmentFlag.AlignLeft,
                )
                vendor = allocation.vendor_name.strip()
                self._text(
                    p,
                    QRectF(left + 190, y, 75, row_h),
                    (vendor if vendor.startswith("-") else f"- {vendor}") if vendor else "",
                    7,
                    False,
                    Qt.AlignmentFlag.AlignLeft,
                )
                self._draw_rich_quantity(
                    p, QRectF(left + 260, y, 48, row_h), allocation.quantity
                )
                self._draw_digits(
                    p,
                    QRectF(amount_left + 18, y + 1, right - amount_left - 18, row_h - 2),
                    allocation.amount,
                )
            else:
                self._text(
                    p,
                    QRectF(left + 5, y, body_left_w - 10, row_h),
                    line.text,
                    8 if line.kind == "account" else 9,
                    False,
                    Qt.AlignmentFlag.AlignLeft,
                )

        allocation_count = len(self.document.active_allocations)
        if allocation_count < body_rows:
            slash_top = content_top + allocation_count * row_h + 5
            slash_bottom = content_top + body_rows * row_h - 8
            path = QPainterPath()
            path.moveTo(amount_left + 45, slash_top)
            path.cubicTo(right - 18, slash_top - 5, right + 15, slash_top + 25, right - 10, slash_top + 62)
            path.lineTo(amount_left + 48, slash_bottom)
            p.setPen(QPen(Qt.GlobalColor.black, 1.1))
            p.drawPath(path)

        account_top = content_top + body_rows * row_h
        account_h = 34
        p.drawRect(QRectF(left, account_top, right - left, account_h))
        self._text(p, QRectF(left + 300, account_top, 72, account_h), "수량계", 8)
        self._draw_rich_quantity(
            p, QRectF(left + 370, account_top, 75, account_h), self.document.total_quantity
        )

        total_top = account_top + account_h
        total_h = 36
        p.drawRect(QRectF(left, total_top, right - left, total_h))
        self._text(p, QRectF(left, total_top, body_left_w, total_h), "合       計", 10, False, family="굴림")
        self._draw_digits(p, QRectF(amount_left + 18, total_top + 1, right - amount_left - 18, total_h - 2), self.document.total_amount)

        lower_top = total_top + total_h
        lower_h = 180
        left_lower_w = 255
        p.drawRect(QRectF(left, lower_top, right - left, lower_h))
        self._line(p, left + left_lower_w, lower_top, left + left_lower_w, lower_top + lower_h)
        self._text(p, QRectF(left, lower_top, left_lower_w, 40), "決  係  代 理  課 長  所 長  協 助", 7, False, family="굴림")
        self._line(p, left, lower_top + 70, left + left_lower_w, lower_top + 70)
        self._text(p, QRectF(left, lower_top + 70, left_lower_w, 35), "決裁指示事項", 8, False, family="굴림")

        right_x = left + left_lower_w
        right_w = right - right_x
        self._text(p, QRectF(right_x, lower_top, right_w, 35), "上記 金額을 如히  請求함.", 8, False, family="굴림")
        date_text = f"{self.document.written_date.year}.   {self.document.written_date.month:02d} .   {self.document.written_date.day:02d} ."
        self._text(p, QRectF(right_x, lower_top + 32, right_w, 28), date_text, 8)
        self._text(p, QRectF(right_x, lower_top + 58, right_w, 30), f"請 求 人     {self.document.writer}  (印)", 8, False, family="굴림")
        self._line(p, right_x, lower_top + 92, right, lower_top + 92)
        self._text(p, QRectF(right_x, lower_top + 92, right_w, 28), "上記 金額을 如히  領收함.", 8, False, family="굴림")
        self._text(p, QRectF(right_x, lower_top + 135, right_w, 28), "請 求 人                    (印)", 8, False, family="굴림")

        footer_y = lower_top + lower_h + 4
        self._text(p, QRectF(left, footer_y, 140, 26), "양우F-152(O)", 9, True, Qt.AlignmentFlag.AlignLeft)
        self._text(p, QRectF(right - 180, footer_y, 180, 26), "양우건설주식회사", 9, True, Qt.AlignmentFlag.AlignRight)
