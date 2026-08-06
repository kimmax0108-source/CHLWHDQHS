from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import hashlib


def round_won(value: float) -> int:
    """숫자를 원 단위로 사사오입한다."""
    try:
        return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    except (InvalidOperation, ValueError, TypeError):
        return 0


def calc_vat(amount: float) -> int:
    """공급가액의 10%를 원 단위 반올림한다."""
    try:
        return int(
            (Decimal(str(amount)) * Decimal("0.1")).quantize(
                Decimal("1"), rounding=ROUND_HALF_UP
            )
        )
    except (InvalidOperation, ValueError, TypeError):
        return 0


@dataclass(frozen=True)
class LedgerRow:
    source_row: int
    intake_date: date
    trade: str
    item: str
    spec: str
    unit: str
    length: str
    quantity: float
    unit_price: float
    amount: float
    vendor: str
    usage: str
    note: str
    source_sheet: str = ""
    quantity_entered: bool = True
    unit_price_entered: bool = True
    amount_entered: bool = True

    @property
    def vat(self) -> int:
        return calc_vat(self.amount)

    @property
    def total(self) -> float:
        return self.amount + self.vat

    @property
    def intake_month(self) -> str:
        return self.intake_date.strftime("%Y-%m")

    @property
    def calculated_amount(self) -> int | None:
        if not self.quantity_entered or not self.unit_price_entered:
            return None
        return round_won(self.quantity * self.unit_price)

    @property
    def amount_difference(self) -> int | None:
        calculated = self.calculated_amount
        if calculated is None or not self.amount_entered:
            return None
        return round_won(self.amount) - calculated

    @property
    def amount_review_status(self) -> str:
        difference = self.amount_difference
        if difference is None:
            return "검토불가"
        return "정상" if difference == 0 else "불일치"

    @property
    def fingerprint(self) -> str:
        # v1.1 단일시트 대장과의 관리값 호환을 위해 source_sheet가 비어 있으면
        # 기존 fingerprint 규칙을 그대로 사용한다. 표준 다중시트 대장은 시트명을 포함해
        # 서로 다른 시트의 동일 행이 충돌하지 않게 한다.
        parts = [
            self.intake_date.isoformat(),
            self.trade,
            self.item,
            self.spec,
            self.unit,
            f"{self.quantity:.6f}",
            f"{self.unit_price:.4f}",
            f"{self.amount:.2f}",
            self.vendor,
            self.usage,
            str(self.source_row),
        ]
        normalized_sheet = "".join(self.source_sheet.split()).casefold()
        if self.source_sheet and normalized_sheet != "자재입출고대장".casefold():
            parts.insert(0, self.source_sheet)
        payload = "|".join(parts)
        return hashlib.sha1(payload.encode("utf-8")).hexdigest()[:24]


@dataclass
class LedgerData:
    path: str
    sheet_name: str
    site_name: str
    rows: list[LedgerRow]
    header_row: int
    data_start_row: int
    sheet_names: list[str]
    loaded_sheet_names: list[str] = field(default_factory=list)


@dataclass
class MoneySummary:
    supply: float = 0.0
    vat: int = 0
    total: float = 0.0
    count: int = 0
    quantity: float = 0.0
