from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any, Literal

QTY_STEP = Decimal("0.001")


def parse_decimal(value: Any) -> Decimal:
    if isinstance(value, Decimal):
        return value
    text = str(value or "").strip().replace(",", "")
    if not text:
        return Decimal("0")
    try:
        return Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"숫자로 해석할 수 없습니다: {value}") from exc


def parse_money(value: Any) -> int:
    return int(parse_decimal(value).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def format_money(value: int | Decimal) -> str:
    return f"{int(value):,}"


def normalize_quantity(value: Any) -> Decimal:
    return parse_decimal(value).quantize(QTY_STEP, rounding=ROUND_HALF_UP)


def format_quantity(value: Any) -> str:
    return f"{normalize_quantity(value):.3f}"


def quantity_parts(value: Any) -> tuple[str, str]:
    text = format_quantity(value)
    integer, fraction = text.split(".", 1)
    return f"{integer}.", fraction


@dataclass(slots=True)
class AllocationRow:
    month: int = 4
    site_name: str = ""
    vendor_name: str = ""
    quantity: Decimal = Decimal("0.000")
    amount: int = 0

    @property
    def is_active(self) -> bool:
        return bool(self.site_name.strip() or self.vendor_name.strip() or self.amount)

    @property
    def display_site(self) -> str:
        site = self.site_name.strip()
        # Continuation vendor rows intentionally leave the month/site cell blank.
        return f"{self.month:02d}월 {site}" if site else ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "month": self.month,
            "site_name": self.site_name,
            "vendor_name": self.vendor_name,
            "quantity": format_quantity(self.quantity),
            "amount": self.amount,
        }

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "AllocationRow":
        return cls(
            month=int(raw.get("month", 4)),
            site_name=str(raw.get("site_name", "")),
            vendor_name=str(raw.get("vendor_name", "")),
            quantity=normalize_quantity(raw.get("quantity", 0)),
            amount=parse_money(raw.get("amount", 0)),
        )


@dataclass(slots=True)
class VendorAccount:
    vendor_name: str = ""
    bank_name: str = ""
    account_number: str = ""
    account_holder: str = ""

    def display_text(self) -> str:
        vendor = self.vendor_name.strip()
        bank = self.bank_name.strip()
        account = self.account_number.strip()
        if not vendor:
            return ""
        if bank and account:
            return f"{vendor} : ({bank}){account}"
        if account:
            return f"{vendor} : {account}"
        return vendor


@dataclass(frozen=True, slots=True)
class BodyLine:
    kind: Literal["allocation", "description", "account"]
    allocation: AllocationRow | None = None
    text: str = ""


@dataclass(slots=True)
class ExpenseDocument:
    payment_month: int = 4
    written_date: date = field(default_factory=date.today)
    writer: str = ""
    payment_title: str = "정기결제"
    vat_included: bool = True
    allocations: list[AllocationRow] = field(default_factory=list)
    descriptions: list[str] = field(default_factory=list)
    vendor_accounts: list[VendorAccount] = field(default_factory=list)

    @property
    def active_allocations(self) -> list[AllocationRow]:
        return [row for row in self.allocations if row.is_active]

    @property
    def total_quantity(self) -> Decimal:
        total = sum((row.quantity for row in self.active_allocations), Decimal("0"))
        return total.quantize(QTY_STEP, rounding=ROUND_HALF_UP)

    @property
    def total_amount(self) -> int:
        return sum(row.amount for row in self.active_allocations)

    @property
    def used_line_count(self) -> int:
        return len(self.export_body_lines())

    def unique_accounts(self) -> list[VendorAccount]:
        order: list[str] = []
        by_name: dict[str, VendorAccount] = {}
        for account in self.vendor_accounts:
            key = account.vendor_name.strip()
            if key and key not in by_name:
                order.append(key)
                by_name[key] = account
        return [by_name[key] for key in order]

    def export_body_lines(self) -> list[BodyLine]:
        """Return every printable body row in display order.

        The Excel writer grows the form vertically when additional vendor-account
        rows are required, so account information never has to share the quantity
        summary row or use a reduced font.
        """
        lines: list[BodyLine] = [
            BodyLine("allocation", allocation=row) for row in self.active_allocations
        ]
        lines.extend(
            BodyLine("description", text=text.strip())
            for text in self.descriptions
            if text.strip()
        )
        lines.extend(
            BodyLine("account", text=account.display_text())
            for account in self.unique_accounts()
            if account.display_text()
        )
        return lines

    def validate(self) -> None:
        if not self.active_allocations:
            raise ValueError("원가 안분 내역을 한 행 이상 입력해 주세요.")
        if self.total_amount <= 0:
            raise ValueError("금액 합계가 0원입니다.")
        if self.total_amount > 9_999_999_999:
            raise ValueError("양식의 금액 칸은 99억 9,999만 9,999원까지 지원합니다.")
        if len(self.export_body_lines()) > 20:
            raise ValueError(
                "안분·적요·계좌 행이 20줄을 초과했습니다. "
                "한 페이지 가독성을 위해 지출결의서를 나눠 작성해 주세요."
            )

    def to_dict(self) -> dict[str, Any]:
        return {
            "payment_month": self.payment_month,
            "written_date": self.written_date.isoformat(),
            "writer": self.writer,
            "payment_title": self.payment_title,
            "vat_included": self.vat_included,
            "allocations": [row.to_dict() for row in self.allocations],
            "descriptions": list(self.descriptions),
            "vendor_accounts": [asdict(account) for account in self.vendor_accounts],
        }

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "ExpenseDocument":
        written = raw.get("written_date") or date.today().isoformat()
        return cls(
            payment_month=int(raw.get("payment_month", 4)),
            written_date=date.fromisoformat(str(written)),
            writer=str(raw.get("writer", "")),
            payment_title=str(raw.get("payment_title", "정기결제")),
            vat_included=bool(raw.get("vat_included", True)),
            allocations=[AllocationRow.from_dict(v) for v in raw.get("allocations", [])],
            descriptions=[str(v) for v in raw.get("descriptions", [])],
            vendor_accounts=[VendorAccount(**v) for v in raw.get("vendor_accounts", [])],
        )
