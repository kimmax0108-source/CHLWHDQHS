from __future__ import annotations

import re
import textwrap
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_CEILING, ROUND_DOWN, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Literal

INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]+')
INVALID_SHEET_CHARS = re.compile(r'[\\/*?:\[\]]+')
FAX_SPLIT_RE = re.compile(r"(?i)(?:\s*[.,;/·]?\s*)(?:F(?:AX)?|팩스)\s*[:：]?", re.UNICODE)
PAYMENT_SUFFIX_RE = re.compile(r"\s*\(\s*100\s*%\s*\)\s*(?:지급)?\s*$")
PHONE_PREFIX_RE = re.compile(r"(?i)^\s*(?:T(?:EL)?|전화)\s*[:：]?\s*")


ZERO = Decimal("0")


def parse_decimal(value: str | int | float | Decimal | None) -> Decimal:
    """쉼표/원화기호를 제거하되 소수 자릿수는 입력 그대로 보존한다."""

    if value is None or value == "":
        return ZERO
    if isinstance(value, bool):
        return Decimal(int(value))
    if isinstance(value, Decimal):
        return value
    text = str(value).strip().replace(",", "")
    text = text.replace("₩", "").replace("원", "").strip()
    text = re.sub(r"[^0-9eE+\-.]", "", text)
    if text in {"", "+", "-", ".", "+.", "-."}:
        return ZERO
    try:
        return Decimal(text)
    except InvalidOperation:
        return ZERO


def parse_money(value: str | int | float | Decimal | None) -> Decimal:
    return parse_decimal(value)


def decimal_text(value: str | int | float | Decimal | None, *, trim: bool = False) -> str:
    decimal_value = parse_decimal(value)
    text = format(decimal_value, "f")
    if trim and "." in text:
        text = text.rstrip("0").rstrip(".")
    if text in {"-0", "-0.0"}:
        return "0"
    return text


def format_money(value: str | int | float | Decimal | None) -> str:
    text = decimal_text(value)
    sign = ""
    if text.startswith("-"):
        sign, text = "-", text[1:]
    integer, dot, fraction = text.partition(".")
    integer = f"{int(integer or '0'):,}"
    return f"{sign}{integer}{dot}{fraction}" if dot else f"{sign}{integer}"


def compact_decimal(value: str | int | float | Decimal | None) -> Decimal:
    """값은 바꾸지 않고 계산 과정에서 생긴 불필요한 끝자리 0만 제거한다."""

    decimal_value = parse_decimal(value)
    return ZERO if decimal_value == 0 else decimal_value.normalize()


def apply_won_policy(value: str | int | float | Decimal | None, mode: str = "round") -> Decimal:
    """원 단위 처리 정책을 최종 합계에만 적용한다.

    품목 수량×단가 계산은 원본 정밀도를 그대로 유지하고, 공급가/부가세/합계처럼
    실제 결재 금액으로 쓰이는 최종값에만 반올림·올림·버림을 적용한다.
    """

    decimal_value = parse_decimal(value)
    normalized = str(mode or "round").strip().lower()
    if normalized in {"keep", "decimal", "소수점 유지"}:
        return compact_decimal(decimal_value)
    rounding = {
        "ceil": ROUND_CEILING,
        "올림": ROUND_CEILING,
        "floor": ROUND_DOWN,
        "버림": ROUND_DOWN,
        "round": ROUND_HALF_UP,
        "반올림": ROUND_HALF_UP,
    }.get(normalized, ROUND_HALF_UP)
    return decimal_value.quantize(Decimal("1"), rounding=rounding)


def round_won(value: float | Decimal) -> Decimal:
    # 하위 호환 이름. 기본 정책은 원 단위 반올림이다.
    return apply_won_policy(value, "round")

def safe_filename(text: str, fallback: str = "구매품의서") -> str:
    cleaned = INVALID_FILENAME_CHARS.sub("_", text).strip().strip(".")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned[:150] or fallback


def safe_sheet_label(text: str, fallback: str = "품목") -> str:
    cleaned = INVALID_SHEET_CHARS.sub("", text).strip()
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned[:20] or fallback


def build_sheet_name(prefix: str, label: str) -> str:
    safe = safe_sheet_label(label)
    available = 31 - len(prefix) - 2
    return f"{prefix}({safe[:available]})"


def build_classification(year: int, site_short: str, sequence: int) -> str:
    return f"자재 제{year} - {site_short.strip()} {sequence}호".strip()


def ensure_xlsx_suffix(path: str | Path) -> Path:
    p = Path(path)
    return p if p.suffix.lower() == ".xlsx" else p.with_suffix(".xlsx")


def wrap_body_text(text: str, width: int = 42, max_lines: int = 3) -> list[str]:
    normalized = str(text or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    if not normalized:
        return []
    wrapper = textwrap.TextWrapper(
        width=max(1, width),
        break_long_words=True,
        break_on_hyphens=False,
        replace_whitespace=False,
        drop_whitespace=True,
    )
    lines: list[str] = []
    for manual_line in normalized.split("\n"):
        manual_line = manual_line.strip()
        if manual_line:
            lines.extend(wrapper.wrap(manual_line) or [manual_line])
    if max_lines > 0 and len(lines) > max_lines:
        lines = lines[: max_lines - 1] + [" ".join(lines[max_lines - 1 :])]
    return lines


def normalize_quote_payment(value: str | None) -> str:
    """견적대비표에는 간결한 결제조건만 남긴다."""

    text = re.sub(r"\s+", " ", str(value or "")).strip()
    if not text:
        return "기성결제 현금"
    text = PAYMENT_SUFFIX_RE.sub("", text).strip()
    if text.endswith(" 지급"):
        text = text[:-3].strip()
    if "기성결제" in text and "현금" in text:
        return "기성결제 현금"
    return text


def purchase_payment_text(value: str | None) -> str:
    """구매품의서용 문구를 중복 없이 표준화한다."""

    base = normalize_quote_payment(value)
    if base == "기성결제 현금":
        return "기성결제 현금 (100%) 지급"
    return base


def extract_phone_only(value: str | None) -> str:
    """연락처 문자열에서 팩스 부분을 제거하고 전화번호만 반환한다."""

    text = str(value or "").strip()
    if not text:
        return ""
    phone = FAX_SPLIT_RE.split(text, maxsplit=1)[0]
    phone = PHONE_PREFIX_RE.sub("", phone)
    return phone.rstrip(" .,/;·").strip()


@dataclass(slots=True)
class Vendor:
    name: str = ""
    phone: str = ""
    manager: str = ""
    payment: str = "기성결제 현금"
    delivery_place: str = ""
    delivery_date: str = ""
    submitted: bool = True
    include_in_average: bool = True

    def normalize(self) -> None:
        self.payment = normalize_quote_payment(self.payment)

    @classmethod
    def from_dict(cls, value: dict[str, Any]) -> Vendor:
        vendor = cls(
            name=str(value.get("name", "")),
            phone=str(value.get("phone", "")),
            manager=str(value.get("manager", "")),
            payment=str(value.get("payment", "기성결제 현금")),
            delivery_place=str(value.get("delivery_place", "")),
            delivery_date=str(value.get("delivery_date", "")),
            submitted=bool(value.get("submitted", True)),
            include_in_average=bool(value.get("include_in_average", True)),
        )
        vendor.normalize()
        return vendor


@dataclass(slots=True)
class QuoteItem:
    # 기존 프로젝트의 위치 인수 호환을 위해 기존 필드 순서를 유지한다.
    name: str = ""
    spec: str = ""
    unit: str = "EA"
    quantity: Decimal = Decimal("1")
    unit_prices: list[Decimal] = field(default_factory=list)
    note: str = ""
    group_title: str = ""
    group_sequence: str = ""

    def normalize_prices(self, vendor_count: int) -> None:
        self.unit_prices = [parse_money(value) for value in self.unit_prices[:vendor_count]]
        self.unit_prices.extend([0] * (vendor_count - len(self.unit_prices)))

    def amount_for(self, vendor_index: int) -> Decimal:
        if vendor_index < 0 or vendor_index >= len(self.unit_prices):
            return ZERO
        return compact_decimal(self.quantity * self.unit_prices[vendor_index])

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["quantity"] = decimal_text(self.quantity)
        payload["unit_prices"] = [decimal_text(value) for value in self.unit_prices]
        return payload

    @classmethod
    def from_dict(cls, value: dict[str, Any]) -> QuoteItem:
        quantity_number = parse_decimal(value.get("quantity", 1))
        return cls(
            name=str(value.get("name", "")),
            spec=str(value.get("spec", "")),
            unit=str(value.get("unit", "")),
            quantity=quantity_number,
            unit_prices=[parse_money(v) for v in value.get("unit_prices", [])],
            note=str(value.get("note", "")),
            group_title=str(value.get("group_title", value.get("representative_name", ""))),
            group_sequence=str(value.get("group_sequence", "")),
        )


@dataclass(slots=True)
class StatementItem:
    group_title: str = ""
    number: str = ""
    name: str = ""
    spec: str = ""
    unit: str = "EA"
    quantity: Decimal = Decimal("1")
    unit_price: Decimal = ZERO
    amount: Decimal = ZERO
    note: str = ""

    def recalculate(self) -> None:
        self.unit_price = parse_money(self.unit_price)
        self.amount = compact_decimal(self.quantity * self.unit_price)

    @classmethod
    def from_dict(cls, value: dict[str, Any]) -> StatementItem:
        quantity = parse_decimal(value.get("quantity", 0))
        unit_price = parse_money(value.get("unit_price", 0))
        amount = parse_money(value.get("amount", 0))
        if not amount and unit_price:
            amount = compact_decimal(quantity * unit_price)
        return cls(
            group_title=str(value.get("group_title", "")),
            number=str(value.get("number", "")),
            name=str(value.get("name", "")),
            spec=str(value.get("spec", "")),
            unit=str(value.get("unit", "")),
            quantity=quantity,
            unit_price=unit_price,
            amount=amount,
            note=str(value.get("note", "")),
        )


QuoteRowKind = Literal["group", "detail"]


@dataclass(slots=True)
class ProjectData:
    # 공통/견적대비표
    site_name: str = ""
    site_short: str = ""
    item_label: str = "품목"
    quote_title: str = ""
    author: str = ""
    quote_date: date | str = ""
    common_delivery_place: str = ""
    vendors: list[Vendor] = field(default_factory=list)
    items: list[QuoteItem] = field(default_factory=list)
    selected_vendor_index: int | None = None
    budget_mode: str = "average"  # average | manual
    manual_budget_supply: Decimal = ZERO
    won_rounding: str = "round"  # round | ceil | floor | keep

    # 구매물품내역서 — 견적대비표에서 최초 생성 후 독립 편집
    statement_items: list[StatementItem] = field(default_factory=list)

    # 구매품의서
    classification: str = ""
    department: str = "자 재 부"
    draft_date: date | str = ""
    effective_date: str = "결재후 즉시"
    drafter: str = ""
    approval_note: str = "전결규정  제   조   항에 의한 전결사항임."
    purchase_title: str = ""
    purchase_item_name: str = ""
    period_kind: str = "임차기간"
    period: str = ""
    attachment: str = "구매물품내역서"
    payment: str = "기성결제 현금 (100%) 지급"
    body_text: str = ""
    note: str = ""
    statement_title: str = ""

    # 구매품의서 자동 연동값은 필요할 때 문서별로 덮어쓸 수 있다.
    purchase_site_override: str = ""
    purchase_vendor_override: str = ""
    purchase_phone_override: str = ""
    purchase_budget_override: Decimal = ZERO
    purchase_contract_override: Decimal = ZERO
    purchase_ratio_override: str = ""
    purchase_override_fields: list[str] = field(default_factory=list)

    def normalize(self) -> None:
        for vendor in self.vendors:
            vendor.normalize()
        for item in self.items:
            item.quantity = parse_decimal(item.quantity)
            item.normalize_prices(len(self.vendors))
        for item in self.statement_items:
            item.quantity = parse_decimal(item.quantity)
            item.unit_price = parse_money(item.unit_price)
            item.amount = parse_money(item.amount)
        if self.items and not any(item.group_title.strip() for item in self.items):
            self.items[0].group_title = (
                self.statement_title or self.quote_title or self.purchase_title or self.items[0].name
            )
        sequence = 0
        for item in self.items:
            if item.group_title.strip():
                sequence += 1
                if not item.group_sequence.strip():
                    item.group_sequence = str(sequence)
        if self.selected_vendor_index is not None and not (
            0 <= self.selected_vendor_index < len(self.vendors)
        ):
            self.selected_vendor_index = None
        if not self.common_delivery_place and self.vendors:
            self.common_delivery_place = self.vendors[0].delivery_place
        self.budget_mode = "manual" if str(self.budget_mode).lower() == "manual" else "average"
        self.manual_budget_supply = parse_money(self.manual_budget_supply)
        self.won_rounding = str(self.won_rounding or "round").lower()
        if self.won_rounding not in {"round", "ceil", "floor", "keep"}:
            self.won_rounding = "round"
        self.purchase_budget_override = parse_money(self.purchase_budget_override)
        self.purchase_contract_override = parse_money(self.purchase_contract_override)
        self.purchase_override_fields = list(dict.fromkeys(str(v) for v in self.purchase_override_fields))
        self.payment = purchase_payment_text(self.payment)

    @property
    def quote_date_text(self) -> str:
        return self._date_text(self.quote_date)

    @property
    def draft_date_text(self) -> str:
        return self._date_text(self.draft_date)

    @staticmethod
    def _date_text(value: date | str) -> str:
        if isinstance(value, datetime):
            return value.strftime("%Y.%m.%d")
        if isinstance(value, date):
            return value.strftime("%Y.%m.%d")
        return str(value or "").strip()

    def vendor_supply_raw(self, vendor_index: int) -> Decimal:
        return compact_decimal(sum((item.amount_for(vendor_index) for item in self.items if item.name.strip()), ZERO))

    def vendor_supply_total(self, vendor_index: int) -> Decimal:
        return apply_won_policy(self.vendor_supply_raw(vendor_index), self.won_rounding)

    def vendor_vat(self, vendor_index: int) -> Decimal:
        return apply_won_policy(self.vendor_supply_total(vendor_index) * Decimal("0.1"), self.won_rounding)

    def vendor_total(self, vendor_index: int) -> Decimal:
        return compact_decimal(self.vendor_supply_total(vendor_index) + self.vendor_vat(vendor_index))

    def eligible_vendor_indices(self) -> list[int]:
        result: list[int] = []
        for index, vendor in enumerate(self.vendors):
            if (
                vendor.name.strip()
                and vendor.submitted
                and vendor.include_in_average
                and self.vendor_total(index) > 0
            ):
                result.append(index)
        return result

    @property
    def auto_winner_index(self) -> int | None:
        eligible = self.eligible_vendor_indices()
        if not eligible:
            return None
        return min(eligible, key=lambda index: (self.vendor_total(index), index))

    @property
    def winner_index(self) -> int | None:
        if self.selected_vendor_index is not None:
            index = self.selected_vendor_index
            if 0 <= index < len(self.vendors) and self.vendor_total(index) > 0:
                return index
        return self.auto_winner_index

    @property
    def selected_vendor(self) -> Vendor | None:
        index = self.winner_index
        return None if index is None else self.vendors[index]

    @property
    def selected_vendor_phone(self) -> str:
        vendor = self.selected_vendor
        return extract_phone_only(vendor.phone if vendor else "")

    @property
    def purchase_payment(self) -> str:
        vendor = self.selected_vendor
        source = vendor.payment if vendor and vendor.payment else self.payment
        return purchase_payment_text(source)

    @property
    def contract_amount(self) -> Decimal:
        index = self.winner_index
        return ZERO if index is None else self.vendor_total(index)

    @property
    def budget_supply(self) -> Decimal:
        if self.budget_mode == "manual":
            return apply_won_policy(self.manual_budget_supply, self.won_rounding)
        eligible = self.eligible_vendor_indices()
        if not eligible:
            return ZERO
        raw_average = sum((self.vendor_supply_raw(i) for i in eligible), ZERO) / Decimal(len(eligible))
        return apply_won_policy(raw_average, self.won_rounding)

    @property
    def budget_vat(self) -> Decimal:
        return apply_won_policy(self.budget_supply * Decimal("0.1"), self.won_rounding)

    @property
    def budget_amount(self) -> Decimal:
        return compact_decimal(self.budget_supply + self.budget_vat)

    @property
    def ratio(self) -> Decimal | None:
        if not self.budget_amount:
            return None
        return self.contract_amount / self.budget_amount

    @property
    def ratio_text(self) -> str:
        if self.ratio is None:
            return "-"
        return f"{self.ratio * Decimal(100):.1f}%"

    def item_average_amount(self, item_index: int) -> Decimal:
        eligible = self.eligible_vendor_indices()
        if not eligible or not (0 <= item_index < len(self.items)):
            return ZERO
        item = self.items[item_index]
        return compact_decimal(sum((item.amount_for(i) for i in eligible), ZERO) / Decimal(len(eligible)))

    def rank_for(self, vendor_index: int) -> int | None:
        eligible = self.eligible_vendor_indices()
        if vendor_index not in eligible:
            return None
        ordered = sorted(eligible, key=lambda index: (self.vendor_total(index), index))
        return ordered.index(vendor_index) + 1

    def quote_output_rows(self) -> list[tuple[QuoteRowKind, str, QuoteItem | None]]:
        rows: list[tuple[QuoteRowKind, str, QuoteItem | None]] = []
        auto_sequence = 0
        for item in (item for item in self.items if item.name.strip()):
            if item.group_title.strip():
                auto_sequence += 1
                sequence = item.group_sequence.strip() or str(auto_sequence)
                rows.append(("group", sequence, item))
            rows.append(("detail", "", item))
        return rows

    @property
    def quote_output_row_count(self) -> int:
        return len(self.quote_output_rows())

    def build_statement_items_from_quote(self) -> list[StatementItem]:
        winner = self.winner_index
        result: list[StatementItem] = []
        detail_number = 0
        for item in (item for item in self.items if item.name.strip()):
            detail_number += 1
            unit_price = item.unit_prices[winner] if winner is not None else 0
            result.append(
                StatementItem(
                    group_title=item.group_title,
                    number=str(detail_number),
                    name=item.name,
                    spec=item.spec,
                    unit=item.unit,
                    quantity=item.quantity,
                    unit_price=unit_price,
                    amount=item.amount_for(winner) if winner is not None else 0,
                    note=item.note,
                )
            )
        return result

    def sync_statement_from_quote(self) -> None:
        self.statement_items = self.build_statement_items_from_quote()

    @property
    def statement_supply_raw(self) -> Decimal:
        return compact_decimal(sum((parse_money(item.amount) for item in self.statement_items if item.name.strip()), ZERO))

    @property
    def statement_supply_total(self) -> Decimal:
        return apply_won_policy(self.statement_supply_raw, self.won_rounding)

    @property
    def statement_vat(self) -> Decimal:
        return apply_won_policy(self.statement_supply_total * Decimal("0.1"), self.won_rounding)

    @property
    def statement_total(self) -> Decimal:
        return compact_decimal(self.statement_supply_total + self.statement_vat)

    @property
    def purchase_site_effective(self) -> str:
        return self.purchase_site_override if "site" in self.purchase_override_fields else self.site_name

    @property
    def purchase_vendor_effective(self) -> str:
        vendor = self.selected_vendor
        automatic = vendor.name if vendor else ""
        return self.purchase_vendor_override if "vendor" in self.purchase_override_fields else automatic

    @property
    def purchase_phone_effective(self) -> str:
        automatic = self.selected_vendor_phone
        return self.purchase_phone_override if "phone" in self.purchase_override_fields else automatic

    @property
    def purchase_budget_effective(self) -> Decimal:
        return self.purchase_budget_override if "budget" in self.purchase_override_fields else self.budget_amount

    @property
    def purchase_contract_effective(self) -> Decimal:
        return self.purchase_contract_override if "contract" in self.purchase_override_fields else self.contract_amount

    @property
    def purchase_ratio_effective(self) -> Decimal | None:
        if "ratio" in self.purchase_override_fields:
            text = str(self.purchase_ratio_override or "").strip().replace("%", "")
            return parse_decimal(text) / Decimal(100) if text else None
        budget = self.purchase_budget_effective
        return self.purchase_contract_effective / budget if budget else None

    @property
    def purchase_ratio_text(self) -> str:
        ratio = self.purchase_ratio_effective
        return "-" if ratio is None else f"{ratio * Decimal(100):.1f}%"

    @property
    def statement_output_row_count(self) -> int:
        return sum(1 + int(bool(item.group_title.strip())) for item in self.statement_items if item.name.strip())

    @property
    def statement_sheet_name(self) -> str:
        return build_sheet_name("내역서", self.item_label)

    @property
    def quote_sheet_name(self) -> str:
        return build_sheet_name("견적대비표", self.item_label)

    def suggested_filename(self) -> str:
        classification = self.classification.replace("자재 제", "").strip()
        title = self.purchase_title or self.quote_title or self.item_label
        base = "-".join(part for part in [classification, title] if part)
        return safe_filename(base or "자재구매품의서") + ".xlsx"

    def validate(self) -> list[str]:
        self.normalize()
        if not self.statement_items and self.items:
            self.sync_statement_from_quote()
        errors: list[str] = []
        if not self.site_name.strip():
            errors.append("현장명을 입력해 주세요.")
        if not self.item_label.strip():
            errors.append("시트용 품목명을 입력해 주세요.")
        if not self.classification.strip():
            errors.append("분류번호를 입력해 주세요.")
        valid_items = [item for item in self.items if item.name.strip()]
        if not valid_items:
            errors.append("견적 품목을 한 개 이상 입력해 주세요.")
        if self.quote_output_row_count > 17:
            errors.append("견적대비표 양식의 제목행·상세행 합계는 최대 17행입니다.")
        if self.statement_output_row_count > 51:
            errors.append("구매물품내역서의 제목행·상세행 합계는 최대 51행입니다.")
        if not self.vendors:
            errors.append("견적업체를 한 곳 이상 추가해 주세요.")
        if self.auto_winner_index is None:
            errors.append("업체명과 견적단가가 입력된 정상 견적업체가 없습니다.")
        if not self.purchase_title.strip():
            errors.append("구매품의서 제목을 입력해 주세요.")
        if not self.purchase_item_name.strip():
            errors.append("구매품의서 품명을 입력해 주세요.")
        return errors

    def to_dict(self) -> dict[str, Any]:
        def convert(value: Any) -> Any:
            if isinstance(value, Decimal):
                return decimal_text(value)
            if isinstance(value, list):
                return [convert(item) for item in value]
            if isinstance(value, dict):
                return {key: convert(item) for key, item in value.items()}
            return value

        payload = convert(asdict(self))
        payload["quote_date"] = self.quote_date_text
        payload["draft_date"] = self.draft_date_text
        return payload

    @classmethod
    def from_dict(cls, value: dict[str, Any]) -> ProjectData:
        data = cls(
            site_name=str(value.get("site_name", "")),
            site_short=str(value.get("site_short", "")),
            item_label=str(value.get("item_label", "품목")),
            quote_title=str(value.get("quote_title", "")),
            author=str(value.get("author", "")),
            quote_date=str(value.get("quote_date", "")),
            common_delivery_place=str(value.get("common_delivery_place", "")),
            vendors=[Vendor.from_dict(v) for v in value.get("vendors", [])],
            items=[QuoteItem.from_dict(v) for v in value.get("items", [])],
            selected_vendor_index=value.get("selected_vendor_index"),
            budget_mode=str(value.get("budget_mode", "average")),
            manual_budget_supply=parse_money(value.get("manual_budget_supply", 0)),
            won_rounding=str(value.get("won_rounding", "round")),
            statement_items=[StatementItem.from_dict(v) for v in value.get("statement_items", [])],
            classification=str(value.get("classification", "")),
            department=str(value.get("department", "자 재 부")),
            draft_date=str(value.get("draft_date", "")),
            effective_date=str(value.get("effective_date", "결재후 즉시")),
            drafter=str(value.get("drafter", "")),
            approval_note=str(value.get("approval_note", "전결규정  제   조   항에 의한 전결사항임.")),
            purchase_title=str(value.get("purchase_title", "")),
            purchase_item_name=str(value.get("purchase_item_name", "")),
            period_kind=str(value.get("period_kind", "임차기간")),
            period=str(value.get("period", "")),
            attachment=str(value.get("attachment", "구매물품내역서")),
            payment=str(value.get("payment", "기성결제 현금 (100%) 지급")),
            body_text=str(value.get("body_text", "")),
            note=str(value.get("note", "")),
            statement_title=str(value.get("statement_title", "")),
            purchase_site_override=str(value.get("purchase_site_override", "")),
            purchase_vendor_override=str(value.get("purchase_vendor_override", "")),
            purchase_phone_override=str(value.get("purchase_phone_override", "")),
            purchase_budget_override=parse_money(value.get("purchase_budget_override", 0)),
            purchase_contract_override=parse_money(value.get("purchase_contract_override", 0)),
            purchase_ratio_override=str(value.get("purchase_ratio_override", "")),
            purchase_override_fields=[str(v) for v in value.get("purchase_override_fields", [])],
        )
        if data.selected_vendor_index is not None:
            try:
                data.selected_vendor_index = int(data.selected_vendor_index)
            except (TypeError, ValueError):
                data.selected_vendor_index = None
        data.normalize()
        if not data.statement_items and data.items:
            data.sync_statement_from_quote()
        return data
