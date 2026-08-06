from __future__ import annotations

from collections import defaultdict
from typing import Iterable, Optional

from .models import LedgerRow, MoneySummary, calc_vat
from .storage import ClaimOverrideStore


def month_key(year: int, month: int) -> str:
    return f"{year:04d}-{month:02d}"


def shift_month(value: str, offset: int) -> str:
    year, month = map(int, value.split("-"))
    serial = year * 12 + (month - 1) + offset
    return f"{serial // 12:04d}-{serial % 12 + 1:02d}"


def effective_claim_month(row: LedgerRow, store: ClaimOverrideStore) -> str:
    return store.get(row.fingerprint) or row.intake_month


def effective_trade(row: LedgerRow, store: ClaimOverrideStore) -> str:
    return store.classification(row.fingerprint) or row.trade


def is_excluded(row: LedgerRow, store: ClaimOverrideStore) -> bool:
    return store.is_excluded(row.fingerprint)


def processing_status(row: LedgerRow, store: ClaimOverrideStore) -> str:
    states: list[str] = []
    if effective_trade(row, store) != row.trade:
        states.append("분류변경")
    if is_excluded(row, store):
        states.append("제외")
    return ", ".join(states) if states else "청구대상"


def rollover_status(row: LedgerRow, store: ClaimOverrideStore) -> str:
    claim_month = effective_claim_month(row, store)
    if claim_month == row.intake_month:
        return "당월 입고"
    if claim_month > row.intake_month:
        return "이월 청구"
    return "선청구"


def money_summary(rows: Iterable[LedgerRow]) -> MoneySummary:
    values = list(rows)
    supply = sum(row.amount for row in values)
    vat = calc_vat(supply)
    return MoneySummary(
        supply=supply,
        vat=vat,
        total=supply + vat,
        count=len(values),
        quantity=sum(row.quantity for row in values),
    )


def period_matches_month(month_value: str, year: Optional[int], month: Optional[int]) -> bool:
    row_year, row_month = map(int, month_value.split("-"))
    if year is not None and row_year != year:
        return False
    if month is not None and row_month != month:
        return False
    return True


def period_matches_row(row: LedgerRow, year: Optional[int], month: Optional[int]) -> bool:
    if year is not None and row.intake_date.year != year:
        return False
    if month is not None and row.intake_date.month != month:
        return False
    return True


def filter_base(
    rows: Iterable[LedgerRow],
    store: ClaimOverrideStore,
    *,
    year: Optional[int] = None,
    month: Optional[int] = None,
    trade: str = "전체",
    vendors: Optional[set[str]] = None,
    items: Optional[set[str]] = None,
    source_sheets: Optional[set[str]] = None,
    keyword: str = "",
    include_excluded: bool = True,
) -> list[LedgerRow]:
    needle = keyword.strip().lower()
    output: list[LedgerRow] = []
    for row in rows:
        if source_sheets and row.source_sheet not in source_sheets:
            continue
        if not period_matches_row(row, year, month):
            continue
        if trade not in ("", "전체") and effective_trade(row, store) != trade:
            continue
        if vendors and row.vendor not in vendors:
            continue
        if items and row.item not in items:
            continue
        if not include_excluded and is_excluded(row, store):
            continue
        if needle:
            haystack = " ".join(
                [
                    row.source_sheet,
                    row.trade,
                    effective_trade(row, store),
                    row.item,
                    row.spec,
                    row.vendor,
                    row.usage,
                    row.note,
                    store.management_note(row.fingerprint),
                ]
            ).lower()
            if needle not in haystack:
                continue
        output.append(row)
    return output


def claim_breakdown_period(
    rows: Iterable[LedgerRow],
    store: ClaimOverrideStore,
    *,
    year: Optional[int],
    month: Optional[int],
    trade: str = "전체",
    vendors: Optional[set[str]] = None,
    items: Optional[set[str]] = None,
    source_sheets: Optional[set[str]] = None,
    keyword: str = "",
) -> dict[str, list[LedgerRow]]:
    # 분류/업체/품명/원본시트 검색은 적용하되 기간 조건은 입고월과 청구월에 따로 적용한다.
    base = filter_base(
        rows,
        store,
        year=None,
        month=None,
        trade=trade,
        vendors=vendors,
        items=items,
        source_sheets=source_sheets,
        keyword=keyword,
        include_excluded=True,
    )
    intake_all = [row for row in base if period_matches_row(row, year, month)]
    claim_all = [
        row
        for row in base
        if period_matches_month(effective_claim_month(row, store), year, month)
    ]
    brought_all = [row for row in claim_all if not period_matches_row(row, year, month)]
    moved_all = [
        row
        for row in intake_all
        if not period_matches_month(effective_claim_month(row, store), year, month)
    ]

    active = lambda values: [row for row in values if not is_excluded(row, store)]
    return {
        "current_intake": active(intake_all),
        "brought_in": active(brought_all),
        "moved_out": active(moved_all),
        "claim_target": active(claim_all),
        "display_current_intake": intake_all,
        "display_brought_in": brought_all,
        "display_moved_out": moved_all,
        "display_claim_rows": claim_all,
    }


def grouped_summary(rows: Iterable[LedgerRow], key) -> list[tuple[str, MoneySummary]]:
    groups: dict[str, list[LedgerRow]] = defaultdict(list)
    for row in rows:
        groups[key(row) or "(미입력)"].append(row)
    return [(name, money_summary(values)) for name, values in groups.items()]


def row_info(row: LedgerRow) -> dict[str, object]:
    return {
        "source_sheet": row.source_sheet,
        "source_row": row.source_row,
        "date": row.intake_date.isoformat(),
        "trade": row.trade,
        "item": row.item,
        "vendor": row.vendor,
        "amount": row.amount,
    }


def set_claim_month(
    rows: Iterable[LedgerRow],
    store: ClaimOverrideStore,
    target_month: str,
    *,
    reason: str = "",
) -> None:
    for row in rows:
        store.set(
            row.fingerprint,
            target_month,
            row.intake_month,
            reason=reason,
            row_info=row_info(row),
            autosave=False,
        )
    store.save()


def cancel_rollover(
    rows: Iterable[LedgerRow], store: ClaimOverrideStore, *, reason: str = ""
) -> None:
    for row in rows:
        store.remove(
            row.fingerprint,
            intake_month=row.intake_month,
            reason=reason,
            row_info=row_info(row),
            autosave=False,
        )
    store.save()


def set_classification(
    rows: Iterable[LedgerRow],
    store: ClaimOverrideStore,
    classification: str,
    *,
    reason: str = "",
) -> None:
    for row in rows:
        store.set_classification(
            row.fingerprint,
            classification,
            row.trade,
            reason=reason,
            row_info=row_info(row),
            autosave=False,
        )
    store.save()


def reset_classification(
    rows: Iterable[LedgerRow], store: ClaimOverrideStore, *, reason: str = ""
) -> None:
    for row in rows:
        store.reset_classification(
            row.fingerprint,
            row.trade,
            reason=reason,
            row_info=row_info(row),
            autosave=False,
        )
    store.save()


def exclude_rows(
    rows: Iterable[LedgerRow], store: ClaimOverrideStore, *, reason: str = ""
) -> None:
    for row in rows:
        store.exclude(
            row.fingerprint,
            reason=reason,
            row_info=row_info(row),
            autosave=False,
        )
    store.save()


def restore_excluded_rows(
    rows: Iterable[LedgerRow], store: ClaimOverrideStore, *, reason: str = ""
) -> None:
    for row in rows:
        store.restore_excluded(
            row.fingerprint,
            reason=reason,
            row_info=row_info(row),
            autosave=False,
        )
    store.save()
