from datetime import date
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from material_claim_manager.models import LedgerRow
from material_claim_manager.services import (
    claim_breakdown_period,
    effective_trade,
    filter_base,
    shift_month,
)


class MemoryStore:
    def __init__(self):
        self.values = {}
        self.classes = {}
        self.excluded = set()
        self.notes = {}

    def get(self, key):
        return self.values.get(key)

    def classification(self, key):
        return self.classes.get(key)

    def is_excluded(self, key):
        return key in self.excluded

    def management_note(self, key):
        return self.notes.get(key, "")


def row(day, amount, *, trade="주자재", item="레미콘", vendor="성신레미컨", qty=1, price=None, sheet=""):
    unit_price = amount if price is None else price
    return LedgerRow(
        1,
        day,
        trade,
        item,
        "",
        "㎥",
        "",
        qty,
        unit_price,
        amount,
        vendor,
        "",
        "",
        source_sheet=sheet,
    )


def test_shift_month_crosses_year():
    assert shift_month("2026-12", 1) == "2027-01"
    assert shift_month("2026-01", -1) == "2025-12"


def test_claim_breakdown_handles_brought_moved_and_excluded():
    may = row(date(2026, 5, 29), 100)
    june = row(date(2026, 6, 2), 200)
    excluded = row(date(2026, 6, 3), 300, item="제외품")
    store = MemoryStore()
    store.values[may.fingerprint] = "2026-06"
    store.values[june.fingerprint] = "2026-07"
    store.excluded.add(excluded.fingerprint)

    result = claim_breakdown_period(
        [may, june, excluded], store, year=2026, month=6
    )
    assert result["current_intake"] == [june]
    assert result["brought_in"] == [may]
    assert result["moved_out"] == [june]
    assert result["claim_target"] == [may]
    assert excluded in result["display_claim_rows"]


def test_whole_year_period_is_supported():
    jan = row(date(2026, 1, 2), 100)
    dec = row(date(2026, 12, 2), 200, item="물차")
    next_year = row(date(2027, 1, 2), 300, item="안전모")
    store = MemoryStore()
    result = claim_breakdown_period(
        [jan, dec, next_year], store, year=2026, month=None
    )
    assert result["claim_target"] == [jan, dec]


def test_manual_classification_drives_trade_filter():
    tarp = row(date(2026, 6, 18), 350000, trade="잡자재", item="고급천막")
    store = MemoryStore()
    store.classes[tarp.fingerprint] = "가설"
    assert effective_trade(tarp, store) == "가설"
    assert filter_base([tarp], store, trade="잡자재") == []
    assert filter_base([tarp], store, trade="가설") == [tarp]


def test_vendor_multiselect_uses_or_condition():
    a = row(date(2026, 6, 1), 100, vendor="A상사")
    b = row(date(2026, 6, 2), 100, vendor="B상사", item="물차")
    c = row(date(2026, 6, 3), 100, vendor="C상사", item="안전모")
    store = MemoryStore()
    result = filter_base([a, b, c], store, vendors={"A상사", "C상사"})
    assert result == [a, c]


def test_amount_review_status_flags_difference_and_blank_inputs():
    mismatch = row(date(2026, 6, 1), 128000, qty=20, price=6410)
    assert mismatch.calculated_amount == 128200
    assert mismatch.amount_difference == -200
    assert mismatch.amount_review_status == "불일치"

    blank_price = LedgerRow(
        2,
        date(2026, 6, 2),
        "잡자재",
        "샘플",
        "",
        "개",
        "",
        3,
        0,
        3000,
        "A상사",
        "",
        "",
        quantity_entered=True,
        unit_price_entered=False,
    )
    assert blank_price.amount_review_status == "검토불가"


def test_source_sheet_filter_uses_selected_standard_sheets():
    misc = row(date(2026, 6, 1), 100, item="고급천막", sheet="잡자재")
    main = row(date(2026, 6, 2), 200, item="레미콘", sheet="주자재")
    safety = row(date(2026, 6, 3), 300, item="안전모", sheet="안전")
    store = MemoryStore()

    result = filter_base(
        [misc, main, safety], store, source_sheets={"잡자재", "안전"}
    )
    assert result == [misc, safety]


def test_fingerprint_is_unique_between_standard_sheets():
    misc = row(date(2026, 6, 1), 100, sheet="잡자재")
    main = row(date(2026, 6, 1), 100, sheet="주자재")
    assert misc.fingerprint != main.fingerprint
