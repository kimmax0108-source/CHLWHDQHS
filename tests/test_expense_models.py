from datetime import date
from decimal import Decimal

from expense_statement_app.models import AllocationRow, ExpenseDocument, VendorAccount


def sample_document() -> ExpenseDocument:
    return ExpenseDocument(
        payment_month=8,
        written_date=date(2025, 9, 30),
        writer="김태강",
        allocations=[
            AllocationRow(8, "연신내 주상복합", "리소스뱅크㈜", Decimal("187.114"), 149_547_354),
            AllocationRow(8, "", "동경강업㈜", Decimal("25.976"), 22_001_672),
        ],
        descriptions=["철근구입비-정기결제", "부가세 포함"],
        vendor_accounts=[
            VendorAccount("리소스뱅크㈜", "우리", "1005-204-378238", "리소스뱅크㈜"),
            VendorAccount("동경강업㈜", "기업", "378-023864-04-011", "동경강업㈜"),
        ],
    )


def test_totals_and_serialization() -> None:
    doc = sample_document()
    assert doc.total_quantity == Decimal("213.090")
    assert doc.total_amount == 171_549_026
    loaded = ExpenseDocument.from_dict(doc.to_dict())
    assert loaded.total_amount == doc.total_amount
    assert loaded.total_quantity == doc.total_quantity


def test_blank_site_does_not_print_month() -> None:
    doc = sample_document()
    assert doc.active_allocations[0].display_site == "08월 연신내 주상복합"
    assert doc.active_allocations[1].display_site == ""


def test_accounts_are_separate_body_lines() -> None:
    doc = sample_document()
    lines = doc.export_body_lines()
    assert [line.kind for line in lines] == [
        "allocation",
        "allocation",
        "description",
        "description",
        "account",
        "account",
    ]
    assert lines[-2].text.startswith("리소스뱅크㈜")
    assert lines[-1].text.startswith("동경강업㈜")


def test_three_vendors_fit_last_account_on_summary_row() -> None:
    doc = sample_document()
    doc.allocations.append(
        AllocationRow(8, "", "환영철강㈜", Decimal("42.536"), 30_000_000)
    )
    doc.vendor_accounts.append(
        VendorAccount("환영철강㈜", "하나", "176-910006-57004", "환영철강㈜")
    )
    lines = doc.export_body_lines()
    assert len(lines) == 8
    assert lines[-1].kind == "account"


def test_additional_accounts_are_allowed_with_dynamic_rows() -> None:
    doc = sample_document()
    for index in range(2):
        vendor = f"추가업체{index + 1}"
        doc.allocations.append(AllocationRow(8, "", vendor, Decimal("1.000"), 10_000))
        doc.vendor_accounts.append(VendorAccount(vendor, "은행", f"000-{index}", vendor))
    doc.validate()
    assert len(doc.export_body_lines()) == 10
