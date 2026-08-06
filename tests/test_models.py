from __future__ import annotations

from decimal import Decimal

from purchase_request_app.models import (
    ProjectData,
    QuoteItem,
    Vendor,
    extract_phone_only,
    normalize_quote_payment,
    purchase_payment_text,
    wrap_body_text,
)


def test_payment_and_phone_normalization() -> None:
    assert normalize_quote_payment("기성결제 현금 (100%) 지급") == "기성결제 현금"
    assert purchase_payment_text("기성결제 현금") == "기성결제 현금 (100%) 지급"
    assert purchase_payment_text("기성결제 현금 (100%)") == "기성결제 현금 (100%) 지급"
    assert extract_phone_only("T 031)355-3000 F 031)366-4545") == "031)355-3000"
    assert extract_phone_only("T 031)355-3000, FAX: 031)366-4545") == "031)355-3000"


def test_quote_group_rows_and_statement_sync_are_editable() -> None:
    data = ProjectData(
        vendors=[Vendor("A", payment="기성결제 현금"), Vendor("B")],
        items=[
            QuoteItem(
                "포세식 화장실", "3 * 6", "EA", 2, [8_200_000, 9_700_000],
                group_title="포세식 화장실 임차(18개월)", group_sequence="1"
            ),
            QuoteItem("운반비", "25톤", "대", 1, [500_000, 600_000]),
        ],
    )
    data.normalize()
    rows = data.quote_output_rows()
    assert [(kind, seq) for kind, seq, _ in rows] == [("group", "1"), ("detail", ""), ("detail", "")]
    data.sync_statement_from_quote()
    assert [item.number for item in data.statement_items] == ["1", "2"]
    assert data.statement_items[0].group_title == "포세식 화장실 임차(18개월)"
    assert data.statement_items[0].unit_price == 8_200_000
    data.statement_items[0].name = "수정 가능한 품명"
    data.statement_items[0].number = "A-1"
    assert data.statement_items[0].name == "수정 가능한 품명"
    assert data.statement_items[0].number == "A-1"


def test_budget_contract_and_winner() -> None:
    data = ProjectData(
        vendors=[Vendor("A"), Vendor("B"), Vendor("C", submitted=False)],
        items=[QuoteItem("품목", "", "EA", 1, [100, 110, 90])],
    )
    assert data.contract_amount == 110
    assert data.budget_supply == 105
    assert data.budget_amount == 116
    data.selected_vendor_index = 1
    assert data.contract_amount == 121
    assert data.winner_index == 1


def test_body_auto_wrap_matches_template_width() -> None:
    text = (
        "과천주암 C-2BL 공동주택 건설공사 2공구 현장의 현장직원 및 감리단 임시 "
        "사무실 용도의 컨테이너를 임차하고자 업체를 선정하여 품의하오니 결재 하여 "
        "주시기 바랍니다."
    )
    assert wrap_body_text(text, width=42, max_lines=3) == [
        "과천주암 C-2BL 공동주택 건설공사 2공구 현장의 현장직원 및 감리단 임시",
        "사무실 용도의 컨테이너를 임차하고자 업체를 선정하여 품의하오니 결재 하여",
        "주시기 바랍니다.",
    ]


def test_decimal_precision_is_not_rounded() -> None:
    data = ProjectData(
        vendors=[Vendor("A"), Vendor("B")],
        items=[
            QuoteItem(
                "정밀 품목",
                "",
                "EA",
                Decimal("1.234567"),
                [Decimal("3500000.123456"), Decimal("3648888.987654")],
            )
        ],
    )
    expected = Decimal("1.234567") * Decimal("3500000.123456")
    assert data.items[0].amount_for(0) == expected.normalize()
    # 품목 계산은 정밀도를 유지하고, 기본 원단위 정책은 최종 합계에서만 반올림한다.
    assert data.vendor_supply_raw(0) == expected.normalize()
    assert data.vendor_supply_total(0) == Decimal("4320985")
    data.won_rounding = "keep"
    assert data.vendor_supply_total(0) == expected.normalize()
    assert data.vendor_vat(0) == (expected * Decimal("0.1")).normalize()


def test_manual_budget_and_won_rounding_modes() -> None:
    data = ProjectData(
        vendors=[Vendor("A")],
        items=[QuoteItem("품목", "", "EA", Decimal("1.5"), [Decimal("100.25")])],
        budget_mode="manual",
        manual_budget_supply=Decimal("999.6"),
    )
    assert data.budget_supply == Decimal("1000")
    assert data.budget_amount == Decimal("1100")
    data.won_rounding = "floor"
    assert data.budget_supply == Decimal("999")
    assert data.budget_vat == Decimal("99")
    data.won_rounding = "keep"
    assert data.budget_supply == Decimal("999.6")
    assert data.budget_vat == Decimal("99.96")


def test_ratio_display_is_one_decimal_place() -> None:
    data = ProjectData(
        vendors=[Vendor("A")],
        items=[QuoteItem("품목", "", "EA", 1, [100])],
        manual_budget_supply=Decimal("100"),
        budget_mode="manual",
    )
    assert data.ratio_text.endswith("%")
    assert len(data.ratio_text.split(".")[-1].rstrip("%")) == 1
    assert len(data.purchase_ratio_text.split(".")[-1].rstrip("%")) == 1
