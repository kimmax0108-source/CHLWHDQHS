from datetime import date

from purchase_request_app.models import ProjectData, QuoteItem, Vendor, build_classification
from purchase_request_app.preview import build_preview_html


def make_data() -> ProjectData:
    data = ProjectData(
        site_name="세종 테스트 현장",
        site_short="세종",
        item_label="포세식화장실",
        quote_title="포세식 화장실 임차",
        author="김테스트",
        quote_date=date(2026, 7, 27),
        common_delivery_place="세종 스마트시티",
        vendors=[
            Vendor("A업체", phone="T 010-1111-2222 F 02-3333-4444", payment="기성결제 현금", delivery_place="세종 스마트시티"),
            Vendor("B업체", delivery_place="세종 스마트시티"),
        ],
        items=[QuoteItem("포세식 화장실", "3*6", "EA", 2, [8_000_000, 9_000_000], group_title="포세식 화장실 임차(18개월)", group_sequence="1")],
        classification=build_classification(2026, "세종", 11),
        draft_date=date(2026, 7, 27), drafter="김테스트",
        purchase_title="포세식 화장실 임차", purchase_item_name="포세식 화장실",
        payment="기성결제 현금", body_text="세종 테스트 현장의 포세식 화장실을 임차하고자 품의합니다.",
    )
    data.sync_statement_from_quote()
    return data


def test_preview_contains_all_three_documents_and_new_rules() -> None:
    data = make_data()
    purchase = build_preview_html(data, "purchase")
    assert "자재구매품의서" in purchase
    assert "010-1111-2222" in purchase
    assert "02-3333-4444" not in purchase
    assert "기성결제 현금 (100%) 지급" in purchase
    statement = build_preview_html(data, "statement")
    assert "구매물품내역서" in statement
    assert "포세식 화장실 임차(18개월)" in statement
    quote = build_preview_html(data, "quote")
    assert "견적대비표" in quote
    assert "label-right" in quote
    assert "attachment-right" in quote
    assert "포세식 화장실 임차(18개월)" in quote
    assert "기성결제 현금" in quote
