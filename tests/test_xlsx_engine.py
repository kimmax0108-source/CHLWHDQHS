from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path
from zipfile import ZipFile

from purchase_request_app.models import ProjectData, QuoteItem, Vendor, build_classification
from purchase_request_app.preview import build_preview_html
from purchase_request_app.xlsx_engine import NS_MAIN, XlsxTemplateEngine, qn


def cell_value(root: ET.Element, ref: str) -> str:
    for cell in root.findall(f".//{qn(NS_MAIN, 'c')}"):
        if cell.get("r") != ref:
            continue
        if cell.get("t") == "inlineStr":
            return "".join(t.text or "" for t in cell.iter(qn(NS_MAIN, "t")))
        formula = cell.find(qn(NS_MAIN, "f"))
        if formula is not None:
            return "=" + (formula.text or "")
        value = cell.find(qn(NS_MAIN, "v"))
        return value.text or "" if value is not None else ""
    raise AssertionError(f"Cell not found: {ref}")


def cell_style_id(root: ET.Element, ref: str) -> int:
    for cell in root.findall(f".//{qn(NS_MAIN, 'c')}"):
        if cell.get("r") == ref:
            return int(cell.get("s", "0") or 0)
    raise AssertionError(f"Cell not found: {ref}")


def style_horizontal(styles: ET.Element, style_id: int) -> str:
    xfs = styles.find(qn(NS_MAIN, "cellXfs"))
    assert xfs is not None
    xf = list(xfs)[style_id]
    alignment = xf.find(qn(NS_MAIN, "alignment"))
    return "" if alignment is None else alignment.get("horizontal", "")




def style_num_format(styles: ET.Element, style_id: int) -> str:
    xfs = styles.find(qn(NS_MAIN, "cellXfs"))
    assert xfs is not None
    xf = list(xfs)[style_id]
    num_fmt_id = xf.get("numFmtId", "0")
    num_fmts = styles.find(qn(NS_MAIN, "numFmts"))
    if num_fmts is not None:
        for node in list(num_fmts):
            if node.get("numFmtId") == num_fmt_id:
                return node.get("formatCode", "")
    return {"9": "0%", "10": "0.00%"}.get(num_fmt_id, "General")

def template_path() -> Path:
    return Path(__file__).resolve().parents[1] / "templates" / "purchase_request_3sheet_template.xlsx"


def make_data(vendor_count: int = 5) -> ProjectData:
    names = ["㈜신기콘테이너", "뉴월드인텍㈜", "대신콘테이너㈜", "세종컨테이너", "한성컨테이너"]
    totals = [8_200_000, 8_200_000, 9_700_000, 8_500_000, 8_600_000]
    vendors = [
        Vendor(
            name=names[i], phone=f"T 031)355-300{i} F 031)366-400{i}", manager=f"담당{i+1}",
            payment="기성결제 현금 (100%) 지급", delivery_place="세종 스마트시티",
            delivery_date="2026.08 ~ (18개월)"
        ) for i in range(vendor_count)
    ]
    data = ProjectData(
        site_name="세종 행정중심 복합도시 5-1생활권 L12BL 아파트건설공사",
        site_short="세종(5-1)", item_label="포세식 화장실", quote_title="포세식 화장실 임차",
        author="김 태 강", quote_date=date(2026, 7, 27), common_delivery_place="세종 스마트시티",
        vendors=vendors,
        items=[QuoteItem("포세식 화장실", "3 * 6", "EA", 2, totals[:vendor_count], group_title="포세식 화장실 임차(18개월)", group_sequence="1")],
        selected_vendor_index=0,
        classification=build_classification(2026, "세종(5-1)", 11), department="자 재 부",
        draft_date=date(2026, 7, 27), effective_date="결재후 즉시", drafter="김 태 강",
        purchase_title="포세식 화장실 임차", purchase_item_name="포세식 화장실(3 * 6, 2EA)",
        period_kind="임차기간", period="2026.08월 ~ 18개월", attachment="구매물품내역서",
        payment="기성결제 현금", body_text="세종 현장의 포세식 화장실 임차업체를 선정하여 품의하오니 결재하여 주시기 바랍니다.",
        statement_title="포세식 화장실 임차(18개월)",
    )
    data.sync_statement_from_quote()
    return data



def test_preview_and_excel_use_the_same_key_values(tmp_path: Path) -> None:
    """화면 미리보기와 최종 Excel이 같은 ProjectData 값을 사용하는지 검증한다."""

    output = tmp_path / "preview-parity.xlsx"
    data = make_data(3)
    data.note = "※철근자재: 당사구매 지급, 보증서 발급 1억원(가공장 부담)"
    XlsxTemplateEngine().export(data, template_path(), output)

    purchase_html = build_preview_html(data, "purchase")
    statement_html = build_preview_html(data, "statement")
    quote_html = build_preview_html(data, "quote")

    # 구매품의서: 원화 기호·가실행·금액·비율·특이사항이 모델값과 일치한다.
    assert f"₩ {data.purchase_budget_effective:,.0f}" in purchase_html
    assert f"₩ {data.purchase_contract_effective:,.0f}" in purchase_html
    assert data.purchase_ratio_text in purchase_html
    assert data.note in purchase_html

    # 내역서: 공급가·부가세·합계·가실행·비율이 실제 계산값과 일치한다.
    ratio_text = f"{data.statement_total / data.budget_amount:.1%}"
    for value in (
        data.statement_supply_total,
        data.statement_vat,
        data.statement_total,
        data.budget_supply,
        data.budget_vat,
        data.budget_amount,
    ):
        assert f"{value:,.0f}" in statement_html
    assert ratio_text in statement_html

    # 견적대비표: 품목, 수량, 업체명, 단가, 금액이 입력 모델과 일치한다.
    item = data.items[0]
    assert item.name in quote_html
    assert item.spec in quote_html
    assert f"{item.quantity:g}" in quote_html
    assert data.vendors[0].name in quote_html
    assert f"{item.unit_prices[0]:,.0f}" in quote_html
    assert f"{item.amount_for(0):,.0f}" in quote_html

    with ZipFile(output) as archive:
        purchase = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        statement = ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
        quote = ET.fromstring(archive.read("xl/worksheets/sheet3.xml"))

        assert cell_value(purchase, "F16") == format(data.purchase_budget_effective, "f")
        assert cell_value(purchase, "F17") == format(data.purchase_contract_effective, "f")
        assert cell_value(statement, "G58") == format(data.statement_total, "f")
        assert cell_value(statement, "H58") == format(data.budget_amount, "f")
        assert cell_value(quote, "B10") == item.name
        assert cell_value(quote, "C10") == item.spec
        assert cell_value(quote, "E10") == format(item.quantity, "f")
        assert cell_value(quote, "F10") == format(item.unit_prices[0], "f")
        assert cell_value(quote, "G10") == "=E10*F10"
        assert item.amount_for(0) == item.quantity * item.unit_prices[0]

def test_export_new_layout_rules_and_right_aligned_metadata(tmp_path: Path) -> None:
    output = tmp_path / "result.xlsx"
    data = make_data(5)
    XlsxTemplateEngine().export(data, template_path(), output)
    with ZipFile(output) as archive:
        assert archive.testzip() is None
        purchase = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        statement = ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
        quote = ET.fromstring(archive.read("xl/worksheets/sheet3.xml"))
        styles_xml = archive.read("xl/styles.xml")
        styles = ET.fromstring(styles_xml)
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        # v2.0.0은 수량·단가의 입력 정밀도를 유지하기 위해 안전한 사용자 지정
        # 숫자 형식을 styles.xml에 추가한다. 원본 네임스페이스와 기존 스타일은
        # 보존하면서 cellXfs 뒤에 새 스타일만 덧붙여야 한다.
        with ZipFile(template_path()) as template_archive:
            template_styles = ET.fromstring(template_archive.read("xl/styles.xml"))
        template_xfs = template_styles.find(qn(NS_MAIN, "cellXfs"))
        output_xfs = styles.find(qn(NS_MAIN, "cellXfs"))
        assert template_xfs is not None and output_xfs is not None
        assert len(list(output_xfs)) >= len(list(template_xfs))
        num_fmts = styles.find(qn(NS_MAIN, "numFmts"))
        assert num_fmts is not None
        format_codes = {node.get("formatCode") for node in list(num_fmts)}
        assert "#,##0" in format_codes
        assert "#,##0.000" in format_codes

        sheets = workbook.find(qn(NS_MAIN, "sheets"))
        assert sheets is not None
        assert [sheet.get("name") for sheet in sheets] == ["구매품의서", "내역서(포세식 화장실)", "견적대비표(포세식 화장실)"]

        # 구매품의서: 전화만, 구매품의서형 결제조건.
        assert cell_value(purchase, "K19") == "031)355-3000"
        assert cell_value(purchase, "F21") == "기성결제 현금 (100%) 지급"
        assert "₩" in style_num_format(styles, cell_style_id(purchase, "F16"))
        assert "₩" in style_num_format(styles, cell_style_id(purchase, "F17"))
        assert style_horizontal(styles, cell_style_id(purchase, "F16")) == "right"
        assert style_horizontal(styles, cell_style_id(purchase, "F17")) == "right"

        # 내역서: 대표 제목행 + 상세행, 상세행 가실행 공란, 하단에만 가실행.
        assert cell_value(statement, "B5") == "포세식 화장실 임차(18개월)"
        assert cell_value(statement, "A6") == "1"
        assert cell_value(statement, "B6") == "포세식 화장실"
        assert cell_value(statement, "F6") == "8200000"
        assert cell_value(statement, "G6") == "16400000"
        assert cell_value(statement, "H6") == ""
        assert cell_value(statement, "G58") == "18040000"
        assert cell_value(statement, "H58") == format(data.budget_amount, "f")
        assert style_num_format(styles, cell_style_id(statement, "I58")) == "0.0%"

        # 견적대비표: 제목행 + 상세행, 결제조건 기본형, 공통 납품장소.
        assert cell_value(quote, "A9") == "1"
        assert cell_value(quote, "B9") == "포세식 화장실 임차(18개월)"
        assert cell_value(quote, "A10") == ""
        assert cell_value(quote, "B10") == "포세식 화장실"
        assert cell_value(quote, "N30") == "기성결제 현금"
        assert cell_value(quote, "N31") == "세종 스마트시티"

        # 5개 업체일 때 M/N/O = 라벨/값/별첨, 라벨은 우측 정렬.
        assert cell_value(quote, "M2") == "작성자 :"
        assert cell_value(quote, "N2") == "김 태 강"
        assert cell_value(quote, "M3") == "작성일 :"
        assert cell_value(quote, "N3") == "2026.07.27"
        assert cell_value(quote, "O3") == "별첨(2)"
        assert style_horizontal(styles, cell_style_id(quote, "M2")) == "right"
        assert style_horizontal(styles, cell_style_id(quote, "M3")) == "right"
        assert style_horizontal(styles, cell_style_id(quote, "N2")) == "left"
        assert style_horizontal(styles, cell_style_id(quote, "O3")) == "right"

        xfs = styles.find(qn(NS_MAIN, "cellXfs"))
        assert xfs is not None
        style_count = len(list(xfs))
        for sheet_root in (purchase, statement, quote):
            for cell in sheet_root.findall(f".//{qn(NS_MAIN, 'c')}"):
                assert int(cell.get("s", "0") or 0) < style_count


def test_export_two_vendors_reduces_quote_width_and_keeps_metadata(tmp_path: Path) -> None:
    output = tmp_path / "two.xlsx"
    XlsxTemplateEngine().export(make_data(2), template_path(), output)
    with ZipFile(output) as archive:
        quote = ET.fromstring(archive.read("xl/worksheets/sheet3.xml"))
        dimension = quote.find(qn(NS_MAIN, "dimension"))
        assert dimension is not None and dimension.get("ref") == "A1:I33"
        assert cell_value(quote, "G2") == "작성자 :"
        assert cell_value(quote, "H2") == "김 태 강"
        assert cell_value(quote, "I3") == "별첨(2)"


def test_export_removes_invalid_office_extension_metadata(tmp_path: Path) -> None:
    output = tmp_path / "clean.xlsx"
    XlsxTemplateEngine().export(make_data(3), template_path(), output)
    with ZipFile(output) as archive:
        assert archive.testzip() is None
        for name in ("xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml"):
            xml_text = archive.read(name).decode("utf-8")
            ET.fromstring(xml_text)
            assert "mc:Ignorable" not in xml_text
            assert "revisionPtr" not in xml_text
            assert "spreadsheetml/2014/revision" not in xml_text
            assert "spreadsheetml/2009/9/ac" not in xml_text


def test_export_package_relationship_targets_exist(tmp_path: Path) -> None:
    output = tmp_path / "package.xlsx"
    XlsxTemplateEngine().export(make_data(5), template_path(), output)
    relationship_namespace = "http://schemas.openxmlformats.org/package/2006/relationships"
    with ZipFile(output) as archive:
        names = set(archive.namelist())
        for name in sorted(names):
            if not name.endswith(".xml") and not name.endswith(".rels"):
                continue
            ET.fromstring(archive.read(name))
            if not name.endswith(".rels"):
                continue
            root = ET.fromstring(archive.read(name))
            if name == "_rels/.rels":
                base_dir = ""
            else:
                rels_dir, rels_file = posixpath.split(name)
                source_dir = posixpath.dirname(rels_dir)
                source_file = rels_file.removesuffix(".rels")
                base_dir = posixpath.dirname(posixpath.join(source_dir, source_file))
            for relationship in root.findall(f"{{{relationship_namespace}}}Relationship"):
                if relationship.get("TargetMode") == "External":
                    continue
                target = relationship.get("Target", "")
                resolved = posixpath.normpath(posixpath.join(base_dir, target)).lstrip("/")
                assert resolved in names, f"Missing relationship target: {name} -> {resolved}"


def test_statement_auto_width_and_integer_number_format(tmp_path: Path) -> None:
    output = tmp_path / "widths.xlsx"
    data = make_data(3)
    data.statement_items[0].name = "철근 공장가공 및 현장 반입 장기 품목명"
    data.statement_items[0].spec = "나사가공,D32 장기 규격 표기"
    data.statement_items[0].quantity = 1644
    data.statement_items.append(
        data.statement_items[0].__class__(
            number="2", name="소수 수량", spec="보통", unit="TON",
            quantity=__import__("decimal").Decimal("1149.313"), unit_price=47000, amount=54017711
        )
    )
    XlsxTemplateEngine().export(data, template_path(), output)
    with ZipFile(output) as archive:
        statement = ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
        styles = ET.fromstring(archive.read("xl/styles.xml"))
        cols = statement.find(qn(NS_MAIN, "cols"))
        assert cols is not None
        widths = {int(node.get("min", "0")): float(node.get("width", "0")) for node in list(cols) if int(node.get("min", "0")) <= 9}
        assert widths[2] >= 11
        assert widths[3] >= 16
        assert widths[5] >= 9
        xfs = styles.find(qn(NS_MAIN, "cellXfs"))
        num_fmts = styles.find(qn(NS_MAIN, "numFmts"))
        assert xfs is not None and num_fmts is not None
        fmt_by_id = {node.get("numFmtId"): node.get("formatCode") for node in list(num_fmts)}
        integer_style = list(xfs)[cell_style_id(statement, "E6")]
        decimal_style = list(xfs)[cell_style_id(statement, "E7")]
        assert fmt_by_id.get(integer_style.get("numFmtId")) == "#,##0"
        assert fmt_by_id.get(decimal_style.get("numFmtId")) == "#,##0.000"


def test_purchase_note_is_written_immediately_after_last_body_line(tmp_path: Path) -> None:
    output = tmp_path / "note.xlsx"
    data = make_data(3)
    data.body_text = "25-A-00부대 시설공사(1450)의 철근 공장가공 업체를 선정하여 품의하오니 결재하여 주시기 바랍니다."
    data.note = "※철근자재: 당사구매 지급, 보증서 발급 1억원(가공장 부담)"
    XlsxTemplateEngine().export(data, template_path(), output)
    with ZipFile(output) as archive:
        purchase = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        assert "업체를 선정하여" in cell_value(purchase, "C22")
        assert "품의하오니" in cell_value(purchase, "C23")
        assert cell_value(purchase, "C24") == data.note
        assert cell_value(purchase, "C25") == ""
        assert cell_style_id(purchase, "C24") == cell_style_id(purchase, "C23")


def test_quote_excel_columns_expand_for_long_names_quantity_and_vendor(tmp_path: Path) -> None:
    output = tmp_path / "quote-widths.xlsx"
    data = make_data(3)
    data.vendors[0].name = "동일철강산업주식회사 세종공장 장기 업체명"
    data.items[0].name = "철근연결 이음재-커플러 초장기 품명 전체 표시"
    data.items[0].spec = "나사가공,D32 특수규격 전체표시"
    data.items[0].quantity = __import__("decimal").Decimal("1149.313")
    data.items[0].unit_prices[0] = __import__("decimal").Decimal("123456789")
    XlsxTemplateEngine().export(data, template_path(), output)
    with ZipFile(output) as archive:
        quote = ET.fromstring(archive.read("xl/worksheets/sheet3.xml"))
        cols = quote.find(qn(NS_MAIN, "cols"))
        assert cols is not None
        widths = {int(node.get("min", "0")): float(node.get("width", "0")) for node in list(cols)}
        assert widths[2] > 20.0  # 품명
        assert widths[3] > 16.0  # 규격
        assert widths[5] >= 9.0  # 수량
        assert widths[6] >= 11.5  # 단가
        assert widths[7] >= 13.5  # 금액
        assert widths[6] + widths[7] >= 28.0  # 병합된 업체명 전체 폭
