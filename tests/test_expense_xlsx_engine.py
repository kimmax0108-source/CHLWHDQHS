from datetime import date
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from expense_statement_app.models import AllocationRow, ExpenseDocument, VendorAccount
from expense_statement_app.xlsx_engine import ExpenseXlsxEngine, MAIN_NS

NS = {
    "m": MAIN_NS,
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}


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


def cell_text(sheet: etree._Element, reference: str) -> str:
    cell = sheet.xpath(f'//m:c[@r="{reference}"]', namespaces=NS)[0]
    return "".join(cell.xpath(".//m:t/text()", namespaces=NS))


def test_export_starts_at_row_one_and_preserves_native_shapes(tmp_path: Path) -> None:
    engine = ExpenseXlsxEngine(Path("templates/expense_statement_template.xlsx"))
    output = tmp_path / "result.xlsx"
    engine.export(sample_document(), output)

    with ZipFile("templates/expense_statement_template.xlsx") as template, ZipFile(output) as result:
        assert result.testzip() is None
        assert result.read("xl/styles.xml") == template.read("xl/styles.xml")
        assert not any(name.startswith("xl/externalLinks/") for name in result.namelist())
        assert not any("vmlDrawing" in name for name in result.namelist())
        assert not any(name.startswith("xl/media/") for name in result.namelist())

        drawing = etree.fromstring(result.read("xl/drawings/drawing1.xml"))
        assert len(drawing) == 2
        assert not drawing.xpath("//*[local-name()='AlternateContent']")
        assert drawing.xpath("string(//xdr:oneCellAnchor/xdr:from/xdr:row)", namespaces=NS) == "0"

        sheet = etree.fromstring(result.read("xl/worksheets/sheet1.xml"))
        assert sheet.xpath("string(m:dimension/@ref)", namespaces=NS) == "X1:AT27"
        rows = sheet.xpath("//m:sheetData/m:row", namespaces=NS)
        assert [int(row.get("r")) for row in rows] == list(range(1, 28))
        assert not sheet.xpath("//m:legacyDrawing", namespaces=NS)
        assert sheet.xpath("string(//m:selection/@activeCell)", namespaces=NS) == "X1"
        assert cell_text(sheet, "AP5") == "監"
        assert cell_text(sheet, "AP6") == "事"
        assert cell_text(sheet, "AT5") == ""
        assert cell_text(sheet, "AT6") == ""
        merge_refs = {node.get("ref") for node in sheet.xpath("//m:mergeCell", namespaces=NS)}
        assert "AP5:AP6" not in merge_refs
        assert "AQ5:AS6" in merge_refs


def test_blank_site_accounts_and_rich_quantity_are_written(tmp_path: Path) -> None:
    output = tmp_path / "result.xlsx"
    ExpenseXlsxEngine(Path("templates/expense_statement_template.xlsx")).export(
        sample_document(), output
    )

    with ZipFile(output) as result:
        sheet = etree.fromstring(result.read("xl/worksheets/sheet1.xml"))
        assert cell_text(sheet, "Y8") == "08월 연신내 주상복합"
        assert cell_text(sheet, "Y9") == ""
        assert cell_text(sheet, "AC9") == "- 동경강업㈜"
        assert cell_text(sheet, "Y12").startswith("리소스뱅크㈜")
        assert cell_text(sheet, "Y13").startswith("동경강업㈜")
        assert cell_text(sheet, "Y15") == ""

        rich = sheet.xpath('//m:c[@r="AG8"]/m:is', namespaces=NS)[0]
        assert "".join(rich.xpath(".//m:t/text()", namespaces=NS)) == "187.114"
        assert rich.xpath('.//m:vertAlign[@val="superscript"]', namespaces=NS)
        assert cell_text(sheet, "AG15") == "213.090"
        assert sheet.xpath('string(//m:c[@r="Y6"]/m:v)', namespaces=NS) == "171549026"

        account_rich = sheet.xpath('//m:c[@r="Y12"]/m:is', namespaces=NS)[0]
        assert account_rich.xpath('.//m:sz[@val="11"]', namespaces=NS)


def test_third_account_gets_own_row_and_pushes_summary_down(tmp_path: Path) -> None:
    document = sample_document()
    document.allocations.append(
        AllocationRow(8, "", "환영철강㈜", Decimal("42.536"), 30_000_000)
    )
    document.vendor_accounts.append(
        VendorAccount("환영철강㈜", "하나", "176-910006-57004", "환영철강㈜")
    )
    output = tmp_path / "three_vendors.xlsx"
    ExpenseXlsxEngine(Path("templates/expense_statement_template.xlsx")).export(document, output)
    with ZipFile(output) as result:
        sheet = etree.fromstring(result.read("xl/worksheets/sheet1.xml"))
        assert sheet.xpath("string(m:dimension/@ref)", namespaces=NS) == "X1:AT28"
        assert cell_text(sheet, "Y13").startswith("리소스뱅크㈜")
        assert cell_text(sheet, "Y14").startswith("동경강업㈜")
        assert cell_text(sheet, "Y15").startswith("환영철강㈜")
        assert cell_text(sheet, "AC16") == "수량계"
        assert cell_text(sheet, "AG16") == "255.626"
        assert cell_text(sheet, "Y16") == ""

        drawing = etree.fromstring(result.read("xl/drawings/drawing1.xml"))
        line_from = drawing.xpath("string(//xdr:twoCellAnchor/xdr:from/xdr:row)", namespaces=NS)
        line_to = drawing.xpath("string(//xdr:twoCellAnchor/xdr:to/xdr:row)", namespaces=NS)
        assert line_from == "10"
        assert line_to == "15"
