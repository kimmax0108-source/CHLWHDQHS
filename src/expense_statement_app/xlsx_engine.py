from __future__ import annotations

import re
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from .models import BodyLine, ExpenseDocument, quantity_parts
from .resource import resource_path

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DCTERMS_NS = "http://purl.org/dc/terms/"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"m": MAIN_NS, "r": REL_NS, "xdr": DRAWING_NS, "a": A_NS, "mc": MC_NS}

SOURCE_FIRST_ROW = 201
SOURCE_LAST_ROW = 227
TARGET_MIN_COL = 24  # X
TARGET_MAX_COL = 46  # AT
BODY_START_ROW = 208
MIN_BODY_ROWS = 7
SUMMARY_BASE_ROW = 215
AMOUNT_DIGIT_COLS = ("AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS")
CELL_REF_RE = re.compile(r"(?P<col>\$?[A-Z]{1,3})(?P<abs>\$?)(?P<row>\d+)")


def _column_number(reference: str) -> int:
    letters = "".join(ch for ch in reference if ch.isalpha())
    value = 0
    for char in letters:
        value = value * 26 + ord(char.upper()) - 64
    return value


def _column_name(number: int) -> str:
    result = ""
    value = number
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _row_number(reference: str) -> int:
    digits = "".join(ch for ch in reference if ch.isdigit())
    return int(digits)


def _replace_row(reference: str, new_row: int) -> str:
    return "".join(ch for ch in reference if ch.isalpha() or ch == "$") + str(new_row)


def _range_bounds(reference: str) -> tuple[int, int, int, int]:
    first, _, last = reference.partition(":")
    last = last or first
    return (
        _column_number(first),
        _row_number(first),
        _column_number(last),
        _row_number(last),
    )


def _shift_range_rows(reference: str, start_row: int, delta: int) -> str:
    first, sep, last = reference.partition(":")

    def shift_one(cell_ref: str) -> str:
        row = _row_number(cell_ref)
        if row >= start_row:
            row += delta
        return _replace_row(cell_ref, row)

    shifted_first = shift_one(first)
    if not sep:
        return shifted_first
    return f"{shifted_first}:{shift_one(last)}"


def _rebase_range_rows(reference: str, offset: int) -> str:
    first, sep, last = reference.partition(":")

    def rebase_one(cell_ref: str) -> str:
        return _replace_row(cell_ref, _row_number(cell_ref) - offset)

    shifted_first = rebase_one(first)
    if not sep:
        return shifted_first
    return f"{shifted_first}:{rebase_one(last)}"


def _inside_source(reference: str) -> bool:
    min_col, min_row, max_col, max_row = _range_bounds(reference)
    return (
        TARGET_MIN_COL <= min_col <= max_col <= TARGET_MAX_COL
        and SOURCE_FIRST_ROW <= min_row <= max_row
    )


def _cell(sheet: etree._ElementTree, reference: str) -> etree._Element:
    matches = sheet.xpath(f'//m:c[@r="{reference}"]', namespaces=NS)
    if not matches:
        raise KeyError(f"템플릿 셀을 찾을 수 없습니다: {reference}")
    return matches[0]


def _clear_content(cell: etree._Element) -> None:
    for child_name in ("f", "v", "is"):
        for child in cell.findall(f"{{{MAIN_NS}}}{child_name}"):
            cell.remove(child)
    cell.attrib.pop("t", None)


def _set_inline_text(cell: etree._Element, text: str) -> None:
    _clear_content(cell)
    if not text:
        return
    cell.set("t", "inlineStr")
    inline = etree.SubElement(cell, f"{{{MAIN_NS}}}is")
    node = etree.SubElement(inline, f"{{{MAIN_NS}}}t")
    if text[:1].isspace() or text[-1:].isspace():
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    node.text = text


def _set_rich_text(cell: etree._Element, text: str, *, size: int = 11) -> None:
    _clear_content(cell)
    if not text:
        return
    cell.set("t", "inlineStr")
    inline = etree.SubElement(cell, f"{{{MAIN_NS}}}is")
    run = etree.SubElement(inline, f"{{{MAIN_NS}}}r")
    props = etree.SubElement(run, f"{{{MAIN_NS}}}rPr")
    etree.SubElement(props, f"{{{MAIN_NS}}}sz", val=str(size))
    etree.SubElement(props, f"{{{MAIN_NS}}}rFont", val="굴림")
    etree.SubElement(props, f"{{{MAIN_NS}}}family", val="3")
    etree.SubElement(props, f"{{{MAIN_NS}}}charset", val="129")
    node = etree.SubElement(run, f"{{{MAIN_NS}}}t")
    if text[:1].isspace() or text[-1:].isspace():
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    node.text = text


def _set_rich_quantity(cell: etree._Element, value: object) -> None:
    _clear_content(cell)
    main, fraction = quantity_parts(value)
    cell.set("t", "inlineStr")
    inline = etree.SubElement(cell, f"{{{MAIN_NS}}}is")

    first_run = etree.SubElement(inline, f"{{{MAIN_NS}}}r")
    first_props = etree.SubElement(first_run, f"{{{MAIN_NS}}}rPr")
    etree.SubElement(first_props, f"{{{MAIN_NS}}}rFont", val="굴림")
    etree.SubElement(first_props, f"{{{MAIN_NS}}}family", val="3")
    etree.SubElement(first_props, f"{{{MAIN_NS}}}charset", val="129")
    first_text = etree.SubElement(first_run, f"{{{MAIN_NS}}}t")
    first_text.text = main

    second_run = etree.SubElement(inline, f"{{{MAIN_NS}}}r")
    props = etree.SubElement(second_run, f"{{{MAIN_NS}}}rPr")
    etree.SubElement(props, f"{{{MAIN_NS}}}vertAlign", val="superscript")
    etree.SubElement(props, f"{{{MAIN_NS}}}sz", val="10")
    etree.SubElement(props, f"{{{MAIN_NS}}}rFont", val="굴림")
    etree.SubElement(props, f"{{{MAIN_NS}}}family", val="3")
    etree.SubElement(props, f"{{{MAIN_NS}}}charset", val="129")
    second_text = etree.SubElement(second_run, f"{{{MAIN_NS}}}t")
    second_text.text = fraction


def _set_numeric(cell: etree._Element, value: int | float) -> None:
    _clear_content(cell)
    node = etree.SubElement(cell, f"{{{MAIN_NS}}}v")
    node.text = str(value)


def _set_formula_with_cache(cell: etree._Element, formula: str, cached_value: int) -> None:
    _clear_content(cell)
    formula_node = etree.SubElement(cell, f"{{{MAIN_NS}}}f")
    formula_node.text = formula
    value_node = etree.SubElement(cell, f"{{{MAIN_NS}}}v")
    value_node.text = str(cached_value)


def _write_amount_digits(sheet: etree._ElementTree, row: int, amount: int) -> None:
    digits = str(max(0, int(amount))).rjust(len(AMOUNT_DIGIT_COLS), "0")[-len(AMOUNT_DIGIT_COLS) :]
    leading = True
    for col, digit in zip(AMOUNT_DIGIT_COLS, digits):
        cell = _cell(sheet, f"{col}{row}")
        if leading and digit == "0":
            _clear_content(cell)
            continue
        leading = False
        _set_numeric(cell, int(digit))


def _clear_amount_digits(sheet: etree._ElementTree, row: int) -> None:
    for col in AMOUNT_DIGIT_COLS:
        _clear_content(_cell(sheet, f"{col}{row}"))


def _spaced_name(name: str) -> str:
    compact = "".join(name.split())
    return "  ".join(compact)


class ExpenseXlsxEngine:
    def __init__(self, template_path: Path | None = None) -> None:
        self.template_path = template_path or resource_path("templates/expense_statement_template.xlsx")

    def export(self, document: ExpenseDocument, output_path: Path) -> Path:
        document.validate()
        if not self.template_path.exists():
            raise FileNotFoundError(f"엑셀 양식 파일을 찾을 수 없습니다: {self.template_path}")

        with ZipFile(self.template_path, "r") as source:
            files = {name: source.read(name) for name in source.namelist()}

        sheet = etree.ElementTree(etree.fromstring(files["xl/worksheets/sheet1.xml"]))
        workbook = etree.ElementTree(etree.fromstring(files["xl/workbook.xml"]))
        workbook_rels = etree.ElementTree(etree.fromstring(files["xl/_rels/workbook.xml.rels"]))
        content_types = etree.ElementTree(etree.fromstring(files["[Content_Types].xml"]))
        core = etree.ElementTree(etree.fromstring(files["docProps/core.xml"]))

        self._crop_sheet_to_source_page(sheet)
        layout = self._prepare_dynamic_rows(sheet, document)
        self._fill_document(sheet, document, layout)
        self._update_strikeout_shape(files, document, layout)
        self._sanitize_drawings(files)
        self._remove_external_links(files, workbook, workbook_rels, content_types)
        self._sanitize_workbook_metadata(workbook)
        self._rebase_page_to_row_one(files, sheet, workbook, layout["last_row"])
        self._replace_calc_chain(files)
        self._set_modified(core)

        files["xl/worksheets/sheet1.xml"] = etree.tostring(
            sheet, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        files["xl/workbook.xml"] = etree.tostring(
            workbook, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        files["xl/_rels/workbook.xml.rels"] = etree.tostring(
            workbook_rels, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        files["[Content_Types].xml"] = etree.tostring(
            content_types, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        files["docProps/core.xml"] = etree.tostring(
            core, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        output_path.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(output_path, "w", ZIP_DEFLATED, compresslevel=9) as target:
            for name, payload in files.items():
                target.writestr(name, payload)
        return output_path

    def _crop_sheet_to_source_page(self, sheet: etree._ElementTree) -> None:
        root = sheet.getroot()
        dimension = root.find(f"{{{MAIN_NS}}}dimension")
        if dimension is not None:
            dimension.set("ref", "X201:AT227")

        cols = root.find(f"{{{MAIN_NS}}}cols")
        if cols is not None:
            for col in list(cols):
                minimum = int(col.get("min", "1"))
                maximum = int(col.get("max", str(minimum)))
                overlap_min = max(minimum, TARGET_MIN_COL)
                overlap_max = min(maximum, TARGET_MAX_COL)
                if overlap_min > overlap_max:
                    cols.remove(col)
                else:
                    col.set("min", str(overlap_min))
                    col.set("max", str(overlap_max))

        sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
        if sheet_data is not None:
            for row in list(sheet_data):
                row_number = int(row.get("r", "0"))
                if not SOURCE_FIRST_ROW <= row_number <= SOURCE_LAST_ROW:
                    sheet_data.remove(row)
                    continue
                row.attrib.pop("spans", None)
                for cell in list(row):
                    reference = cell.get("r", "")
                    if not TARGET_MIN_COL <= _column_number(reference) <= TARGET_MAX_COL:
                        row.remove(cell)

        merge_cells = root.find(f"{{{MAIN_NS}}}mergeCells")
        if merge_cells is not None:
            for merge in list(merge_cells):
                if not _inside_source(merge.get("ref", "")):
                    merge_cells.remove(merge)
            merge_cells.set("count", str(len(merge_cells)))

        for tag in ("rowBreaks", "colBreaks", "legacyDrawing"):
            node = root.find(f"{{{MAIN_NS}}}{tag}")
            if node is not None:
                root.remove(node)

    def _prepare_dynamic_rows(
        self, sheet: etree._ElementTree, document: ExpenseDocument
    ) -> dict[str, int | list[int]]:
        planned = document.export_body_lines()
        body_count = max(MIN_BODY_ROWS, len(planned))
        extra_rows = body_count - MIN_BODY_ROWS
        if extra_rows:
            self._insert_rows(sheet, SUMMARY_BASE_ROW, extra_rows, template_row=213)

        summary_row = BODY_START_ROW + body_count
        total_row = summary_row + 1
        layout: dict[str, int | list[int]] = {
            "body_count": body_count,
            "body_rows": list(range(BODY_START_ROW, summary_row)),
            "summary_row": summary_row,
            "total_row": total_row,
            "date_row": 221 + extra_rows,
            "writer_row": 222 + extra_rows,
            "footer_row": 226 + extra_rows,
            "last_row": SOURCE_LAST_ROW + extra_rows,
            "extra_rows": extra_rows,
        }
        return layout

    def _insert_rows(
        self, sheet: etree._ElementTree, start_row: int, count: int, *, template_row: int
    ) -> None:
        if count <= 0:
            return
        root = sheet.getroot()
        sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
        if sheet_data is None:
            raise ValueError("템플릿 sheetData를 찾을 수 없습니다.")

        template = deepcopy(
            next(row for row in sheet_data if int(row.get("r", "0")) == template_row)
        )

        for row in reversed(list(sheet_data)):
            old_row = int(row.get("r", "0"))
            if old_row < start_row:
                continue
            new_row = old_row + count
            row.set("r", str(new_row))
            for cell in row.findall(f"{{{MAIN_NS}}}c"):
                cell.set("r", _replace_row(cell.get("r", ""), new_row))
                formula = cell.find(f"{{{MAIN_NS}}}f")
                if formula is not None and formula.text:
                    formula.text = self._shift_formula_rows(formula.text, start_row, count)

        for offset in range(count):
            new_row_number = start_row + offset
            new_row = deepcopy(template)
            new_row.set("r", str(new_row_number))
            new_row.attrib.pop("spans", None)
            for cell in new_row.findall(f"{{{MAIN_NS}}}c"):
                cell.set("r", _replace_row(cell.get("r", ""), new_row_number))
                _clear_content(cell)
            sheet_data.append(new_row)

        rows = sorted(list(sheet_data), key=lambda row: int(row.get("r", "0")))
        for row in list(sheet_data):
            sheet_data.remove(row)
        for row in rows:
            sheet_data.append(row)

        merge_cells = root.find(f"{{{MAIN_NS}}}mergeCells")
        if merge_cells is not None:
            for merge in merge_cells:
                merge.set("ref", _shift_range_rows(merge.get("ref", ""), start_row, count))

    @staticmethod
    def _shift_formula_rows(formula: str, start_row: int, delta: int) -> str:
        def replace(match: re.Match[str]) -> str:
            row = int(match.group("row"))
            if row >= start_row:
                row += delta
            return f"{match.group('col')}{match.group('abs')}{row}"

        return CELL_REF_RE.sub(replace, formula)

    def _sanitize_drawings(self, files: dict[str, bytes]) -> None:
        # The 2017 source page contains exactly two valid native shapes:
        # the title box and the curved strikeout line. Remove only broken camera/VML parts.
        for name in list(files):
            if name.startswith("xl/externalLinks/") or "vmlDrawing" in name:
                files.pop(name, None)
        for name in list(files):
            if name.startswith("xl/media/") and name.lower().endswith((".emf", ".wmf")):
                files.pop(name, None)

    def _copy_row_styles(self, sheet: etree._ElementTree, source_row: int, target_row: int) -> None:
        for col_number in range(TARGET_MIN_COL, TARGET_MAX_COL + 1):
            col = _column_name(col_number)
            source = _cell(sheet, f"{col}{source_row}")
            target = _cell(sheet, f"{col}{target_row}")
            style = source.get("s")
            if style is None:
                target.attrib.pop("s", None)
            else:
                target.set("s", style)

    def _remove_body_row_merges(self, sheet: etree._ElementTree, row: int) -> None:
        merge_cells = sheet.getroot().find(f"{{{MAIN_NS}}}mergeCells")
        if merge_cells is None:
            return
        for merge in list(merge_cells):
            min_col, min_row, max_col, max_row = _range_bounds(merge.get("ref", ""))
            if min_row == max_row == row and not (max_col < 25 or min_col > 33):
                merge_cells.remove(merge)
        merge_cells.set("count", str(len(merge_cells)))

    def _remove_exact_merge(self, sheet: etree._ElementTree, reference: str) -> None:
        merge_cells = sheet.getroot().find(f"{{{MAIN_NS}}}mergeCells")
        if merge_cells is None:
            return
        for merge in list(merge_cells):
            if merge.get("ref") == reference:
                merge_cells.remove(merge)
        merge_cells.set("count", str(len(merge_cells)))

    def _add_merge(self, sheet: etree._ElementTree, reference: str) -> None:
        merge_cells = sheet.getroot().find(f"{{{MAIN_NS}}}mergeCells")
        if merge_cells is None:
            merge_cells = etree.SubElement(sheet.getroot(), f"{{{MAIN_NS}}}mergeCells")
        if any(merge.get("ref") == reference for merge in merge_cells):
            return
        etree.SubElement(merge_cells, f"{{{MAIN_NS}}}mergeCell", ref=reference)
        merge_cells.set("count", str(len(merge_cells)))

    def _configure_body_row(self, sheet: etree._ElementTree, row: int, line: BodyLine | None) -> None:
        kind = line.kind if line is not None else "description"
        style_source = 208 if kind == "allocation" else (213 if kind == "account" else 211)
        self._copy_row_styles(sheet, style_source, row)
        self._remove_body_row_merges(sheet, row)
        if kind == "allocation":
            self._add_merge(sheet, f"Y{row}:AB{row}")
            self._add_merge(sheet, f"AC{row}:AF{row}")
        else:
            self._add_merge(sheet, f"Y{row}:AG{row}")

    def _fill_document(
        self, sheet: etree._ElementTree, document: ExpenseDocument, layout: dict[str, int | list[int]]
    ) -> None:
        total = document.total_amount
        _set_formula_with_cache(_cell(sheet, "Y205"), "Y206", total)
        _set_numeric(_cell(sheet, "Y206"), total)
        # The source form places 監/事 in the narrow column immediately after
        # the amount text block.  AP was vertically merged in the template;
        # split it so each character can stay on its own row, and leave the
        # far-right AT column blank exactly like the approved reference form.
        self._remove_exact_merge(sheet, "AP205:AP206")
        _set_inline_text(_cell(sheet, "AP205"), "監")
        _set_inline_text(_cell(sheet, "AP206"), "事")
        _clear_content(_cell(sheet, "AT205"))
        _clear_content(_cell(sheet, "AT206"))

        planned = document.export_body_lines()
        body_rows = list(layout["body_rows"])
        for index, row in enumerate(body_rows):
            line = planned[index] if index < len(planned) else None
            self._configure_body_row(sheet, row, line)
            for anchor in (f"Y{row}", f"AC{row}", f"AG{row}"):
                _clear_content(_cell(sheet, anchor))
            _clear_amount_digits(sheet, row)

            if line is None:
                continue
            if line.kind == "allocation" and line.allocation is not None:
                allocation = line.allocation
                _set_inline_text(_cell(sheet, f"Y{row}"), allocation.display_site)
                vendor = allocation.vendor_name.strip()
                _set_inline_text(_cell(sheet, f"AC{row}"), (vendor if vendor.startswith("-") else f"- {vendor}") if vendor else "")
                _set_rich_quantity(_cell(sheet, f"AG{row}"), allocation.quantity)
                _write_amount_digits(sheet, row, allocation.amount)
            elif line.kind == "account":
                _set_rich_text(_cell(sheet, f"Y{row}"), line.text, size=11)
            else:
                _set_inline_text(_cell(sheet, f"Y{row}"), line.text)

        summary_row = int(layout["summary_row"])
        _clear_content(_cell(sheet, f"Y{summary_row}"))
        _set_inline_text(_cell(sheet, f"AC{summary_row}"), "수량계")
        _set_rich_quantity(_cell(sheet, f"AG{summary_row}"), document.total_quantity)

        total_row = int(layout["total_row"])
        _write_amount_digits(sheet, total_row, total)

        date_row = int(layout["date_row"])
        writer_row = int(layout["writer_row"])
        footer_row = int(layout["footer_row"])
        date_text = (
            f"                 {document.written_date.year}.      "
            f"{document.written_date.month:02d} .   {document.written_date.day:02d}   ."
        )
        _set_inline_text(_cell(sheet, f"AD{date_row}"), date_text)
        writer = _spaced_name(document.writer)
        _set_inline_text(_cell(sheet, f"AD{writer_row}"), f"            請 求 人        {writer}  (印)")
        _set_inline_text(_cell(sheet, f"Y{footer_row}"), "양우F-152(O)")

    def _update_strikeout_shape(
        self, files: dict[str, bytes], document: ExpenseDocument, layout: dict[str, int | list[int]]
    ) -> None:
        name = "xl/drawings/drawing1.xml"
        if name not in files:
            return
        drawing = etree.ElementTree(etree.fromstring(files[name]))
        anchors = drawing.xpath("//xdr:twoCellAnchor", namespaces=NS)
        if not anchors:
            files[name] = etree.tostring(
                drawing, xml_declaration=True, encoding="UTF-8", standalone=True
            )
            return
        anchor = anchors[0]
        start_row_zero = BODY_START_ROW - 1 + len(document.active_allocations)
        end_row_zero = int(layout["summary_row"]) - 1
        from_node = anchor.find(f"{{{DRAWING_NS}}}from")
        to_node = anchor.find(f"{{{DRAWING_NS}}}to")
        if from_node is not None and to_node is not None:
            from_node.find(f"{{{DRAWING_NS}}}row").text = str(start_row_zero)
            to_node.find(f"{{{DRAWING_NS}}}row").text = str(max(start_row_zero + 1, end_row_zero))
        files[name] = etree.tostring(
            drawing, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    def _rebase_page_to_row_one(
        self,
        files: dict[str, bytes],
        sheet: etree._ElementTree,
        workbook: etree._ElementTree,
        last_row: int,
    ) -> None:
        offset = SOURCE_FIRST_ROW - 1
        root = sheet.getroot()
        sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
        if sheet_data is not None:
            for row in sheet_data:
                old_row = int(row.get("r", "0"))
                new_row = old_row - offset
                row.set("r", str(new_row))
                for cell in row.findall(f"{{{MAIN_NS}}}c"):
                    cell.set("r", _replace_row(cell.get("r", ""), new_row))
                    formula = cell.find(f"{{{MAIN_NS}}}f")
                    if formula is not None and formula.text:
                        formula.text = self._rebase_formula(formula.text, offset)

        merge_cells = root.find(f"{{{MAIN_NS}}}mergeCells")
        if merge_cells is not None:
            for merge in merge_cells:
                merge.set("ref", _rebase_range_rows(merge.get("ref", ""), offset))

        final_last_row = last_row - offset
        dimension = root.find(f"{{{MAIN_NS}}}dimension")
        if dimension is not None:
            dimension.set("ref", f"X1:AT{final_last_row}")

        view = sheet.xpath("//m:sheetView", namespaces=NS)[0]
        view.set("showGridLines", "0")
        view.set("view", "pageBreakPreview")
        view.set("topLeftCell", "X1")
        selections = view.xpath("./m:selection", namespaces=NS)
        if selections:
            selections[0].set("activeCell", "X1")
            selections[0].set("sqref", "X1")

        names = workbook.xpath("//m:definedName[@name='_xlnm.Print_Area']", namespaces=NS)
        if names:
            names[0].text = f"정기!$X$1:$AT${final_last_row}"
        calc_pr = workbook.xpath("//m:calcPr", namespaces=NS)
        if calc_pr:
            calc_pr[0].set("fullCalcOnLoad", "1")
            calc_pr[0].set("forceFullCalc", "1")

        drawing_name = "xl/drawings/drawing1.xml"
        if drawing_name in files:
            drawing = etree.ElementTree(etree.fromstring(files[drawing_name]))
            for row_node in drawing.xpath("//xdr:from/xdr:row | //xdr:to/xdr:row", namespaces=NS):
                row_node.text = str(int(row_node.text or "0") - offset)
            files[drawing_name] = etree.tostring(
                drawing, xml_declaration=True, encoding="UTF-8", standalone=True
            )

    @staticmethod
    def _rebase_formula(formula: str, offset: int) -> str:
        def replace(match: re.Match[str]) -> str:
            row = int(match.group("row"))
            if row >= SOURCE_FIRST_ROW:
                row -= offset
            return f"{match.group('col')}{match.group('abs')}{row}"

        return CELL_REF_RE.sub(replace, formula)

    def _remove_external_links(
        self,
        files: dict[str, bytes],
        workbook: etree._ElementTree,
        workbook_rels: etree._ElementTree,
        content_types: etree._ElementTree,
    ) -> None:
        for node in workbook.xpath("//m:externalReferences", namespaces=NS):
            node.getparent().remove(node)
        for relationship in list(workbook_rels.getroot()):
            if relationship.get("Type", "").endswith("/externalLink"):
                workbook_rels.getroot().remove(relationship)
        for override in list(content_types.getroot()):
            if override.get("PartName", "").startswith("/xl/externalLinks/"):
                content_types.getroot().remove(override)
        for name in list(files):
            if name.startswith("xl/externalLinks/"):
                files.pop(name, None)

    def _sanitize_workbook_metadata(self, workbook: etree._ElementTree) -> None:
        for node in workbook.xpath("//mc:AlternateContent", namespaces=NS):
            node.getparent().remove(node)
        for node in workbook.xpath("//*[local-name()='revisionPtr']"):
            node.getparent().remove(node)

    def _replace_calc_chain(self, files: dict[str, bytes]) -> None:
        root = etree.Element(f"{{{MAIN_NS}}}calcChain", nsmap={None: MAIN_NS})
        etree.SubElement(root, f"{{{MAIN_NS}}}c", r="Y5", i="1")
        files["xl/calcChain.xml"] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    def _set_modified(self, core: etree._ElementTree) -> None:
        nodes = core.xpath("//dcterms:modified", namespaces={"dcterms": DCTERMS_NS})
        if nodes:
            nodes[0].text = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace(
                "+00:00", "Z"
            )
