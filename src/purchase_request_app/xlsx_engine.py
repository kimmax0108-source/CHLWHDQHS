from __future__ import annotations

import re
import unicodedata
import xml.etree.ElementTree as ET
from copy import deepcopy
from datetime import UTC, datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZIP_DEFLATED, ZipFile

from .models import (
    ProjectData,
    decimal_text,
    format_money,
    parse_decimal,
    ensure_xlsx_suffix,
    purchase_payment_text,
    wrap_body_text,
)

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
NS_DC = "http://purl.org/dc/elements/1.1/"
NS_DCTERMS = "http://purl.org/dc/terms/"
NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"
NS_XML = "http://www.w3.org/XML/1998/namespace"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
NS_X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
NS_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
NS_XR2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
NS_XR3 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"
NS_XR6 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6"
NS_XR9 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision9"
NS_XR10 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"
NS_X16R2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"
NS_XCALCF = "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures"

EXTENSION_NAMESPACES = {
    NS_X14AC,
    NS_X15,
    NS_XR,
    NS_XR2,
    NS_XR3,
    NS_XR6,
    NS_XR9,
    NS_XR10,
    NS_X16R2,
    NS_XCALCF,
}

for prefix, namespace in (
    ("x", NS_MAIN),
    ("r", NS_REL),
    ("cp", NS_CP),
    ("dc", NS_DC),
    ("dcterms", NS_DCTERMS),
    ("xsi", NS_XSI),
    ("mc", NS_MC),
):
    ET.register_namespace(prefix, namespace)

CELL_RE = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$")


def qn(namespace: str, tag: str) -> str:
    return f"{{{namespace}}}{tag}"


def col_to_index(column: str) -> int:
    result = 0
    for char in column.upper():
        result = result * 26 + ord(char) - 64
    return result


def index_to_col(index: int) -> str:
    if index < 1:
        raise ValueError("열 번호는 1 이상이어야 합니다.")
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def split_cell_ref(ref: str) -> tuple[int, int]:
    match = CELL_RE.match(ref.replace("$", ""))
    if not match:
        raise ValueError(f"잘못된 셀 주소: {ref}")
    return col_to_index(match.group(1)), int(match.group(2))


def shift_ref(ref: str, delta: int) -> str:
    column, row = split_cell_ref(ref)
    return f"{index_to_col(column + delta)}{row}"


def xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def xml_bytes_default(root: ET.Element, namespace: str) -> bytes:
    ET.register_namespace("", namespace)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


class XlsxTemplateEngine:
    """원본 3개 시트 양식을 유지하면서 값과 업체 열만 OOXML로 수정한다."""

    PURCHASE_SHEET = "xl/worksheets/sheet1.xml"
    STATEMENT_SHEET = "xl/worksheets/sheet2.xml"
    QUOTE_SHEET = "xl/worksheets/sheet3.xml"
    STYLES = "xl/styles.xml"

    def export(
        self,
        data: ProjectData,
        template_path: str | Path,
        output_path: str | Path,
    ) -> Path:
        errors = data.validate()
        if errors:
            raise ValueError("\n".join(errors))

        template = Path(template_path)
        output = ensure_xlsx_suffix(output_path)
        if not template.exists():
            raise FileNotFoundError(f"양식 파일을 찾을 수 없습니다: {template}")
        output.parent.mkdir(parents=True, exist_ok=True)

        with ZipFile(template, "r") as source:
            styles_bytes = source.read(self.STYLES)
            self._prepare_precision_styles(styles_bytes)
            purchase_root = ET.fromstring(source.read(self.PURCHASE_SHEET))
            statement_root = ET.fromstring(source.read(self.STATEMENT_SHEET))
            quote_root = ET.fromstring(source.read(self.QUOTE_SHEET))

            self._write_purchase(purchase_root, data)
            self._write_statement(statement_root, data)
            quote_last_col = self._rebuild_quote_sheet(quote_root, data)

            # Excel의 복구 경고를 유발하는 공동편집/개정 메타데이터를 제거한다.
            for worksheet_root in (purchase_root, statement_root, quote_root):
                self._strip_extension_metadata(worksheet_root)

            replacements = {
                self.PURCHASE_SHEET: xml_bytes(purchase_root),
                self.STATEMENT_SHEET: xml_bytes(statement_root),
                self.QUOTE_SHEET: xml_bytes(quote_root),
                "xl/workbook.xml": self._clean_workbook(
                    source.read("xl/workbook.xml"), data, quote_last_col
                ),
                "xl/_rels/workbook.xml.rels": self._clean_workbook_rels(
                    source.read("xl/_rels/workbook.xml.rels")
                ),
                "[Content_Types].xml": self._clean_content_types(
                    source.read("[Content_Types].xml")
                ),
                "docProps/core.xml": self._clean_core_properties(
                    source.read("docProps/core.xml")
                ),
                self.STYLES: self._augment_precision_styles(styles_bytes),
            }

            excluded = (
                "xl/externalLinks/",
                "xl/calcChain.xml",
            )
            with NamedTemporaryFile(
                prefix="purchase_request_", suffix=".xlsx", delete=False, dir=output.parent
            ) as temp:
                temp_path = Path(temp.name)
            try:
                with ZipFile(temp_path, "w", compression=ZIP_DEFLATED) as target:
                    for info in source.infolist():
                        if info.filename in replacements:
                            target.writestr(info, replacements[info.filename])
                        elif any(info.filename.startswith(prefix) for prefix in excluded):
                            continue
                        else:
                            target.writestr(info, source.read(info.filename))
                temp_path.replace(output)
            except Exception:
                temp_path.unlink(missing_ok=True)
                raise
        return output

    # ------------------------------------------------------------------
    # Cell helpers
    # ------------------------------------------------------------------
    def _sheet_data(self, root: ET.Element) -> ET.Element:
        sheet_data = root.find(qn(NS_MAIN, "sheetData"))
        if sheet_data is None:
            raise ValueError("양식에 sheetData가 없습니다.")
        return sheet_data

    def _find_row(self, root: ET.Element, row_number: int) -> ET.Element | None:
        for row in self._sheet_data(root).findall(qn(NS_MAIN, "row")):
            if int(row.get("r", "0")) == row_number:
                return row
        return None

    def _find_cell(self, root: ET.Element, ref: str) -> ET.Element | None:
        _, row_number = split_cell_ref(ref)
        row = self._find_row(root, row_number)
        if row is None:
            return None
        for cell in row.findall(qn(NS_MAIN, "c")):
            if cell.get("r") == ref:
                return cell
        return None

    def _ensure_cell(
        self,
        root: ET.Element,
        ref: str,
        style_source: ET.Element | None = None,
    ) -> ET.Element:
        existing = self._find_cell(root, ref)
        if existing is not None:
            return existing
        target_col, target_row = split_cell_ref(ref)
        sheet_data = self._sheet_data(root)
        row = self._find_row(root, target_row)
        if row is None:
            row = ET.Element(qn(NS_MAIN, "row"), {"r": str(target_row)})
            inserted = False
            for index, current in enumerate(sheet_data.findall(qn(NS_MAIN, "row"))):
                if int(current.get("r", "0")) > target_row:
                    sheet_data.insert(index, row)
                    inserted = True
                    break
            if not inserted:
                sheet_data.append(row)
        attrs = {"r": ref}
        if style_source is not None and style_source.get("s") is not None:
            attrs["s"] = style_source.get("s", "")
        cell = ET.Element(qn(NS_MAIN, "c"), attrs)
        inserted = False
        for index, current in enumerate(row.findall(qn(NS_MAIN, "c"))):
            col, _ = split_cell_ref(current.get("r", "A1"))
            if col > target_col:
                row.insert(index, cell)
                inserted = True
                break
        if not inserted:
            row.append(cell)
        return cell

    @staticmethod
    def _clear_cell(cell: ET.Element) -> None:
        for child in list(cell):
            cell.remove(child)
        cell.attrib.pop("t", None)

    def _set_text(self, root: ET.Element, ref: str, value: str | None) -> None:
        cell = self._ensure_cell(root, ref)
        self._clear_cell(cell)
        text = "" if value is None else str(value)
        if not text:
            return
        cell.set("t", "inlineStr")
        inline = ET.SubElement(cell, qn(NS_MAIN, "is"))
        node = ET.SubElement(inline, qn(NS_MAIN, "t"))
        node.set(qn(NS_XML, "space"), "preserve")
        node.text = text

    def _set_number(
        self,
        root: ET.Element,
        ref: str,
        value: object,
        *,
        precision: bool = True,
    ) -> None:
        cell = self._ensure_cell(root, ref)
        self._clear_cell(cell)
        if precision:
            self._apply_precision_style(cell, value)
        node = ET.SubElement(cell, qn(NS_MAIN, "v"))
        node.text = decimal_text(value)

    def _set_formula(
        self,
        root: ET.Element,
        ref: str,
        formula: str,
        cached: object | None,
        *,
        precision: bool = True,
    ) -> None:
        cell = self._ensure_cell(root, ref)
        self._clear_cell(cell)
        if precision and cached is not None:
            self._apply_precision_style(cell, cached)
        f = ET.SubElement(cell, qn(NS_MAIN, "f"))
        f.text = formula
        v = ET.SubElement(cell, qn(NS_MAIN, "v"))
        if cached is not None:
            v.text = decimal_text(cached)

    def _prepare_precision_styles(self, styles_bytes: bytes) -> None:
        text = styles_bytes.decode("utf-8")
        cell_xfs = re.search(r'<cellXfs\b[^>]*count="(\d+)"[^>]*>', text)
        if cell_xfs is None:
            raise ValueError("양식 styles.xml에 cellXfs가 없습니다.")
        self._precision_style_start = int(cell_xfs.group(1))
        self._precision_style_map: dict[tuple[int, int], int] = {}
        self._precision_style_reverse: dict[int, int] = {}
        custom_ids = [int(value) for value in re.findall(r'<numFmt\b[^>]*numFmtId="(\d+)"', text)]
        self._precision_num_fmt_start = max([182, *custom_ids]) + 1

    def _apply_precision_style(self, cell: ET.Element, value: object) -> None:
        decimal_value = parse_decimal(value)
        # 정수는 '#,##0', 소수는 사용자가 입력한 자릿수만큼 0을 유지한다.
        # Excel의 유효숫자 한계를 고려해 표시 자릿수는 최대 15자리로 제한하지만
        # 실제 셀 값 자체는 반올림하거나 잘라내지 않는다.
        scale = max(0, min(15, -decimal_value.as_tuple().exponent))
        original = int(cell.get("s", "0"))
        if original >= self._precision_style_start:
            original = self._precision_style_reverse.get(original, 0)
        key = (original, scale)
        mapped = self._precision_style_map.get(key)
        if mapped is None:
            mapped = self._precision_style_start + len(self._precision_style_map)
            self._precision_style_map[key] = mapped
            self._precision_style_reverse[mapped] = original
        cell.set("s", str(mapped))

    def _augment_precision_styles(self, styles_bytes: bytes) -> bytes:
        if not getattr(self, "_precision_style_map", None):
            return styles_bytes
        text = styles_bytes.decode("utf-8")
        scales = sorted({0, 3, *{scale for _original, scale in self._precision_style_map}})
        fmt_ids = {scale: self._precision_num_fmt_start + index for index, scale in enumerate(scales)}

        num_nodes = []
        for scale in scales:
            code = "#,##0" if scale == 0 else "#,##0." + ("0" * scale)
            num_nodes.append(f'<numFmt numFmtId="{fmt_ids[scale]}" formatCode="{code}"/>')
        nodes_text = "".join(num_nodes)
        num_match = re.search(
            r'(<numFmts\b[^>]*count=")(\d+)("[^>]*>)(.*?)(</numFmts>)',
            text,
            flags=re.DOTALL,
        )
        if num_match:
            count = int(num_match.group(2)) + len(num_nodes)
            replacement = (
                num_match.group(1) + str(count) + num_match.group(3)
                + num_match.group(4) + nodes_text + num_match.group(5)
            )
            text = text[: num_match.start()] + replacement + text[num_match.end() :]
        else:
            opening = re.search(r'<styleSheet\b[^>]*>', text)
            if opening is None:
                raise ValueError("양식 styles.xml의 styleSheet를 찾을 수 없습니다.")
            insertion = f'<numFmts count="{len(num_nodes)}">{nodes_text}</numFmts>'
            text = text[: opening.end()] + insertion + text[opening.end() :]

        xfs_match = re.search(
            r'(<cellXfs\b[^>]*count=")(\d+)("[^>]*>)(.*?)(</cellXfs>)',
            text,
            flags=re.DOTALL,
        )
        if xfs_match is None:
            raise ValueError("양식 styles.xml의 cellXfs를 찾을 수 없습니다.")
        old_count = int(xfs_match.group(2))
        body = xfs_match.group(4)
        xfs = re.findall(r'<xf\b[^>]*?/>|<xf\b[^>]*>.*?</xf>', body, flags=re.DOTALL)
        clones: list[str] = []
        for (original, scale), mapped in sorted(self._precision_style_map.items(), key=lambda item: item[1]):
            if mapped != old_count + len(clones) or not (0 <= original < len(xfs)):
                raise ValueError("정밀 숫자 스타일 매핑이 올바르지 않습니다.")
            clone = xfs[original]
            fmt_id = fmt_ids[scale]
            if 'numFmtId=' in clone:
                clone = re.sub(r'numFmtId="\d+"', f'numFmtId="{fmt_id}"', clone, count=1)
            else:
                clone = clone.replace('<xf ', f'<xf numFmtId="{fmt_id}" ', 1)
            if 'applyNumberFormat=' in clone:
                clone = re.sub(r'applyNumberFormat="[^"]*"', 'applyNumberFormat="1"', clone, count=1)
            else:
                clone = clone.replace('/>', ' applyNumberFormat="1"/>', 1) if clone.rstrip().endswith('/>') else clone.replace('>', ' applyNumberFormat="1">', 1)
            if 'applyAlignment=' in clone:
                clone = re.sub(r'applyAlignment="[^"]*"', 'applyAlignment="1"', clone, count=1)
            else:
                clone = clone.replace('/>', ' applyAlignment="1"/>', 1) if clone.rstrip().endswith('/>') else clone.replace('>', ' applyAlignment="1">', 1)
            if '<alignment ' in clone:
                if 'shrinkToFit=' in clone:
                    clone = re.sub(r'shrinkToFit="[^"]*"', 'shrinkToFit="1"', clone, count=1)
                else:
                    clone = clone.replace('<alignment ', '<alignment shrinkToFit="1" ', 1)
            elif '<alignment/>' in clone:
                clone = clone.replace('<alignment/>', '<alignment shrinkToFit="1"/>', 1)
            elif clone.rstrip().endswith('/>'):
                clone = re.sub(r'/>(\s*)$', r'><alignment shrinkToFit="1"/></xf>\1', clone, count=1)
            else:
                clone = clone.replace('</xf>', '<alignment shrinkToFit="1"/></xf>', 1)
            clones.append(clone)
        new_count = old_count + len(clones)
        replacement = (
            xfs_match.group(1) + str(new_count) + xfs_match.group(3)
            + body + ''.join(clones) + xfs_match.group(5)
        )
        text = text[: xfs_match.start()] + replacement + text[xfs_match.end() :]
        return text.encode("utf-8")

    def _clear_range(self, root: ET.Element, min_col: int, max_col: int, min_row: int, max_row: int) -> None:
        for row in self._sheet_data(root).findall(qn(NS_MAIN, "row")):
            row_number = int(row.get("r", "0"))
            if not min_row <= row_number <= max_row:
                continue
            for cell in row.findall(qn(NS_MAIN, "c")):
                col, _ = split_cell_ref(cell.get("r", "A1"))
                if min_col <= col <= max_col:
                    self._clear_cell(cell)

    def _copy_cell_style(self, root: ET.Element, source_ref: str, target_ref: str) -> None:
        source = self._find_cell(root, source_ref)
        target = self._ensure_cell(root, target_ref)
        if source is not None and source.get("s") is not None:
            target.set("s", source.get("s", "0"))

    def _copy_row_styles(
        self,
        root: ET.Element,
        source_row: int,
        target_row: int,
        min_col: int,
        max_col: int,
    ) -> None:
        for col_index in range(min_col, max_col + 1):
            column = index_to_col(col_index)
            self._copy_cell_style(root, f"{column}{source_row}", f"{column}{target_row}")

    def _apply_style_from_template_cell(
        self,
        root: ET.Element,
        target_ref: str,
        template_cell: ET.Element | None,
    ) -> None:
        """Copy a known-valid style id without rewriting styles.xml.

        The template style table remains byte-for-byte unchanged. Rebuilding cellXfs with
        ElementTree can alter Office namespace prefixes or xf child order and may make
        desktop Excel repair the workbook.
        """

        target = self._ensure_cell(root, target_ref)
        if template_cell is not None and template_cell.get("s") is not None:
            target.set("s", template_cell.get("s", "0"))

    def _strip_extension_metadata(self, root: ET.Element) -> None:
        """불필요한 Office 개정/공동편집 확장 메타데이터를 제거한다.

        원본 양식에는 xr/x14ac 접두사를 사용하는 속성과 mc:Ignorable 목록이
        포함되어 있다. XML 재직렬화 과정에서 사용되지 않는 namespace 선언은
        사라질 수 있으므로, Excel에서 복구 경고가 발생하지 않도록 비필수 확장
        요소와 속성을 제거하고 순수 OOXML 본문만 남긴다.
        """

        ignorable = qn(NS_MC, "Ignorable")
        for element in root.iter():
            element.attrib.pop(ignorable, None)
            for attr_name in list(element.attrib):
                if not attr_name.startswith("{"):
                    continue
                namespace = attr_name[1:].split("}", 1)[0]
                if namespace in EXTENSION_NAMESPACES:
                    del element.attrib[attr_name]

            for child in list(element):
                if not child.tag.startswith("{"):
                    continue
                namespace = child.tag[1:].split("}", 1)[0]
                if namespace in EXTENSION_NAMESPACES or namespace == NS_MC:
                    element.remove(child)

    # ------------------------------------------------------------------
    # 구매품의서
    # ------------------------------------------------------------------
    def _write_purchase(self, root: ET.Element, data: ProjectData) -> None:
        vendor = data.selected_vendor
        common = {
            "C3": data.classification,
            "I3": data.department,
            "C4": "",  # 문서번호는 사용하지 않음
            "C5": data.draft_date_text,
            "I5": data.effective_date,
            "C6": data.drafter,
            "E6": data.approval_note,
            "C13": data.purchase_title,
            "F14": data.purchase_site_effective,
            "F15": data.purchase_item_name,
            "I16": "(부가세 포함)",
            "I17": "(부가세 포함)",
            "F18": data.period,
            "F19": data.purchase_vendor_effective,
            "I19": "전화",
            "K19": data.purchase_phone_effective,
            "F20": data.attachment,
            "F21": purchase_payment_text(data.payment or (vendor.payment if vendor else "")),
        }
        for ref, value in common.items():
            self._set_text(root, ref, value)
        period_labels = {
            "계약기간": "계   약    기    간:",
            "임차기간": "임   차    기    간:",
            "납품일자": "납   품    일    자:",
        }
        self._set_text(root, "C18", period_labels.get(data.period_kind, f"{data.period_kind}:"))
        # 구매품의서 가실행/금액 셀은 회계 형식을 유지한다. 원화 기호는
        # 셀 왼쪽에 고정되고 금액 숫자는 오른쪽 정렬되어야 한다.
        # 정밀도 스타일로 덮어쓰면 이 배치와 ₩ 표시가 사라진다.
        self._set_number(root, "F16", data.purchase_budget_effective, precision=False)
        self._set_number(root, "F17", data.purchase_contract_effective, precision=False)
        self._set_formula(
            root,
            "M17",
            'IF(F16=0,"",F17/F16)',
            data.purchase_ratio_effective if data.purchase_ratio_effective is not None else None,
            precision=False,
        )
        lines = wrap_body_text(data.body_text, width=42, max_lines=3)
        # 본문과 추가 문구 영역을 먼저 비운 뒤, 실제 본문 줄 수만큼 연속 배치한다.
        # 본문이 두 줄이면 추가 문구는 C24에 들어가며 빈 행을 만들지 않는다.
        for ref in ("C22", "C23", "C24", "C25"):
            self._set_text(root, ref, "")
        for index, line in enumerate(lines):
            self._set_text(root, f"C{22 + index}", line)
        note_row = min(25, 22 + len(lines))
        if data.note.strip():
            style_source_row = max(22, note_row - 1)
            self._copy_cell_style(root, f"C{style_source_row}", f"C{note_row}")
            self._set_text(root, f"C{note_row}", data.note.strip())
        dimension = root.find(qn(NS_MAIN, "dimension"))
        if dimension is not None:
            dimension.set("ref", "A1:M31")

    # ------------------------------------------------------------------
    # 구매물품내역서
    # ------------------------------------------------------------------
    def _write_statement(self, root: ET.Element, data: ProjectData) -> None:
        if not data.statement_items:
            data.sync_statement_from_quote()
        self._set_text(root, "A1", "구매물품내역서")
        self._set_text(root, "A2", f"현장명 : {data.site_name}")
        self._set_text(root, "I2", "별첨(1)")
        self._clear_range(root, 1, 9, 5, 55)

        row = 5
        current_group = ""
        for item in (item for item in data.statement_items if item.name.strip()):
            group_title = item.group_title.strip()
            if group_title and group_title != current_group:
                if row > 55:
                    break
                self._copy_row_styles(root, 5, row, 1, 9)
                self._set_text(root, f"B{row}", group_title)
                current_group = group_title
                row += 1
            if row > 55:
                break
            self._copy_row_styles(root, 6, row, 1, 9)
            self._set_text(root, f"A{row}", item.number)
            self._set_text(root, f"B{row}", item.name)
            self._set_text(root, f"C{row}", item.spec)
            self._set_text(root, f"D{row}", item.unit)
            self._set_number(root, f"E{row}", item.quantity)
            self._set_number(root, f"F{row}", item.unit_price)
            self._set_number(root, f"G{row}", item.amount)
            self._set_text(root, f"H{row}", "")  # 품목별 가실행 금액은 표시하지 않음
            self._set_text(root, f"I{row}", item.note)
            row += 1

        self._set_text(root, "B56", "공  급  가")
        self._set_text(root, "B57", "부  가  세")
        self._set_text(root, "B58", "합      계")
        self._set_number(root, "G56", data.statement_supply_total)
        self._set_number(root, "G57", data.statement_vat)
        self._set_number(root, "G58", data.statement_total)
        self._set_number(root, "H56", data.budget_supply)
        self._set_number(root, "H57", data.budget_vat)
        self._set_number(root, "H58", data.budget_amount)
        self._set_formula(
            root,
            "I58",
            'IF(H58=0,"",G58/H58)',
            data.statement_total / data.budget_amount if data.budget_amount else None,
            precision=False,
        )
        self._set_text(root, "A59", "YTF-자재-001(0)")
        self._set_text(root, "H59", "양우건설㈜")
        self._fit_statement_columns(root, data)
        dimension = root.find(qn(NS_MAIN, "dimension"))
        if dimension is not None:
            dimension.set("ref", "A1:I59")

    def _fit_statement_columns(self, root: ET.Element, data: ProjectData) -> None:
        """내역서 품명·규격·수량이 뭉개지지 않도록 인쇄 폭 안에서 자동 배분한다."""

        valid = [item for item in data.statement_items if item.name.strip()]
        longest_name = max([len(item.name) for item in valid] + [4])
        longest_spec = max([len(item.spec) for item in valid] + [4])
        quantity_len = max([len(decimal_text(item.quantity)) for item in valid] + [4])
        widths = {
            1: 5.5,
            2: min(20.0, max(11.0, 6.5 + longest_name * 0.75)),
            3: min(25.0, max(16.0, 8.0 + longest_spec * 0.82)),
            4: 5.5,
            5: min(14.0, max(9.0, quantity_len + 2.5)),
            6: 12.5,
            7: 15.0,
            8: 12.5,
            9: 8.0,
        }
        old_cols = root.find(qn(NS_MAIN, "cols"))
        source_attrs: dict[int, dict[str, str]] = {}
        if old_cols is not None:
            for index in range(1, 10):
                for node in old_cols.findall(qn(NS_MAIN, "col")):
                    if int(node.get("min", "1")) <= index <= int(node.get("max", "1")):
                        source_attrs[index] = dict(node.attrib)
                        break
            root.remove(old_cols)
        new_cols = ET.Element(qn(NS_MAIN, "cols"))
        for index in range(1, 10):
            attrs = source_attrs.get(index, {"style": "16"})
            attrs.update({
                "min": str(index),
                "max": str(index),
                "width": f"{widths[index]:.3f}",
                "customWidth": "1",
            })
            attrs.pop("bestFit", None)
            new_cols.append(ET.Element(qn(NS_MAIN, "col"), attrs))
        # 인쇄영역 밖은 기본 너비 하나로 정리해 불필요한 반복 열 정의도 제거한다.
        new_cols.append(ET.Element(qn(NS_MAIN, "col"), {
            "min": "10", "max": "16384", "width": "8.886", "style": "16"
        }))
        sheet_data = self._sheet_data(root)
        root.insert(list(root).index(sheet_data), new_cols)

    # ------------------------------------------------------------------
    # 견적대비표 - 업체 수에 맞춰 열 재구성
    # ------------------------------------------------------------------
    def _rebuild_quote_sheet(
        self,
        root: ET.Element,
        data: ProjectData,
    ) -> str:
        original = deepcopy(root)
        vendor_count = max(1, len(data.vendors))
        last_col_index = 5 + vendor_count * 2
        last_col = index_to_col(last_col_index)
        output_rows = data.quote_output_rows()[:17]
        row_kind_by_number = {
            9 + index: kind for index, (kind, _sequence, _item) in enumerate(output_rows)
        }

        original_cells = {
            cell.get("r", ""): deepcopy(cell)
            for cell in original.findall(f".//{qn(NS_MAIN, 'c')}")
            if cell.get("r")
        }
        original_rows = {
            int(row.get("r", "0")): deepcopy(row)
            for row in self._sheet_data(original).findall(qn(NS_MAIN, "row"))
        }

        new_sheet_data = ET.Element(qn(NS_MAIN, "sheetData"))
        old_sheet_data = self._sheet_data(root)
        parent_index = list(root).index(old_sheet_data)
        root.remove(old_sheet_data)
        root.insert(parent_index, new_sheet_data)

        def source_cell(
            ref: str,
            fallback: str | None = None,
            *,
            prefer_fallback: bool = False,
        ) -> ET.Element:
            source = original_cells.get(fallback) if prefer_fallback and fallback else original_cells.get(ref)
            if source is None and fallback:
                source = original_cells.get(fallback)
            if source is None:
                source = ET.Element(qn(NS_MAIN, "c"), {"r": ref})
            clone = deepcopy(source)
            clone.set("r", ref)
            return clone

        def add_row(row_number: int) -> ET.Element:
            source_row_number = row_number
            if 9 <= row_number <= 25:
                source_row_number = 9 if row_kind_by_number.get(row_number) == "group" else 10
            source = original_rows.get(source_row_number)
            if source is None:
                source = original_rows.get(10)
            attrs = {"r": str(row_number)}
            if source is not None:
                attrs.update({key: value for key, value in source.attrib.items() if key != "spans"})
                attrs["r"] = str(row_number)
            attrs["spans"] = f"1:{last_col_index}"
            row = ET.Element(qn(NS_MAIN, "row"), attrs)
            new_sheet_data.append(row)
            return row

        def append_cell(
            row: ET.Element,
            ref: str,
            fallback: str | None = None,
            keep_value: bool = False,
            *,
            prefer_fallback: bool = False,
        ) -> None:
            cell = source_cell(ref, fallback, prefer_fallback=prefer_fallback)
            if not keep_value:
                self._clear_cell(cell)
            row.append(cell)

        for row_number in range(1, 34):
            row = add_row(row_number)
            source_item_row = 9 if row_kind_by_number.get(row_number) == "group" else 10
            for col_index in range(1, 6):
                ref = f"{index_to_col(col_index)}{row_number}"
                fallback = ref
                if 9 <= row_number <= 25:
                    fallback = f"{index_to_col(col_index)}{source_item_row}"
                keep = row_number in {1, 4, 5, 6, 7, 8, 27, 28, 29, 30, 31, 32, 33}
                append_cell(row, ref, fallback, keep_value=keep)
            for vendor_index in range(vendor_count):
                delta = vendor_index * 2
                for template_col in (6, 7):
                    target_col = template_col + delta
                    target_ref = f"{index_to_col(target_col)}{row_number}"
                    template_ref = f"{index_to_col(template_col)}{row_number}"
                    if 9 <= row_number <= 25:
                        template_ref = f"{index_to_col(template_col)}{source_item_row}"
                    keep = row_number == 8
                    append_cell(
                        row,
                        target_ref,
                        template_ref,
                        keep_value=keep,
                        prefer_fallback=vendor_index > 0,
                    )

        self._rebuild_quote_columns(root, original, data)
        self._rebuild_quote_merges(root, vendor_count, last_col)

        self._set_text(root, "A1", f"견적대비표({data.item_label})")
        self._set_text(root, "A2", f"현장명 : {data.site_name}")
        self._set_text(root, "A3", f"품  명 : {data.quote_title or data.purchase_title}")
        self._set_text(root, "E3", "")

        label_col = index_to_col(max(1, last_col_index - 2))
        value_col = index_to_col(max(1, last_col_index - 1))
        attachment_col = last_col
        self._set_text(root, f"{label_col}2", "작성자 :")
        self._set_text(root, f"{value_col}2", data.author or data.drafter)
        self._set_text(root, f"{label_col}3", "작성일 :")
        self._set_text(root, f"{value_col}3", data.quote_date_text)
        self._set_text(root, f"{attachment_col}3", "별첨(2)")
        # 원본 양식에서 검증된 정렬 스타일 ID만 복사한다. styles.xml 자체는
        # 수정하지 않으므로 Excel의 XML 복구 경고가 발생하지 않는다.
        self._apply_style_from_template_cell(root, f"{label_col}2", original_cells.get("G2"))
        self._apply_style_from_template_cell(root, f"{label_col}3", original_cells.get("G3"))
        self._apply_style_from_template_cell(root, f"{value_col}2", original_cells.get("J2"))
        self._apply_style_from_template_cell(root, f"{value_col}3", original_cells.get("J3"))
        self._apply_style_from_template_cell(root, f"{attachment_col}3", original_cells.get("K3"))

        for vendor_index, vendor in enumerate(data.vendors):
            unit_col = 6 + vendor_index * 2
            amount_col = unit_col + 1
            unit_letter = index_to_col(unit_col)
            amount_letter = index_to_col(amount_col)
            self._set_text(root, f"{unit_letter}4", vendor.name)
            self._set_text(root, f"{unit_letter}5", vendor.phone)
            self._set_text(root, f"{unit_letter}6", vendor.manager)
            self._set_number(root, f"{unit_letter}7", data.vendor_supply_total(vendor_index))
            self._set_text(root, f"{unit_letter}8", "단가")
            self._set_text(root, f"{amount_letter}8", "금액")

        for output_index, (kind, sequence, item) in enumerate(output_rows):
            row = 9 + output_index
            if item is None:
                continue
            if kind == "group":
                self._set_text(root, f"A{row}", sequence)
                self._set_text(root, f"B{row}", item.group_title)
                continue
            self._set_text(root, f"A{row}", "")
            self._set_text(root, f"B{row}", item.name)
            self._set_text(root, f"C{row}", item.spec)
            self._set_text(root, f"D{row}", item.unit)
            self._set_number(root, f"E{row}", item.quantity)
            for vendor_index in range(len(data.vendors)):
                unit_col = 6 + vendor_index * 2
                amount_col = unit_col + 1
                unit_letter = index_to_col(unit_col)
                amount_letter = index_to_col(amount_col)
                unit_price = item.unit_prices[vendor_index]
                amount = item.amount_for(vendor_index)
                self._set_number(root, f"{unit_letter}{row}", unit_price)
                self._set_formula(
                    root,
                    f"{amount_letter}{row}",
                    f"E{row}*{unit_letter}{row}",
                    amount,
                )

        self._set_text(root, "A27", "공 급 가")
        self._set_text(root, "A28", "부 가 세")
        self._set_text(root, "A29", "합    계")
        self._set_text(root, "A30", "결 재   조 건")
        self._set_text(root, "A31", "납 품   장 소")
        self._set_text(root, "A32", "납 부(설치)일")
        for vendor_index, vendor in enumerate(data.vendors):
            unit_col = 6 + vendor_index * 2
            amount_col = unit_col + 1
            unit_letter = index_to_col(unit_col)
            amount_letter = index_to_col(amount_col)
            self._set_number(root, f"{amount_letter}27", data.vendor_supply_total(vendor_index))
            self._set_number(root, f"{amount_letter}28", data.vendor_vat(vendor_index))
            self._set_number(root, f"{amount_letter}29", data.vendor_total(vendor_index))
            self._set_text(root, f"{unit_letter}30", vendor.payment)
            self._set_text(root, f"{unit_letter}31", vendor.delivery_place)
            self._set_text(root, f"{unit_letter}32", vendor.delivery_date)
        self._set_text(root, "A33", "양우F-32  (0)")
        self._set_text(root, f"{last_col}33", "양우건설주식회사")

        dimension = root.find(qn(NS_MAIN, "dimension"))
        if dimension is not None:
            dimension.set("ref", f"A1:{last_col}33")
        self._set_quote_page_setup(root)
        return last_col

    @staticmethod
    def _display_width(value: object) -> float:
        """Excel 열 너비 계산용 시각 폭. 한글/전각 문자는 두 칸으로 본다."""

        text = "" if value is None else str(value)
        width = 0.0
        for char in text:
            width += 2.0 if unicodedata.east_asian_width(char) in {"W", "F", "A"} else 1.0
        return width

    def _rebuild_quote_columns(self, root: ET.Element, original: ET.Element, data: ProjectData) -> None:
        """견적대비표의 모든 표시 열을 실제 데이터 길이에 맞춰 자동 조정한다."""

        vendor_count = max(1, len(data.vendors))
        old_cols = root.find(qn(NS_MAIN, "cols"))
        if old_cols is not None:
            root.remove(old_cols)
        new_cols = ET.Element(qn(NS_MAIN, "cols"))
        sheet_data = self._sheet_data(root)
        root.insert(list(root).index(sheet_data), new_cols)

        def attrs_for(column_index: int) -> dict[str, str]:
            cols = original.find(qn(NS_MAIN, "cols"))
            if cols is not None:
                for node in cols.findall(qn(NS_MAIN, "col")):
                    if int(node.get("min", "1")) <= column_index <= int(node.get("max", "1")):
                        attrs = dict(node.attrib)
                        attrs["min"] = str(column_index)
                        attrs["max"] = str(column_index)
                        return attrs
            return {"min": str(column_index), "max": str(column_index), "width": "11", "customWidth": "1"}

        output_rows = data.quote_output_rows()[:17]
        normal_items = [item for kind, _seq, item in output_rows if kind != "group" and item is not None]
        all_items = [item for _kind, _seq, item in output_rows if item is not None]

        sequence_width = max([self._display_width(seq) for _kind, seq, _item in output_rows] + [4.0])
        name_width = max(
            [self._display_width(item.name) for item in normal_items]
            + [self._display_width(item.group_title) for item in all_items]
            + [self._display_width("품명")]
        )
        spec_width = max([self._display_width(item.spec) for item in normal_items] + [self._display_width("규격")])
        unit_width = max([self._display_width(item.unit) for item in normal_items] + [self._display_width("단위")])
        quantity_width = max(
            [self._display_width(decimal_text(item.quantity)) for item in normal_items]
            + [self._display_width("수량")]
        )

        fixed_widths = {
            1: min(10.0, max(5.5, sequence_width + 2.0)),
            2: min(42.0, max(16.0, name_width * 0.92 + 3.5)),
            3: min(34.0, max(14.0, spec_width * 0.92 + 3.5)),
            4: min(12.0, max(6.0, unit_width + 2.5)),
            5: min(18.0, max(9.0, quantity_width + 3.0)),
        }

        for index in range(1, 6):
            attrs = attrs_for(index)
            attrs.update({
                "width": f"{fixed_widths[index]:.3f}",
                "customWidth": "1",
                "bestFit": "1",
            })
            new_cols.append(ET.Element(qn(NS_MAIN, "col"), attrs))

        for vendor_index in range(vendor_count):
            vendor = data.vendors[vendor_index] if vendor_index < len(data.vendors) else None
            unit_values = []
            amount_values = []
            for item in normal_items:
                price = item.unit_prices[vendor_index] if vendor_index < len(item.unit_prices) else 0
                unit_values.append(format_money(price))
                amount_values.append(format_money(item.amount_for(vendor_index)))
            if vendor_index < len(data.vendors):
                amount_values.extend([
                    format_money(data.vendor_supply_total(vendor_index)),
                    format_money(data.vendor_vat(vendor_index)),
                    format_money(data.vendor_total(vendor_index)),
                ])
            unit_width = min(19.0, max(11.5, max([self._display_width(v) for v in unit_values] + [4.0]) + 2.5))
            amount_width = min(23.0, max(13.5, max([self._display_width(v) for v in amount_values] + [4.0]) + 2.5))

            # 업체명은 단가·금액 두 열에 병합되어 표시된다. 두 열의 합이 업체명
            # 전체 폭보다 좁으면 금액 열 쪽에 여유를 더해 잘림을 방지한다.
            vendor_name_width = self._display_width(vendor.name if vendor else f"업체 {vendor_index + 1}") + 4.0
            shortage = max(0.0, vendor_name_width - (unit_width + amount_width))
            unit_width += shortage * 0.4
            amount_width += shortage * 0.6
            unit_width = min(unit_width, 23.0)
            amount_width = min(amount_width, 29.0)

            for offset, width in enumerate((unit_width, amount_width)):
                target = 6 + vendor_index * 2 + offset
                attrs = attrs_for(6 + offset)
                attrs.update({
                    "min": str(target),
                    "max": str(target),
                    "width": f"{width:.3f}",
                    "customWidth": "1",
                    "bestFit": "1",
                })
                new_cols.append(ET.Element(qn(NS_MAIN, "col"), attrs))

    def _rebuild_quote_merges(self, root: ET.Element, vendor_count: int, last_col: str) -> None:
        old = root.find(qn(NS_MAIN, "mergeCells"))
        if old is not None:
            root.remove(old)
        merges = [
            f"A1:{last_col}1",
            "A3:C3",
            "A4:E4",
            "A5:E5",
            "A6:E6",
            "A7:E7",
            "A27:B27",
            "A28:B28",
            "A29:B29",
            "A30:E30",
            "A31:E31",
            "A32:E32",
        ]
        for vendor_index in range(vendor_count):
            start = index_to_col(6 + vendor_index * 2)
            end = index_to_col(7 + vendor_index * 2)
            for row in (4, 5, 6, 7, 30, 31, 32):
                merges.append(f"{start}{row}:{end}{row}")
        merge_node = ET.Element(qn(NS_MAIN, "mergeCells"), {"count": str(len(merges))})
        for ref in merges:
            merge_node.append(ET.Element(qn(NS_MAIN, "mergeCell"), {"ref": ref}))
        # sheetData 다음에 배치
        sheet_data = self._sheet_data(root)
        root.insert(list(root).index(sheet_data) + 1, merge_node)

    def _set_quote_page_setup(self, root: ET.Element) -> None:
        page_setup = root.find(qn(NS_MAIN, "pageSetup"))
        if page_setup is None:
            page_setup = ET.SubElement(root, qn(NS_MAIN, "pageSetup"))
        page_setup.set("paperSize", "9")
        page_setup.set("orientation", "landscape")
        page_setup.set("fitToWidth", "1")
        page_setup.set("fitToHeight", "1")
        page_setup.attrib.pop("scale", None)
        sheet_pr = root.find(qn(NS_MAIN, "sheetPr"))
        if sheet_pr is None:
            sheet_pr = ET.Element(qn(NS_MAIN, "sheetPr"))
            root.insert(0, sheet_pr)
        page_set_up_pr = sheet_pr.find(qn(NS_MAIN, "pageSetUpPr"))
        if page_set_up_pr is None:
            page_set_up_pr = ET.SubElement(sheet_pr, qn(NS_MAIN, "pageSetUpPr"))
        page_set_up_pr.set("fitToPage", "1")

    # ------------------------------------------------------------------
    # Package metadata / workbook names
    # ------------------------------------------------------------------
    def _clean_workbook(self, xml_data: bytes, data: ProjectData, quote_last_col: str) -> bytes:
        root = ET.fromstring(xml_data)
        sheets = root.find(qn(NS_MAIN, "sheets"))
        if sheets is None or len(sheets) < 3:
            raise ValueError("3개 시트 양식이 아닙니다.")
        sheets[0].set("name", "구매품의서")
        sheets[1].set("name", data.statement_sheet_name)
        sheets[2].set("name", data.quote_sheet_name)

        external = root.find(qn(NS_MAIN, "externalReferences"))
        if external is not None:
            root.remove(external)
        for alternate in list(root.findall(qn(NS_MC, "AlternateContent"))):
            root.remove(alternate)

        # 공동편집 revisionPtr, xr 계열 uid, 미사용 mc:Ignorable 및 계산 확장을
        # 제거해 Excel 데스크톱에서 복구 대화상자가 뜨지 않게 한다.
        ext_list = root.find(qn(NS_MAIN, "extLst"))
        if ext_list is not None:
            root.remove(ext_list)
        self._strip_extension_metadata(root)

        defined_names = root.find(qn(NS_MAIN, "definedNames"))
        if defined_names is not None:
            root.remove(defined_names)
        defined_names = ET.Element(qn(NS_MAIN, "definedNames"))
        sheets_index = list(root).index(sheets)
        root.insert(sheets_index + 1, defined_names)
        print_areas = (
            (0, "구매품의서", "$A$2:$M$31"),
            (1, data.statement_sheet_name, "$A$1:$I$59"),
            (2, data.quote_sheet_name, f"$A$1:${quote_last_col}$33"),
        )
        for local_id, sheet_name, area in print_areas:
            node = ET.SubElement(
                defined_names,
                qn(NS_MAIN, "definedName"),
                {"name": "_xlnm.Print_Area", "localSheetId": str(local_id)},
            )
            escaped = sheet_name.replace("'", "''")
            node.text = f"'{escaped}'!{area}"

        book_views = root.find(qn(NS_MAIN, "bookViews"))
        if book_views is not None:
            for view in book_views.findall(qn(NS_MAIN, "workbookView")):
                view.set("activeTab", "0")
        calc_pr = root.find(qn(NS_MAIN, "calcPr"))
        if calc_pr is None:
            calc_pr = ET.SubElement(root, qn(NS_MAIN, "calcPr"))
        calc_pr.set("calcMode", "auto")
        calc_pr.set("fullCalcOnLoad", "1")
        calc_pr.set("forceFullCalc", "1")
        return xml_bytes(root)

    def _clean_workbook_rels(self, xml_data: bytes) -> bytes:
        root = ET.fromstring(xml_data)
        for node in list(root):
            rel_type = node.get("Type", "")
            if rel_type.endswith("/externalLink") or rel_type.endswith("/calcChain"):
                root.remove(node)
        return xml_bytes_default(root, NS_PACKAGE_REL)

    def _clean_content_types(self, xml_data: bytes) -> bytes:
        root = ET.fromstring(xml_data)
        for node in list(root):
            part = node.get("PartName", "")
            if part.startswith("/xl/externalLinks/") or part == "/xl/calcChain.xml":
                root.remove(node)
        return xml_bytes_default(root, NS_CT)

    def _clean_core_properties(self, xml_data: bytes) -> bytes:
        root = ET.fromstring(xml_data)
        creator = root.find(qn(NS_DC, "creator"))
        if creator is None:
            creator = ET.SubElement(root, qn(NS_DC, "creator"))
        creator.text = "자재구매품의서 자동작성"
        modified_by = root.find(qn(NS_CP, "lastModifiedBy"))
        if modified_by is None:
            modified_by = ET.SubElement(root, qn(NS_CP, "lastModifiedBy"))
        modified_by.text = "자재구매품의서 자동작성"
        revision = root.find(qn(NS_CP, "revision"))
        if revision is not None:
            revision.text = "1"
        now = datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")
        for tag in ("created", "modified"):
            node = root.find(qn(NS_DCTERMS, tag))
            if node is not None:
                node.text = now
        return xml_bytes(root)
