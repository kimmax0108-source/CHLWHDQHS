from __future__ import annotations

from html import escape

from .models import ProjectData, decimal_text, format_money, purchase_payment_text, wrap_body_text


def _e(value: object) -> str:
    return escape("" if value is None else str(value))


def _money(value: object) -> str:
    return _e(format_money(value))


def _base_html(title: str, body: str, min_width: int = 720) -> str:
    return f"""
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<style>
  html, body {{ margin: 0; padding: 0; background: #dfe4ec; }}
  body {{ font-family: 'Malgun Gothic', '맑은 고딕', sans-serif; color: #111827; }}
  .canvas {{ padding: 14px; }}
  .paper {{ min-width: {min_width}px; margin: 0 auto; background: white; border: 1px solid #7b8492;
            box-shadow: 0 2px 10px rgba(0,0,0,.12); padding: 15px; }}
  .document-title {{ text-align: center; font-size: 25px; font-weight: 700; letter-spacing: 6px;
                     margin: 4px 0 12px; }}
  .sub-title {{ text-align: center; font-size: 21px; font-weight: 700; margin: 4px 0 10px; }}
  table {{ width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 11px; }}
  th, td {{ border: 1px solid #303642; padding: 5px 6px; vertical-align: middle; }}
  th {{ background: #edf2f8; font-weight: 700; text-align: center; }}
  .label {{ width: 92px; background: #f3f6fa; text-align: center; font-weight: 700; }}
  .center {{ text-align: center; }}
  .right {{ text-align: right; }}
  .left {{ text-align: left; }}
  .small {{ font-size: 10px; color: #4b5563; }}
  .body-line {{ height: 28px; font-size: 12px; }}
  .winner {{ background: #e8f5e9; font-weight: 700; }}
  .group {{ font-weight: 700; background: #f8fafc; }}
  .muted {{ color: #6b7280; }}
  .footer {{ display: flex; justify-content: space-between; font-size: 10px; margin-top: 8px; }}
  .meta {{ display: grid; grid-template-columns: 1fr auto auto; gap: 12px; align-items: center;
           font-size: 10px; margin-bottom: 5px; }}
  .meta .label-right {{ text-align: right; }}
</style>
<title>{_e(title)}</title>
</head>
<body><div class="canvas"><div class="paper">{body}</div></div></body>
</html>
"""


def purchase_preview_html(data: ProjectData) -> str:
    vendor = data.selected_vendor
    body_lines = wrap_body_text(data.body_text, width=42, max_lines=3)
    while len(body_lines) < 3:
        body_lines.append("")
    period_label = {
        "계약기간": "계 약 기 간",
        "임차기간": "임 차 기 간",
        "납품일자": "납 품 일 자",
    }.get(data.period_kind, data.period_kind)
    body = f"""
<div class="document-title">자재구매품의서</div>
<table>
  <tr><td class="label">분류번호</td><td colspan="5">{_e(data.classification)}</td><td class="label">기안부서</td><td colspan="3" class="center">{_e(data.department)}</td></tr>
  <tr><td class="label">기안일자</td><td colspan="3" class="center">{_e(data.draft_date_text)}</td><td class="label">시행일자</td><td colspan="5">{_e(data.effective_date)}</td></tr>
  <tr><td class="label">기안자</td><td colspan="3" class="center">{_e(data.drafter)}</td><td colspan="6" class="center small">{_e(data.approval_note)}</td></tr>
  <tr><td class="label">수신처</td><td colspan="4" class="center">공 의</td><td class="label">발신</td><td colspan="4" class="center">내 부 결 재</td></tr>
  <tr><td class="label">제목</td><td colspan="9">{_e(data.purchase_title)}</td></tr>
  <tr><td class="label">1. 현장명</td><td colspan="9">{_e(data.purchase_site_effective)}</td></tr>
  <tr><td class="label">2. 품명</td><td colspan="9">{_e(data.purchase_item_name)}</td></tr>
  <tr><td class="label">3. 가실행</td><td colspan="7" class="right">₩ {_money(data.purchase_budget_effective)}</td><td colspan="2" class="center">부가세 포함</td></tr>
  <tr><td class="label">4. 금액</td><td colspan="7" class="right">₩ {_money(data.purchase_contract_effective)}</td><td colspan="2" class="center winner">{_e(data.purchase_ratio_text)}</td></tr>
  <tr><td class="label">5. {_e(period_label)}</td><td colspan="9">{_e(data.period)}</td></tr>
  <tr><td class="label">6. 거래처</td><td colspan="5">{_e(data.purchase_vendor_effective)}</td><td class="label">전화</td><td colspan="3">{_e(data.purchase_phone_effective)}</td></tr>
  <tr><td class="label">7. 첨부</td><td colspan="9">{_e(data.attachment)}</td></tr>
  <tr><td class="label">8. 지불조건</td><td colspan="9">{_e(purchase_payment_text(data.payment))}</td></tr>
  <tr><td colspan="10" class="body-line">{_e(body_lines[0])}</td></tr>
  <tr><td colspan="10" class="body-line">{_e(body_lines[1])}</td></tr>
  <tr><td colspan="10" class="body-line">{_e(body_lines[2])}</td></tr>
  <tr><td colspan="10" class="body-line">{_e(data.note)}</td></tr>
</table>
<div class="footer"><span>양우F - 32 (0)</span><span>양우건설주식회사</span></div>
"""
    return _base_html("구매품의서 미리보기", body, min_width=760)


def statement_preview_html(data: ProjectData) -> str:
    """내역서 Excel 출력 구조와 같은 9열·합계 배치로 미리보기를 만든다."""

    vendor = data.selected_vendor
    if not data.statement_items and data.items:
        data.sync_statement_from_quote()
    rows: list[str] = []
    current_group = ""
    for item in (item for item in data.statement_items if item.name.strip()):
        if item.group_title.strip() and item.group_title.strip() != current_group:
            current_group = item.group_title.strip()
            rows.append(
                f"<tr class='group statement-group'><td></td><td colspan='8'>{_e(current_group)}</td></tr>"
            )
        rows.append(
            f"<tr class='statement-item'><td class='center'>{_e(item.number)}</td><td>{_e(item.name)}</td><td>{_e(item.spec)}</td>"
            f"<td class='center'>{_e(item.unit)}</td><td class='right'>{_e(decimal_text(item.quantity))}</td>"
            f"<td class='right'>{_money(item.unit_price)}</td><td class='right'>{_money(item.amount)}</td>"
            f"<td></td><td>{_e(item.note)}</td></tr>"
        )
    if not rows:
        rows.append("<tr><td colspan='9' class='center muted'>견적대비표에서 품목을 불러와 주세요.</td></tr>")
    ratio = data.statement_total / data.budget_amount if data.budget_amount else None
    ratio_text = "-" if ratio is None else f"{ratio:.1%}"
    body = f"""
<style>
  .statement-sheet {{ border: 2px solid #111827; }}
  .statement-sheet th {{ background: #ffffff; padding: 6px 4px; }}
  .statement-sheet td {{ padding: 6px 5px; line-height: 1.25; overflow-wrap: anywhere; }}
  .statement-title {{ border-bottom: 2px solid #111827; padding-bottom: 4px; }}
  .statement-meta {{ display:flex; justify-content:space-between; gap:16px; font-size:10px; margin:4px 0 8px; }}
  .statement-group td {{ background:#f4f7fb; font-weight:700; }}
  .statement-total-label {{ text-align:right; font-weight:700; letter-spacing:2px; }}
  .statement-budget {{ text-align:right; }}
  .statement-ratio {{ text-align:center; font-weight:700; }}
</style>
<div class="sub-title statement-title">구매물품내역서</div>
<div class="statement-meta"><span>현장명 : {_e(data.site_name)}</span><span>선정업체 : {_e(vendor.name if vendor else '선정업체 없음')}</span></div>
<table class="statement-sheet">
  <colgroup>
    <col style="width:6%"><col style="width:20%"><col style="width:17%"><col style="width:7%">
    <col style="width:9%"><col style="width:12%"><col style="width:14%"><col style="width:11%"><col style="width:4%">
  </colgroup>
  <tr><th>번호</th><th>품명</th><th>규격</th><th>단위</th><th>수량</th><th>단가</th><th>금액</th><th>가실행</th><th>비고</th></tr>
  {''.join(rows)}
  <tr><td colspan="6" class="statement-total-label">공 급 가</td><td class="right">{_money(data.statement_supply_total)}</td><td class="statement-budget">{_money(data.budget_supply)}</td><td></td></tr>
  <tr><td colspan="6" class="statement-total-label">부 가 세</td><td class="right">{_money(data.statement_vat)}</td><td class="statement-budget">{_money(data.budget_vat)}</td><td></td></tr>
  <tr class="winner"><td colspan="6" class="statement-total-label">합 계</td><td class="right">{_money(data.statement_total)}</td><td class="statement-budget">{_money(data.budget_amount)}</td><td class="statement-ratio">{_e(ratio_text)}</td></tr>
</table>
<div class="footer"><span>YTF-자재-001(0)</span><span>양우건설㈜</span></div>
"""
    return _base_html("내역서 미리보기", body, min_width=860)


def quote_preview_html(data: ProjectData) -> str:
    vendors = data.vendors
    vendor_count = max(1, len(vendors))
    vendor_headers: list[str] = []
    phone_rows: list[str] = []
    manager_rows: list[str] = []
    amount_rows: list[str] = []
    unit_headers: list[str] = []
    for index in range(vendor_count):
        vendor = vendors[index] if index < len(vendors) else None
        winner_class = " winner" if data.winner_index == index else ""
        vendor_headers.append(
            f"<th colspan='2' class='{winner_class.strip()}'>{_e(vendor.name if vendor else f'업체 {index + 1}')}</th>"
        )
        phone_rows.append(f"<td colspan='2' class='center small{winner_class}'>{_e(vendor.phone if vendor else '')}</td>")
        manager_rows.append(f"<td colspan='2' class='center small{winner_class}'>{_e(vendor.manager if vendor else '')}</td>")
        supply = data.vendor_supply_total(index) if index < len(vendors) else 0
        amount_rows.append(f"<td colspan='2' class='right{winner_class}'>{_money(supply)}</td>")
        unit_headers.append("<th>단가</th><th>금액</th>")

    item_rows: list[str] = []
    for kind, sequence, item in data.quote_output_rows():
        if item is None:
            continue
        if kind == "group":
            item_rows.append(
                f"<tr class='group'><td class='center'>{_e(sequence)}</td><td>{_e(item.group_title)}</td>"
                f"<td></td><td></td><td></td>{'<td></td>' * (vendor_count * 2)}</tr>"
            )
            continue
        price_cells: list[str] = []
        for vendor_index in range(vendor_count):
            price = item.unit_prices[vendor_index] if vendor_index < len(item.unit_prices) else 0
            amount = item.amount_for(vendor_index) if vendor_index < len(vendors) else 0
            css = " winner" if data.winner_index == vendor_index else ""
            price_cells.append(
                f"<td class='right{css}'>{_money(price)}</td><td class='right{css}'>{_money(amount)}</td>"
            )
        item_rows.append(
            f"<tr><td></td><td>{_e(item.name)}</td><td>{_e(item.spec)}</td>"
            f"<td class='center'>{_e(item.unit)}</td><td class='right'>{_e(decimal_text(item.quantity))}</td>{''.join(price_cells)}</tr>"
        )
    if not item_rows:
        item_rows.append(
            f"<tr><td colspan='{5 + vendor_count * 2}' class='center muted'>품목과 업체별 단가를 입력해 주세요.</td></tr>"
        )

    total_rows: list[str] = []
    for label, getter in (
        ("공급가", data.vendor_supply_total),
        ("부가세", data.vendor_vat),
        ("합계", data.vendor_total),
    ):
        cells = []
        for vendor_index in range(vendor_count):
            value = getter(vendor_index) if vendor_index < len(vendors) else 0
            css = " winner" if data.winner_index == vendor_index else ""
            cells.append(f"<td colspan='2' class='right{css}'>{_money(value)}</td>")
        total_rows.append(f"<tr><td colspan='5' class='right'><b>{_e(label)}</b></td>{''.join(cells)}</tr>")

    detail_rows: list[str] = []
    for label, attr in (("결제조건", "payment"), ("납품장소", "delivery_place"), ("납품/설치일", "delivery_date")):
        cells = []
        for vendor_index in range(vendor_count):
            vendor = vendors[vendor_index] if vendor_index < len(vendors) else None
            value = getattr(vendor, attr) if vendor else ""
            css = " winner" if data.winner_index == vendor_index else ""
            cells.append(f"<td colspan='2' class='center{css}'>{_e(value)}</td>")
        detail_rows.append(f"<tr><td colspan='5' class='right'><b>{_e(label)}</b></td>{''.join(cells)}</tr>")

    body = f"""
<div class="sub-title">견적대비표({_e(data.item_label)})</div>
<div class="meta"><span>현장명 : {_e(data.site_name)}</span><span class="label-right">작성자 :</span><span class="value-left">{_e(data.author or data.drafter)}</span><span></span></div>
<div class="meta"><span>품명 : {_e(data.quote_title or data.purchase_title)}</span><span class="label-right">작성일 :</span><span class="value-left">{_e(data.quote_date_text)}</span><span class="attachment-right">별첨(2)</span></div>
<table>
  <tr><th colspan="5">업체명</th>{''.join(vendor_headers)}</tr>
  <tr><td colspan="5" class="center">연락처</td>{''.join(phone_rows)}</tr>
  <tr><td colspan="5" class="center">담당</td>{''.join(manager_rows)}</tr>
  <tr><td colspan="5" class="center">견적금액</td>{''.join(amount_rows)}</tr>
  <tr><th style="width:42px">순서</th><th style="width:150px">품명</th><th style="width:105px">규격</th><th style="width:48px">단위</th><th style="width:52px">수량</th>{''.join(unit_headers)}</tr>
  {''.join(item_rows)}
  {''.join(total_rows)}
  {''.join(detail_rows)}
</table>
<div class="small" style="margin-top:8px">초록색 열은 현재 최종 선정업체입니다. 납품장소는 공통값 자동 적용 후 업체별 수정이 가능합니다.</div>
<div class="footer"><span>양우F-32 (0)</span><span>양우건설주식회사</span></div>
"""
    return _base_html("견적대비표 미리보기", body, min_width=max(820, 390 + vendor_count * 170))


def build_preview_html(data: ProjectData, kind: str) -> str:
    if kind == "statement":
        return statement_preview_html(data)
    if kind == "quote":
        return quote_preview_html(data)
    return purchase_preview_html(data)
