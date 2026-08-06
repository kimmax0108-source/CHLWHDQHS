from decimal import Decimal

from expense_statement_app.calculations import financial_amount_text, korean_integer, quantity_parts


def test_korean_amount_text() -> None:
    assert korean_integer(1_070_849_527) == "일십억칠천팔십사만구천오백이십칠"
    assert financial_amount_text(1_070_849_527) == "一金 일십억칠천팔십사만구천오백이십칠원整"


def test_quantity_parts() -> None:
    assert quantity_parts(Decimal("1178.660")) == ("1178.", "660")
