from __future__ import annotations

from decimal import Decimal

KOREAN_DIGITS = ("영", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구")
SMALL_UNITS = ("", "십", "백", "천")
LARGE_UNITS = ("", "만", "억", "조", "경")


def _group_to_korean(group: int) -> str:
    result: list[str] = []
    for position in range(3, -1, -1):
        divisor = 10**position
        digit = group // divisor
        group %= divisor
        if digit:
            result.append(KOREAN_DIGITS[digit])
            result.append(SMALL_UNITS[position])
    return "".join(result)


def korean_integer(value: int) -> str:
    if value == 0:
        return "영"
    if value < 0:
        return "마이너스" + korean_integer(abs(value))
    groups: list[int] = []
    while value:
        groups.append(value % 10_000)
        value //= 10_000
    parts: list[str] = []
    for idx in range(len(groups) - 1, -1, -1):
        group = groups[idx]
        if not group:
            continue
        parts.append(_group_to_korean(group))
        parts.append(LARGE_UNITS[idx])
    return "".join(parts)


def financial_amount_text(value: int) -> str:
    return f"一金 {korean_integer(int(value))}원整"


def quantity_parts(value: Decimal | str | float | int) -> tuple[str, str]:
    decimal_value = Decimal(str(value)).quantize(Decimal("0.001"))
    integer, fraction = f"{decimal_value:.3f}".split(".")
    return f"{integer}.", fraction
