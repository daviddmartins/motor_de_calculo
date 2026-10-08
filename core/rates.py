from __future__ import annotations

from decimal import Decimal, localcontext


def _validate(rate: Decimal, label: str) -> None:
    if rate <= Decimal("-1"):
        raise ValueError(f"{label} deve ser superior a -100%.")


def annual_to_monthly_equivalent(annual_rate: Decimal) -> Decimal:
    _validate(annual_rate, "Taxa anual")
    with localcontext() as ctx:
        ctx.prec = 40
        return (Decimal("1") + annual_rate) ** (Decimal("1") / Decimal("12")) - Decimal("1")


def monthly_to_daily_equivalent(monthly_rate: Decimal, number_of_days: int) -> Decimal:
    _validate(monthly_rate, "Taxa mensal")
    if number_of_days <= 0:
        raise ValueError("O divisor diário deve ser positivo.")
    with localcontext() as ctx:
        ctx.prec = 40
        return (Decimal("1") + monthly_rate) ** (Decimal("1") / Decimal(number_of_days)) - Decimal("1")


def monthly_index_to_daily_equivalent(monthly_index: Decimal, business_days: int) -> Decimal:
    _validate(monthly_index, "Índice mensal")
    if business_days <= 0:
        raise ValueError("A competência precisa ter ao menos um dia útil.")
    with localcontext() as ctx:
        ctx.prec = 40
        return (Decimal("1") + monthly_index) ** (Decimal("1") / Decimal(business_days)) - Decimal("1")


def price_payment(present_value: Decimal, monthly_rate: Decimal, term: int, timing: int = 0) -> Decimal:
    if present_value <= 0 or term <= 0:
        raise ValueError("Valor presente e prazo devem ser positivos.")
    _validate(monthly_rate, "Taxa mensal")
    if timing not in (0, 1):
        raise ValueError("O tipo deve ser 0 ou 1.")
    if monthly_rate == 0:
        return present_value / Decimal(term)
    with localcontext() as ctx:
        ctx.prec = 40
        one = Decimal("1")
        factor = (one + monthly_rate) ** Decimal(term)
        payment = present_value * monthly_rate * factor / (factor - one)
        if timing == 1:
            payment = payment / (one + monthly_rate)
        return payment
