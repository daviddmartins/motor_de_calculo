from datetime import date
from decimal import Decimal, localcontext

from core.indices import IndexSeries
from core.monetary_update import (
    CORRECTION_PRORATA_CALENDAR,
    INTEREST_NONE,
    UpdateItem,
    UpdateRule,
    calculate_monetary_update,
)


def test_prorata_matches_reference_spreadsheet_first_partial_month():
    """08/03/2021 a 31/03/2021: 24/31 de uma cotação mensal de 0,86%."""
    series = IndexSeries(code="INPC", values={"2021-03": Decimal("0.0086")}, source="referencia")
    original = Decimal("10873.82")
    result = calculate_monetary_update(
        items=[UpdateItem(1, date(2021, 3, 8), original, "referencia")],
        rules=[UpdateRule(
            order=1,
            start_date=date(2021, 3, 8),
            end_date=date(2021, 3, 31),
            correction_index_code="INPC",
            correction_method=CORRECTION_PRORATA_CALENDAR,
            interest_method=INTEREST_NONE,
            competence_start_day=1,
        )],
        base_date=date(2021, 3, 31),
        indices={"INPC": series},
        holidays=set(),
        round_daily_money=False,
        money_places=8,
    )
    with localcontext() as ctx:
        ctx.prec = 40
        expected_factor = Decimal("1.0086") ** (Decimal(24) / Decimal(31))
        expected = original * expected_factor
    assert abs(result.updated_total - expected) < Decimal("0.00000001")
    assert len(result.daily_rows) == 24
    assert result.daily_rows[0].day == date(2021, 3, 8)
    assert result.daily_rows[-1].day == date(2021, 3, 31)
