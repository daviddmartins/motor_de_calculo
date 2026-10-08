from datetime import date
from decimal import Decimal

from core.indices import IndexSeries
from core.monetary_update import (
    CORRECTION_PRORATA_CALENDAR,
    INTEREST_APPLICATION_END,
    INTEREST_BASE_CORRECTED,
    INTEREST_SIMPLE_YEARFRAC,
    UpdateItem,
    UpdateRule,
    calculate_monetary_update,
)


def test_missing_future_index_keeps_interest_until_base_date():
    series = IndexSeries(code="INPC", values={"2021-03": Decimal("0.0086")}, source="teste")
    result = calculate_monetary_update(
        items=[UpdateItem(1, date(2021, 3, 8), Decimal("10873.82"), "teste")],
        rules=[UpdateRule(
            order=1,
            start_date=date(2021, 3, 8),
            end_date=date(2021, 4, 30),
            correction_index_code="INPC",
            correction_method=CORRECTION_PRORATA_CALENDAR,
            monthly_interest_rate=Decimal("0.01"),
            interest_method=INTEREST_SIMPLE_YEARFRAC,
            interest_application=INTEREST_APPLICATION_END,
            interest_start_date=date(2021, 3, 8),
            interest_base=INTEREST_BASE_CORRECTED,
            competence_start_day=1,
        )],
        base_date=date(2021, 4, 30),
        indices={"INPC": series},
        holidays=set(),
        round_daily_money=False,
    )
    assert result.is_complete is False
    assert result.missing_index_reference == "2021-04"
    assert result.daily_rows[-1].day == date(2021, 4, 30)
    assert result.interest_total > 0
