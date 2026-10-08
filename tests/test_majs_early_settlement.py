from datetime import date
from decimal import Decimal

from core.indices import IndexSeries
from core.majs import (
    MAJSEarlySettlement,
    MAJSSettings,
    SEM_CORRECAO,
    calculate_majs,
)


def _constant_series(start_year=2019, end_year=2026, rate=Decimal("0.003")):
    values = {
        f"{year:04d}-{month:02d}": rate
        for year in range(start_year, end_year + 1)
        for month in range(1, 13)
    }
    return IndexSeries(code="INPC", values=values, source="teste")


def _settings():
    return MAJSSettings(
        original_amount=Decimal("80000"),
        credit_date=date(2020, 1, 8),
        monthly_interest_rate=Decimal("0.006"),
        term_months=96,
        base_date=date(2026, 8, 15),
        correction_index_code="INPC",
        correction_lag_months=2,
    )


def test_early_settlement_novation_stops_flow_without_settlement_difference():
    settlement = MAJSEarlySettlement(date(2022, 5, 10), novation=True)
    result = calculate_majs(
        _settings(),
        {"INPC": _constant_series()},
        early_settlement=settlement,
        difference_correction_mode=SEM_CORRECAO,
        difference_interest_mode="SEM_JUROS",
    )

    assert result.was_settled_early
    assert result.early_settlement_balance == Decimal("58737.45")
    assert result.balance_to_mature == Decimal("0")
    assert result.last_evolution_date == date(2022, 5, 10)
    assert not any(r.due_date > settlement.settlement_date for r in result.schedule)
    assert not any(r.origin == "Quitação antecipada" for r in result.differences)


def test_early_settlement_payment_creates_difference_against_theoretical_balance():
    settlement = MAJSEarlySettlement(
        date(2022, 5, 10),
        novation=False,
        amount_paid=Decimal("50000"),
        note="Quitação antecipada por acordo",
    )
    result = calculate_majs(
        _settings(),
        {"INPC": _constant_series()},
        early_settlement=settlement,
        difference_correction_mode=SEM_CORRECAO,
        difference_interest_mode="SEM_JUROS",
    )

    rows = [r for r in result.differences if r.origin == "Quitação antecipada"]
    assert len(rows) == 1
    row = rows[0]
    assert result.early_settlement_balance == Decimal("58737.45")
    assert row.amount_due == Decimal("58737.45")
    assert row.amount_paid == Decimal("50000.00")
    assert row.difference_paid_minus_due == Decimal("-8737.45")
    assert result.balance_to_mature == Decimal("0")
    assert result.last_evolution_date == date(2022, 5, 10)
