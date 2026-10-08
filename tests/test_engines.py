from datetime import date
from decimal import Decimal
import unittest

from core.calendar_rules import first_reference_due_date, next_business_day
from core.engines import calculate_contract_evolution
from core.indices import IndexSeries
from core.models import ContractSettings, ExtraordinaryAmortization
from core.rates import annual_to_monthly_equivalent, monthly_to_daily_equivalent, price_payment


class CalendarTests(unittest.TestCase):
    def test_first_due_is_next_month(self):
        self.assertEqual(first_reference_due_date(date(2023, 8, 22)), date(2023, 9, 20))
        self.assertEqual(first_reference_due_date(date(2023, 8, 5)), date(2023, 9, 20))
        self.assertEqual(first_reference_due_date(date(2023, 8, 20)), date(2023, 9, 20))

    def test_operational_due_moves_but_reference_does_not(self):
        self.assertEqual(next_business_day(date(2026, 6, 20), set()), date(2026, 6, 22))


class RateTests(unittest.TestCase):
    def test_equivalent_daily_recomposes(self):
        monthly = annual_to_monthly_equivalent(Decimal("0.1353"))
        daily = monthly_to_daily_equivalent(monthly, 31)
        recomposed = (Decimal("1") + daily) ** Decimal("31") - Decimal("1")
        self.assertAlmostEqual(float(monthly), float(recomposed), places=12)

    def test_price_type_is_available_and_changes_payment(self):
        monthly = annual_to_monthly_equivalent(Decimal("0.1353"))
        end = price_payment(Decimal("10000"), monthly, 48, 0)
        beginning = price_payment(Decimal("10000"), monthly, 48, 1)
        self.assertLess(beginning, end)


class EngineTests(unittest.TestCase):
    @staticmethod
    def zero_indices():
        return IndexSeries(
            "INPC",
            {f"{y:04d}-{m:02d}": Decimal("0") for y in range(2010, 2040) for m in range(1, 13)},
        )

    def test_fixed_closes_balance(self):
        settings = ContractSettings(
            modality_code="MOD_002", initial_balance=Decimal("10000"),
            credit_date=date(2023, 8, 22), annual_interest_rate=Decimal("0.12"),
            term=24, competence_start_day=21, index_code=None, payment_timing=0,
        )
        result = calculate_contract_evolution(settings, None, set(), [])
        self.assertLess(abs(result.installments[-1].closing_balance), Decimal("0.02"))

    def test_variable_with_zero_index_closes_balance(self):
        settings = ContractSettings(
            modality_code="MOD_001", initial_balance=Decimal("12000"),
            credit_date=date(2023, 8, 22), annual_interest_rate=Decimal("0.08"),
            term=12, competence_start_day=21, index_code="INPC", index_lag_months=2,
        )
        result = calculate_contract_evolution(settings, self.zero_indices(), set(), [])
        self.assertLess(abs(result.installments[-1].closing_balance), Decimal("0.02"))

    def test_variable_amortization_uses_balance_less_period_interest(self):
        settings = ContractSettings(
            modality_code="MOD_001", initial_balance=Decimal("95000"),
            credit_date=date(2017, 11, 1), annual_interest_rate=Decimal("0.0769"),
            term=96, competence_start_day=21, index_code="INPC", index_lag_months=2,
        )
        values = self.zero_indices().values.copy()
        values["2017-08"] = Decimal("-0.0003")
        values["2017-09"] = Decimal("-0.0002")
        result = calculate_contract_evolution(settings, IndexSeries("INPC", values), set(), [])
        first = result.installments[0]
        expected_base = first.balance_before_installment - first.interest_amount
        self.assertLess(abs(first.amortization_base - expected_base), Decimal("0.02"))
        expected_amortization = expected_base / Decimal("96")
        self.assertLess(abs(first.regular_amortization - expected_amortization), Decimal("0.02"))
        self.assertLess(
            abs(first.installment_amount - (first.interest_amount + first.regular_amortization)),
            Decimal("0.02"),
        )

    def test_correction_is_applied_to_balance_after_interest(self):
        settings = ContractSettings(
            modality_code="MOD_001", initial_balance=Decimal("10000"),
            credit_date=date(2023, 8, 22), annual_interest_rate=Decimal("0.08"),
            term=12, competence_start_day=21, index_code="INPC", index_lag_months=2,
        )
        values = self.zero_indices().values.copy()
        values["2023-06"] = Decimal("0.01")
        result = calculate_contract_evolution(settings, IndexSeries("INPC", values), set(), [])
        corrected = next(row for row in result.daily_rows if row.daily_correction_rate != 0)
        expected = corrected.balance_after_interest * corrected.daily_correction_rate
        self.assertLess(abs(corrected.correction_amount - expected), Decimal("0.000001"))

    def test_initial_date_excluded_and_due_date_included(self):
        settings = ContractSettings(
            modality_code="MOD_001", initial_balance=Decimal("10000"),
            credit_date=date(2023, 8, 22), annual_interest_rate=Decimal("0.08"),
            term=12, competence_start_day=21, index_code="INPC", index_lag_months=2,
        )
        result = calculate_contract_evolution(settings, self.zero_indices(), set(), [])
        first_rows = [r for r in result.daily_rows if r.installment_number == 1]
        self.assertEqual(first_rows[0].day, date(2023, 8, 23))
        self.assertEqual(first_rows[-1].day, date(2023, 9, 20))


    def test_daily_rates_and_index_change_on_competence_boundary(self):
        settings = ContractSettings(
            modality_code="MOD_001", initial_balance=Decimal("95000"),
            credit_date=date(2017, 11, 1), annual_interest_rate=Decimal("0.0769"),
            term=96, competence_start_day=21, index_code="INPC", index_lag_months=2,
        )
        values = self.zero_indices().values.copy()
        values["2017-08"] = Decimal("-0.0003")
        values["2017-09"] = Decimal("-0.0002")
        result = calculate_contract_evolution(settings, IndexSeries("INPC", values), set(), [])
        nov20 = next(r for r in result.daily_rows if r.day == date(2017, 11, 20))
        nov21 = next(r for r in result.daily_rows if r.day == date(2017, 11, 21))
        self.assertEqual(nov20.competence, "2017-10")
        self.assertEqual(nov20.index_reference, "2017-08")
        self.assertEqual(nov21.competence, "2017-11")
        self.assertEqual(nov21.index_reference, "2017-09")
        self.assertNotEqual(nov20.daily_interest_rate, nov21.daily_interest_rate)
        self.assertNotEqual(nov20.daily_correction_rate, nov21.daily_correction_rate)

    def test_extraordinary_reduces_balance(self):
        settings = ContractSettings(
            modality_code="MOD_001", initial_balance=Decimal("12000"),
            credit_date=date(2023, 8, 22), annual_interest_rate=Decimal("0.08"),
            term=12, competence_start_day=21, index_code="INPC", index_lag_months=2,
        )
        normal = calculate_contract_evolution(settings, self.zero_indices(), set(), [])
        extra = calculate_contract_evolution(
            settings, self.zero_indices(), set(),
            [ExtraordinaryAmortization(date(2023, 9, 10), Decimal("1000"))],
        )
        self.assertLess(extra.installments[0].closing_balance, normal.installments[0].closing_balance)


if __name__ == "__main__":
    unittest.main()
