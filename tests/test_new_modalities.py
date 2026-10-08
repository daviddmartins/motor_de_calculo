from datetime import date
from decimal import Decimal

from core.engines import calculate_contract_evolution
from core.indices import IndexSeries
from core.models import ContractSettings
from core.pdf_reports import ReportIdentity, build_opinion_pdf


def _base_settings(code: str, *, index_code=None, payment_timing=0, term=96):
    return ContractSettings(
        modality_code=code,
        initial_balance=Decimal("64793.31"),
        credit_date=date(2015, 8, 17),
        annual_interest_rate=Decimal("0.079"),
        term=term,
        competence_start_day=21,
        index_code=index_code,
        index_lag_months=2,
        payment_timing=payment_timing,
    )


def test_mod_004_matches_tab_4_first_installments():
    result = calculate_contract_evolution(
        _base_settings("MOD_004", index_code=None, payment_timing=1),
        None,
        set(),
        [],
    )
    first, second = result.installments[:2]
    assert first.interest_amount == Decimal("451.84")
    assert first.regular_amortization == Decimal("452.46")
    assert first.installment_amount == Decimal("904.30")
    assert first.closing_balance == Decimal("64340.85")
    assert second.interest_amount == Decimal("408.97")
    assert second.regular_amortization == Decimal("495.33")
    assert second.installment_amount == Decimal("904.30")


def test_mod_005_matches_tab_5_monthly_and_annual_rule():
    result = calculate_contract_evolution(
        _base_settings("MOD_005", index_code=None, payment_timing=0),
        None,
        set(),
        [],
    )
    first, second = result.installments[:2]
    thirteenth = result.installments[12]

    assert first.interest_amount == Decimal("0.00")
    assert first.installment_amount == Decimal("903.75")
    assert first.regular_amortization == Decimal("903.75")
    assert first.closing_balance == Decimal("63889.56")

    assert second.installment_amount == Decimal("897.96")
    assert second.regular_amortization == Decimal("897.96")

    assert thirteenth.reference_due_date == date(2016, 9, 20)
    assert thirteenth.interest_amount == Decimal("4291.48")
    assert thirteenth.installment_amount == Decimal("902.73")
    assert thirteenth.regular_amortization == Decimal("902.73")
    assert thirteenth.closing_balance == Decimal("57711.23")


def test_mod_006_matches_tab_6_first_installment():
    indices = IndexSeries(
        "INPC",
        {
            "2015-05": Decimal("0.0099"),
            "2015-06": Decimal("0.0077"),
            "2015-07": Decimal("0.0058"),
        },
    )
    holidays = {date(2015, 9, 7)}
    result = calculate_contract_evolution(
        _base_settings("MOD_006", index_code="INPC", payment_timing=1),
        indices,
        holidays,
        [],
    )
    first = result.installments[0]
    assert first.interest_amount == Decimal("453.96")
    assert first.correction_amount == Decimal("584.80")
    assert first.installment_amount == Decimal("912.44")
    assert first.closing_balance == Decimal("64919.64")


def _zero_index():
    return IndexSeries(
        "INPC",
        {f"{year:04d}-{month:02d}": Decimal("0") for year in range(2014, 2020) for month in range(1, 13)},
    )


def _identity():
    return ReportIdentity(
        contract_number="TESTE-001",
        participant_name="Participante Teste",
        registration="0001",
        cpf="000.000.000-00",
        request_date=date(2026, 8, 15),
        net_amount=Decimal("64793.31"),
        elaborator="Teste automatizado",
        validator="Sem validação",
    )


def test_each_new_modality_generates_its_own_opinion_pdf():
    cases = [
        ("MOD_004", None, 1),
        ("MOD_005", None, 0),
        ("MOD_006", _zero_index(), 1),
    ]
    for code, index_series, timing in cases:
        settings = _base_settings(
            code,
            index_code="INPC" if code == "MOD_006" else None,
            payment_timing=timing,
            term=6,
        )
        result = calculate_contract_evolution(settings, index_series, set(), [])
        pdf = build_opinion_pdf(
            settings=settings,
            result=result,
            modality_name={
                "MOD_004": "Novo Credinâmico Fixo",
                "MOD_005": "Credinâmico Fixo",
                "MOD_006": "Novo Credinâmico Variável",
            }[code],
            identity=_identity(),
            selected_installment_numbers=[1, 2],
            overrides={},
            logo_path=None,
        )
        assert pdf.startswith(b"%PDF")
        assert len(pdf) > 10_000
