from pathlib import Path
from datetime import date
from decimal import Decimal

from core.engines import calculate_contract_evolution
from core.indices import IndexSeries
from core.models import ContractSettings


def test_credplan_variavel_closing_balance_matches_principal_base_less_amortization():
    settings = ContractSettings(
        modality_code="MOD_001",
        initial_balance=Decimal("10000"),
        credit_date=date(2023, 8, 22),
        annual_interest_rate=Decimal("0.1353"),
        term=2,
        competence_start_day=21,
        due_day=20,
        index_code="INPC",
        index_lag_months=2,
    )
    indices = IndexSeries(
        "INPC",
        {
            "2023-06": Decimal("0"),
            "2023-07": Decimal("0"),
        },
    )

    result = calculate_contract_evolution(settings, indices, set(), [])

    for row in result.installments:
        assert row.amortization_base == row.balance_before_installment - row.interest_amount
        assert row.closing_balance == row.amortization_base - row.regular_amortization
        assert row.installment_amount == row.regular_amortization + row.interest_amount


def test_parecer_padrao_v0104_preserva_elementos_tecnicos():
    source = Path("core/pdf_reports.py").read_text(encoding="utf-8")
    assert "Ausência de incorporação de juros vencidos ao principal" in source
    assert "J<sub>d</sub> = SD<sub>0</sub> × i<sub>d</sub>" in source
    assert "_competence_2120_example(settings)" in source
    assert "defasagem de 2 meses da modalidade" not in source
    assert "Fórmulas e aplicação numérica" in source
    assert "Competências e taxas efetivamente utilizadas" in source


def test_parecer_v0106_usa_design_system_aprovado():
    source = Path("core/pdf_reports.py").read_text(encoding="utf-8")
    assert "PARECER DESIGN SYSTEM APROVADO v0.10.5" in source
    assert "PDF_BLUE = HexColor("#00366B")" in source
    assert "_approved_icon" in source
    assert "funcef.opening_story" in source
    assert "funcef.closing_story" in source
    assert "Resumo da Prestação" in source


def test_defasagem_do_parecer_respeita_parametro_do_sistema():
    from core.pdf_reports import _competence_2120_example, _lag_formula_note

    settings = ContractSettings(
        modality_code="MOD_001", initial_balance=Decimal("10000"),
        credit_date=date(2023, 8, 22), annual_interest_rate=Decimal("0.1353"),
        term=48, competence_start_day=21, due_day=20, index_code="INPC",
        index_lag_months=3,
    )
    example = _competence_2120_example(settings)
    formula = _lag_formula_note(settings)
    assert "defasagem de 3 meses" in example
    assert "INPC de outubro" in example
    assert "defasagem de 3 meses" in formula


def test_defasagem_singular_e_zero():
    from core.pdf_reports import _lag_label
    assert _lag_label(0) == "sem defasagem"
    assert _lag_label(1) == "defasagem de 1 mês"
    assert _lag_label(2) == "defasagem de 2 meses"
