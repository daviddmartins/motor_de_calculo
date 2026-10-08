from pathlib import Path


def test_credplan_fixo_usa_design_system_aprovado():
    source = Path('core/pdf_reports.py').read_text(encoding='utf-8')

    assert 'UNIFICAÇÃO VISUAL CREDPLAN FIXO v0.10.8' in source
    assert "settings.modality_code == 'MOD_002'" in source
    assert '_fixed_formula_overview' in source
    assert '_fixed_detail_summary_table' in source
    assert '_fixed_detail_flowables' in source
    assert 'Natureza da metodologia de amortização - Tabela Price' in source
    assert 'Ausência de correção monetária e segregação dos juros' in source


def test_credplan_fixo_preserva_price_sem_correcao():
    source = Path('core/pdf_reports.py').read_text(encoding='utf-8')

    assert 'Prestação pela Tabela Price' in source
    assert 'A = PMT - J<sub>período</sub>' in source
    assert 'SD<sub>final</sub> = SD<sub>antes</sub> - PMT' in source
    assert 'Não se aplica' in source
