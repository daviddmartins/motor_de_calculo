from pathlib import Path


def test_evolution_pdf_source_omits_empty_identity_phrases():
    source = Path("core/pdf_reports.py").read_text(encoding="utf-8")
    assert 'em nome de ' not in source
    assert 'identity.participant_name}</b>, matrícula' not in source
    assert '("Elaboração", identity.elaborator or "Não informado")' not in source
    assert '("Validação", identity.validator or "Sem validação")' not in source


def test_web_form_does_not_require_personal_identity():
    source = Path("app.py").read_text(encoding="utf-8")
    assert 'st.text_input("Nome do participante"' not in source
    assert 'st.text_input("CPF"' not in source
    assert 'st.text_input("Matrícula"' not in source
    assert '("Nome do participante", participant_name)' not in source
    assert '("Matrícula", registration)' not in source
    assert '("CPF", cpf)' not in source
