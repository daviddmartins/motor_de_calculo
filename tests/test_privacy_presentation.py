from pathlib import Path

from core.majs import majs_interest_mode_label, majs_yearfrac_basis_label
from core.monetary_pdf import INTEREST_LABELS, INTEREST_SIMPLE_YEARFRAC


def test_user_facing_day_count_labels_do_not_reference_excel():
    labels = [
        majs_interest_mode_label("UM_PCT_YEARFRAC"),
        majs_interest_mode_label("UM_PCT_YEARFRAC_ATE_TL"),
        majs_yearfrac_basis_label(0),
        INTEREST_LABELS[INTEREST_SIMPLE_YEARFRAC],
    ]
    assert all("Excel" not in value for value in labels)
    assert "fração de ano" in labels[0].lower()


def test_web_app_does_not_collect_personal_identity_for_documents():
    source = Path("app.py").read_text(encoding="utf-8")
    assert 'st.text_input("Nome do participante"' not in source
    assert 'majs_report_participant = st.text_input' not in source
    assert 'participant_input = st.text_input' not in source
    assert 'elaborator=CURRENT_USER_NAME' not in source
    assert 'FRAÇÃOANO do Excel' not in source
    assert 'padrão do Excel' not in source


def test_pdf_sources_omit_placeholder_personal_identity():
    evolution = Path("core/pdf_reports.py").read_text(encoding="utf-8")
    majs = Path("core/majs_pdf.py").read_text(encoding="utf-8")
    monetary = Path("core/monetary_pdf.py").read_text(encoding="utf-8")
    assert 'identity.participant_name or "Não informado"' not in evolution
    assert 'identity.participant_name or "Não informado"' not in majs
    assert 'participant_name or "Não informado"' not in monetary
    assert "FRAÇÃOANO do Excel" not in majs
    assert "FRAÇÃOANO do Excel" not in monetary
