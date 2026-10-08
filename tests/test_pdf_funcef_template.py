from datetime import date
from decimal import Decimal
from pathlib import Path

from core.engines import calculate_contract_evolution
from core.indices import IndexSeries
from core.models import ContractSettings
from core.pdf_funcef_template import (
    CLASSIFICATION_LABEL, DOCUMENT_TITLE, FOOTER_LABEL, ManifestationHeader,
)
from core.pdf_reports import ReportIdentity, build_opinion_pdf


def _identity(header):
    return ReportIdentity(
        contract_number="TESTE-001",
        participant_name="",
        registration="",
        cpf="",
        request_date=date(2026, 8, 15),
        header=header,
    )


def _settings(code):
    return ContractSettings(
        modality_code=code,
        initial_balance=Decimal("10000"),
        credit_date=date(2023, 8, 22),
        annual_interest_rate=Decimal("0.1353"),
        term=4,
        competence_start_day=21,
        due_day=20,
        index_code="INPC" if code == "MOD_001" else None,
        index_lag_months=2 if code == "MOD_001" else 0,
    )


def _indices():
    return IndexSeries(
        "INPC",
        {f"{year:04d}-{month:02d}": Decimal("0") for year in (2023, 2024) for month in range(1, 13)},
    )


def test_parecer_credplan_inclui_abertura_e_responsavel_do_modelo():
    header = ManifestationHeader(
        issue_date=date(2026, 9, 21),
        process_number="0001234-56.2026.8.07.0001",
        district="Brasília",
        court="1ª Vara Cível",
        state="DF",
        lawyer="Advogado Teste & Associados",
        subject="Revisão contratual",
        reference="Ofício <123>",
        responsible="Analista Teste",
    )
    for code in ("MOD_001", "MOD_002"):
        settings = _settings(code)
        result = calculate_contract_evolution(
            settings, _indices() if code == "MOD_001" else None, set(), []
        )
        pdf = build_opinion_pdf(
            settings=settings,
            result=result,
            modality_name="Credplan",
            identity=_identity(header),
            selected_installment_numbers=[1],
            overrides={},
            logo_path=None,
        )
        assert pdf.startswith(b"%PDF")
        assert len(pdf) > 10_000


def test_parecer_sem_cabecalho_informado_usa_valores_padrao():
    settings = _settings("MOD_002")
    result = calculate_contract_evolution(settings, None, set(), [])
    pdf = build_opinion_pdf(
        settings=settings,
        result=result,
        modality_name="Credplan Fixo",
        identity=_identity(None),
        selected_installment_numbers=[1],
        overrides={},
        logo_path=None,
    )
    assert pdf.startswith(b"%PDF")


def test_modelo_funcef_preserva_textos_institucionais():
    source = Path("core/pdf_funcef_template.py").read_text(encoding="utf-8")
    assert DOCUMENT_TITLE == "Manifestação de Subsídios"
    assert CLASSIFICATION_LABEL == "#10 Corporativo - FUNCEF"
    assert FOOTER_LABEL == "COPART · GERAT · DIBEN"
    for text in (
        "Dados do Processo", "Participante e Operação", "Demanda",
        "Responsável pela informação", "#1C2541", "#31859B", "#0B5B91", "#DBE5F1",
    ):
        assert text in source
    assert "Conferido por" not in source


def _all_modalities():
    from core.indices import IndexSeries as _Series
    index = _Series("INPC", {f"{y:04d}-{m:02d}": Decimal("0.004") for y in range(2022, 2026) for m in range(1, 13)})
    for code in ("MOD_001", "MOD_002", "MOD_004", "MOD_005", "MOD_006"):
        variable = code in ("MOD_001", "MOD_006")
        settings = ContractSettings(
            modality_code=code, initial_balance=Decimal("50000"), credit_date=date(2023, 8, 22),
            annual_interest_rate=Decimal("0.1353"), term=12, index_code="INPC" if variable else None,
            index_lag_months=2 if variable else 0, payment_timing=1 if code in ("MOD_004", "MOD_006") else 0,
        )
        yield code, settings, calculate_contract_evolution(settings, index if variable else None, set(), [])


def test_pdf_usa_uma_unica_familia_tipografica():
    import re
    for code, settings, result in _all_modalities():
        pdf = build_opinion_pdf(
            settings=settings, result=result, modality_name=code, identity=_identity(None),
            selected_installment_numbers=[1, 2], overrides={}, logo_path=None,
        )
        fonts = set(re.findall(rb"/BaseFont /([A-Za-z-]+)", pdf))
        assert fonts and fonts <= {b"Helvetica", b"Helvetica-Bold", b"Helvetica-Oblique", b"Helvetica-BoldOblique"}, (code, fonts)



def test_cabecalho_manual_inclui_participante_no_pdf_e_revisa_campos():
    from core.pdf_funcef_template import header_review, opening_story
    from reportlab.platypus import Paragraph, Table

    header = ManifestationHeader(
        process_number="0001234-56.2026.8.07.0001", borrowers="Fulano de Tal; Beltrana",
        registrations="123456", responsible="Analista",
    )
    texts = []

    def collect(item):
        if isinstance(item, Paragraph):
            texts.append(item.getPlainText())
        elif isinstance(item, Table):
            for row in item._cellvalues:
                for cell in row:
                    for sub in (cell if isinstance(cell, list) else [cell]):
                        collect(sub)

    for item in opening_story(header, width=500, logo_path=None, contract_number="999", modality_name="C Fixo"):
        collect(item)
    assert "Fulano de Tal; Beltrana" in texts
    assert "123456" in texts
    assert "999" in texts and "C Fixo" in texts

    notices = header_review(header)
    assert not any("CNJ" in n for n in notices)
    assert any("Comarca" in n for n in notices)
    assert any("CNJ" in n for n in header_review(ManifestationHeader(process_number="123")))
