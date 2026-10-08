from __future__ import annotations

"""Pareceres Credplan V17.

Preserva o padrão visual aprovado e consolida quatro ajustes editoriais:
- remove o bloco redundante "Resumo da Prestação";
- garante que toda menção à defasagem reflita settings.index_lag_months;
- iguala a abertura de "Objeto e parâmetros contratuais" entre Fixo e Variável;
- substitui "Valor da parcela base" por "Defasagem do índice" nos parâmetros.
"""

import re

from reportlab.lib.units import mm
from reportlab.platypus import KeepTogether, Paragraph, Spacer, Table, TableStyle

from . import pdf_reports as base
from . import pdf_reports_v13 as v13
from . import pdf_reports_v14 as v14
from . import pdf_reports_v16 as v16

_BASE_BUILD_STORY = v16._build_story


def _lag_value(settings) -> str:
    if settings.modality_code == "MOD_002" or not settings.index_code:
        return "Não se aplica"
    lag = int(settings.index_lag_months or 0)
    if lag == 0:
        return "Sem defasagem"
    if lag == 1:
        return "1 mês"
    return f"{lag} meses"


def _lag_phrase(settings) -> str:
    lag = int(settings.index_lag_months or 0)
    if lag == 0:
        return "sem defasagem"
    if lag == 1:
        return "defasagem de 1 mês"
    return f"defasagem de {lag} meses"


def _intro_text(identity, modality_name, settings):
    contract = identity.contract_number.strip() or "-"
    return (
        "O presente parecer tem por objeto demonstrar, de forma técnica e padronizada, "
        f"a evolução do contrato de empréstimo <b>{modality_name} nº {contract}</b>. "
        "A análise considera os parâmetros contratuais informados, a metodologia financeira "
        "aplicável à modalidade e a memória de cálculo produzida pelo Motor de Cálculos, "
        "permitindo a reprodução e a conferência das etapas que formam as prestações e o saldo devedor."
    )


def _param_table(identity, settings, result, st):
    left = [
        ("Data da solicitação", v13._date(identity.request_date)),
        ("Data do crédito", v13._date(settings.credit_date)),
        ("Margem consignável", v13._money(identity.margin_amount)),
        ("Valor contratado", v13._money(settings.initial_balance)),
        ("FGQC - concessão", v13._money(identity.fgqc_concession)),
    ]
    right = [
        ("Saldo quitado na operação", v13._money(identity.settled_loan_balance)),
        ("Valor líquido", v13._money(identity.net_amount)),
        ("IOF", v13._money(identity.iof_amount)),
        ("Taxa administrativa", v13._money(identity.administrative_fee)),
        ("Defasagem do índice", _lag_value(settings)),
    ]

    rows = []
    for idx in range(5):
        l = left[idx]
        r = right[idx]
        rows.append([
            Paragraph(l[0], st["label"]),
            Paragraph(l[1], st["value"]),
            Paragraph(r[0], st["label"]),
            Paragraph(r[1], st["value"]),
        ])

    table = Table(rows, colWidths=[41.5 * mm, 47 * mm, 41.5 * mm, 47 * mm], hAlign="LEFT")
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.35, v13.LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (0, -1), v13.BLUE_SOFT2),
        ("BACKGROUND", (2, 0), (2, -1), v13.BLUE_SOFT2),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
    ]))
    return table


def _dynamic_paragraph(flowable: Paragraph, settings):
    """Reescreve defensivamente qualquer menção antiga à defasagem.

    V14 já usa a parametrização dinâmica. Este pós-processamento garante que
    nenhum texto legado da V13 volte a exibir 2 meses por ordem de importação.
    """
    try:
        plain = flowable.getPlainText()
        source = flowable.text
    except Exception:
        return flowable

    if settings.modality_code != "MOD_001":
        return flowable

    if "Nota sobre a competência 21-20." in plain:
        return Paragraph(v14._lag_note(settings), flowable.style)

    if plain.startswith("A correção monetária foi aplicada segundo o índice"):
        text = (
            f"A correção monetária foi aplicada segundo o índice e a {_lag_phrase(settings)} informados no sistema, "
            "observados os dias úteis da competência de aplicação. A redução do principal ocorreu por meio da "
            "amortização recalculada para o saldo atualizado e para o prazo remanescente, razão pela qual seu "
            "valor pode variar entre os ciclos."
        )
        return Paragraph(text, flowable.style)

    if plain.startswith("Os juros remuneram o capital utilizado no período") and "defasagem" in plain:
        text = (
            "Os juros remuneram o capital utilizado no período e não reduzem o principal. A redução do saldo "
            "decorre da amortização. A correção monetária integra a evolução diária do saldo, segundo o índice e a "
            f"{_lag_phrase(settings)} definidos na parametrização do cálculo, e pode alterar a base sobre a qual a "
            "amortização do ciclo será apurada."
        )
        return Paragraph(text, flowable.style)

    # Fórmula de correção: substitui apenas a frase de apoio, preservando a formatação.
    if "A defasagem parametrizada é de" in plain:
        lag = int(settings.index_lag_months or 0)
        if lag == 0:
            replacement = "Não há defasagem parametrizada."
        elif lag == 1:
            replacement = "A defasagem parametrizada é de 1 mês."
        else:
            replacement = f"A defasagem parametrizada é de {lag} meses."
        source = re.sub(
            r"A defasagem parametrizada é de\s+\d+\s+m[eê]s(?:es)?\.",
            replacement,
            source,
            flags=re.IGNORECASE,
        )
        return Paragraph(source, flowable.style)

    # Última proteção contra texto legado literal.
    if "defasagem de 2 meses" in plain.lower():
        source = re.sub(
            r"defasagem de 2 meses",
            _lag_phrase(settings),
            source,
            flags=re.IGNORECASE,
        )
        return Paragraph(source, flowable.style)

    return flowable


def _rewrite_nested(flowable, settings):
    if isinstance(flowable, Paragraph):
        return _dynamic_paragraph(flowable, settings)

    if isinstance(flowable, KeepTogether):
        content = getattr(flowable, "_content", None)
        if content is not None:
            flowable._content = [_rewrite_nested(item, settings) for item in content]
        return flowable

    if isinstance(flowable, Table):
        values = getattr(flowable, "_cellvalues", None)
        if values is not None:
            for row_idx, row in enumerate(values):
                for col_idx, cell in enumerate(row):
                    if isinstance(cell, (Paragraph, Table, KeepTogether)):
                        values[row_idx][col_idx] = _rewrite_nested(cell, settings)
                    elif isinstance(cell, list):
                        values[row_idx][col_idx] = [_rewrite_nested(item, settings) for item in cell]
        return flowable

    return flowable


def _remove_installment_summaries(story):
    cleaned = []
    idx = 0
    while idx < len(story):
        item = story[idx]
        if isinstance(item, Paragraph):
            try:
                plain = item.getPlainText().strip()
            except Exception:
                plain = ""
            if plain.startswith("Resumo da Prestação "):
                # Estrutura V16: título -> tabela de resumo -> espaçador.
                idx += 1
                if idx < len(story) and isinstance(story[idx], Table):
                    idx += 1
                if idx < len(story) and isinstance(story[idx], Spacer):
                    idx += 1
                continue
        cleaned.append(item)
        idx += 1
    return cleaned


def _build_story(tmp_dir, settings, result, modality_name, identity, selected_rows, overrides, st):
    # As funções chamadas pela montagem V14 são resolvidas no namespace V13.
    # Mantemos os dois ajustes abaixo explícitos antes da construção do story.
    v13._intro_text = _intro_text
    v13._param_table = _param_table

    story = list(_BASE_BUILD_STORY(
        tmp_dir,
        settings,
        result,
        modality_name,
        identity,
        selected_rows,
        overrides,
        st,
    ))

    story = _remove_installment_summaries(story)
    return [_rewrite_nested(item, settings) for item in story]


# V16/V14/V13 resolvem _build_story no namespace de V13.
v13._intro_text = _intro_text
v13._param_table = _param_table
v13._build_story = _build_story

build_opinion_pdf = v16.build_opinion_pdf

__all__ = ["build_opinion_pdf"]
