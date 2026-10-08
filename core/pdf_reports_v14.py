from __future__ import annotations

"""Ajustes V14 dos pareceres Credplan.

Mantém integralmente o padrão visual V13 aprovado e altera apenas:
- textos do Credplan Variável para explicitar o SAC;
- qualquer menção textual à defasagem do índice para refletir a parametrização
  efetivamente informada no sistema.
"""

from reportlab.lib.units import mm
from reportlab.platypus import CondPageBreak, HRFlowable, KeepTogether, Paragraph, Spacer

from . import pdf_reports as base
from . import pdf_reports_v13 as v13


def _lag_label(settings) -> str:
    lag = int(settings.index_lag_months or 0)
    if lag == 0:
        return "sem defasagem"
    if lag == 1:
        return "defasagem de 1 mês"
    return f"defasagem de {lag} meses"


def _lag_note(settings) -> str:
    lag = int(settings.index_lag_months or 0)
    index_name = (settings.index_code or "INPC").strip() or "INPC"
    if lag == 0:
        reference_text = (
            f"Sem defasagem, utiliza-se a referência do {index_name} correspondente "
            "à própria competência processada."
        )
    elif lag == 1:
        reference_text = (
            f"Com defasagem de 1 mês, utiliza-se a referência do {index_name} deslocada "
            "em 1 mês em relação à competência processada."
        )
    else:
        reference_text = (
            f"Com defasagem de {lag} meses, utiliza-se a referência do {index_name} "
            f"deslocada em {lag} meses em relação à competência processada."
        )
    return (
        "<b>Nota sobre a competência 21-20.</b> A competência é identificada pelo mês "
        "em que o ciclo se inicia, abrangendo o período do dia 21 ao dia 20 do mês "
        f"seguinte. {reference_text} O quadro acima apresenta a referência efetivamente "
        "utilizada no cálculo. Os juros são apropriados diariamente ao longo de todo o "
        "ciclo; a correção monetária é aplicada somente nos dias úteis."
    )


def _foundation_blocks(settings, st):
    fixed = settings.modality_code == "MOD_002"
    if fixed:
        # O Fixo permanece exatamente com a fundamentação aprovada na V13.
        return v13._foundation_blocks_original(settings, st) if hasattr(v13, "_foundation_blocks_original") else _FOUNDATION_V13(settings, st)

    blocks = [
        (
            "Natureza da metodologia de amortização — SAC",
            [
                "A modalidade Credplan Variável utiliza o <b>Sistema de Amortização Constante (SAC)</b>. "
                "Na metodologia aplicada pelo contrato, a amortização do ciclo é vinculada ao saldo devedor "
                "atualizado e ao prazo remanescente. A prestação é formada por juros e amortização, e a "
                "amortização é recalculada em cada ciclo a partir do saldo atualizado, deduzidos os juros do "
                "período, dividido pela quantidade de prestações remanescentes.",
                "Os juros remuneram o capital utilizado no período e não reduzem o principal. A redução do "
                "saldo decorre da amortização. A correção monetária integra a evolução diária do saldo, "
                f"segundo o índice e a {_lag_label(settings)} definidos na parametrização do cálculo, e pode "
                "alterar a base sobre a qual a amortização do ciclo será apurada.",
            ],
        ),
        (
            "Ausência de incorporação de juros vencidos ao principal",
            [
                "Na memória examinada, os juros são apropriados diariamente sobre o saldo vigente e liquidados "
                "como componente da prestação. Não se verifica a transferência de juros vencidos para o principal "
                "com o objetivo de formar uma nova base de incidência. A variação das prestações decorre da "
                "combinação entre saldo remanescente, correção monetária, juros do período e prazo ainda existente.",
                "Para a correta leitura da evolução, devem ser distinguidos a atualização monetária do saldo, os "
                "juros remuneratórios do período e a amortização do principal. Somente a amortização reduz o saldo "
                "devedor principal.",
            ],
        ),
    ]
    out = []
    for idx, (title, paras) in enumerate(blocks):
        parts = [Paragraph(title, st["subsection"])]
        for p in paras:
            parts.extend([Paragraph(p, st["body"]), Spacer(1, 1.5 * mm)])
        out.append(KeepTogether(parts))
        if idx == 0:
            out.append(HRFlowable(width="100%", thickness=0.4, color=v13.LINE, spaceBefore=1.0 * mm, spaceAfter=1.8 * mm))
    return out


# Guardamos a função original antes de substituí-la, para o Fixo continuar
# exatamente como aprovado.
_FOUNDATION_V13 = v13._foundation_blocks
if not hasattr(v13, "_foundation_blocks_original"):
    v13._foundation_blocks_original = _FOUNDATION_V13


def _build_story(tmp_dir, settings, result, modality_name, identity, selected_rows, overrides, st):
    story = [
        v13._section_title("Objeto e parâmetros contratuais", st),
        Paragraph(v13._intro_text(identity, modality_name, settings), st["body"]),
        Spacer(1, 1.8 * mm),
        KeepTogether([
            Paragraph(
                "INDICADORES PRINCIPAIS",
                v13.ParagraphStyle("v14_kpi_eyebrow", fontName="Helvetica-Bold", fontSize=7.0, leading=8.4, textColor=v13.MUTED),
            ),
            Spacer(1, 1.1 * mm),
            v13._KPIGrid(v13._kpis(settings, result, modality_name)),
        ]),
        Spacer(1, 3 * mm),
        KeepTogether([v13._param_table(identity, settings, result, st)]),
        Spacer(1, 3.3 * mm),
        v13._section_title("Fundamentação técnico-financeira", st),
    ]
    story += _foundation_blocks(settings, st)

    story += [
        Spacer(1, 2 * mm),
        v13._section_title("Fórmulas e aplicação numérica", st),
    ] + v13._formula_story(tmp_dir, settings, result, st)

    story += [
        Paragraph("Competências e taxas efetivamente utilizadas na primeira prestação", st["subsection"]),
        v13._segment_table(settings, result, st),
        Spacer(1, 1.6 * mm),
    ]
    if settings.modality_code == "MOD_002":
        story += [
            v13._info_note(
                "O Credplan Fixo não utiliza correção monetária. O quadro acima apresenta as competências "
                "efetivamente processadas até a primeira prestação e as taxas diárias de juros utilizadas na evolução do saldo.",
                st,
            ),
            Spacer(1, 3.2 * mm),
        ]
    else:
        story += [v13._info_note(_lag_note(settings), st), Spacer(1, 3.2 * mm)]

    story += [
        CondPageBreak(32 * mm),
        v13._section_title("Demonstração da evolução", st),
        Paragraph("Prestações de referência", st["subsection"]),
        v13._evolution_table(selected_rows, overrides, settings.rounding.money_places, st),
        Spacer(1, 4 * mm),
    ]

    for row in selected_rows:
        story += [
            CondPageBreak(34 * mm),
            KeepTogether([
                Paragraph(f"Prestação {row.installment_number} - referência {v13._date(row.reference_due_date)}", st["detail_title"]),
                Spacer(1, 1.5 * mm),
                Paragraph(v13._detail_narrative(row, overrides, settings), st["body"]),
                Spacer(1, 2 * mm),
            ]),
            v13._detail_table(row, overrides, settings, st),
            Spacer(1, 2.2 * mm),
            Paragraph(f"Resumo da Prestação {row.installment_number}", st["subsection"]),
            v13._summary_table(row, overrides, settings, st),
            Spacer(1, 1.3 * mm),
        ]
        for equation in v13._equations_for_row(row, overrides, settings):
            story += [v13._eq_strip(equation, st), Spacer(1, 0.7 * mm)]
        story += [Spacer(1, 2 * mm)]

    story += [CondPageBreak(22 * mm), v13._section_title("Conclusão técnica", st)]
    first = selected_rows[0]
    if settings.modality_code == "MOD_002":
        conclusion = [
            "Sob o ponto de vista estritamente técnico-financeiro e considerando a metodologia demonstrada, não foi identificada a incorporação de juros vencidos ao saldo principal para a formação de novos juros. O fator exponencial da Tabela Price representa a equivalência financeira necessária à apuração de uma prestação uniforme.",
            "Em cada período, os juros remuneram o saldo principal existente e a amortização, obtida pela diferença entre a prestação e os juros, reduz o saldo devedor. A parcela de juros não é abatida do principal nem incorporada a ele como nova base de cálculo. Não há correção monetária na modalidade Credplan Fixo.",
            f"Na primeira prestação apresentada, o valor de {v13._money(base._value(first, 'installment_amount', overrides))} é composto por juros de {v13._money(base._value(first, 'interest_amount', overrides))} e amortização de {v13._money(base._value(first, 'regular_amortization', overrides))}, resultando em saldo principal de {v13._money(base._value(first, 'closing_balance', overrides))} após a amortização. O FGQC, quando informado, é demonstrado separadamente e não integra a redução do principal.",
        ]
    else:
        conclusion = [
            "Sob o ponto de vista estritamente técnico-financeiro e considerando a metodologia demonstrada, não foi identificada a incorporação de juros vencidos ao saldo principal para a geração de novos encargos. Os juros foram apropriados diariamente sobre o saldo vigente e apresentados como componente da prestação.",
            f"A correção monetária foi aplicada segundo o índice e a {_lag_label(settings)} informados no sistema, observados os dias úteis da competência de aplicação. A redução do principal ocorreu por meio da amortização recalculada para o saldo atualizado e para o prazo remanescente, razão pela qual seu valor pode variar entre os ciclos.",
            f"Na primeira prestação de referência, a prestação de {v13._money(base._value(first, 'installment_amount', overrides))} compreende juros de {v13._money(base._value(first, 'interest_amount', overrides))} e amortização de {v13._money(base._value(first, 'regular_amortization', overrides))}, com saldo final de {v13._money(base._value(first, 'closing_balance', overrides))}. O FGQC de {v13._money(base._value(first, 'fgqc_amount', overrides))} é demonstrado separadamente.",
        ]
    for p in conclusion:
        story += [Paragraph(p, st["body"]), Spacer(1, 2.2 * mm)]

    if identity.additional_note.strip():
        story += [
            Spacer(1, 2 * mm),
            Paragraph("Observação da elaboração", st["subsection"]),
            Paragraph(identity.additional_note.strip().replace("\n", "<br/>"), st["body"]),
        ]
    return story


# O build da V13 resolve _build_story e _foundation_blocks no namespace do
# módulo V13. Substituímos somente essas duas rotinas de conteúdo; todo o
# desenho, KPIs, fórmulas, cabeçalho, logo, tabelas e paginação permanecem V13.
v13._foundation_blocks = _foundation_blocks
v13._build_story = _build_story

build_opinion_pdf = v13.build_opinion_pdf
