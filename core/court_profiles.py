from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path
import csv

from .monetary_update import (
    CORRECTION_FACTOR_TABLE,
    CORRECTION_MONTHLY_CLOSE,
    INTEREST_NONE,
    INTEREST_BASE_CORRECTED,
    ORDER_CORRECTION_INTEREST,
    UpdateRule,
)


@dataclass(frozen=True)
class FactorSeries:
    code: str
    name: str
    values: dict[str, Decimal]
    source: str
    observation: str = ""

    @property
    def coverage_start(self) -> str:
        return min(self.values) if self.values else "-"

    @property
    def coverage_end(self) -> str:
        return max(self.values) if self.values else "-"


@dataclass(frozen=True)
class CourtProfile:
    code: str
    tribunal: str
    title: str
    description: str
    source_label: str
    source_url: str
    status: str
    factor_code: str | None = None
    note: str = ""


PROFILE_DEFINITIONS: tuple[CourtProfile, ...] = (
    CourtProfile(
        code="PERSONALIZADO", tribunal="LIVRE", title="Critério personalizado",
        description="Monte livremente uma ou mais regras por período: índice, juros, base e ordem de incidência.",
        source_label="Parâmetros definidos pelo usuário", source_url="", status="EDITÁVEL",
    ),
    CourtProfile(
        code="TJSP_NOVA", tribunal="TJSP", title="Tabela Prática — Lei 14.905/2024",
        description="Tabela oficial nova para débitos judiciais cíveis em geral. Usa a razão entre os fatores dos meses inicial e final.",
        source_label="TJSP · Nova Tabela Prática", source_url="arquivo fornecido pelo usuário", status="OFICIAL · TABELA",
        factor_code="TJSP_NOVA",
        note="A tabela informa somente correção monetária; juros moratórios devem ser configurados separadamente quando cabíveis.",
    ),
    CourtProfile(
        code="TJSP_ANTIGA", tribunal="TJSP", title="Antiga Tabela Prática",
        description="Tabela histórica do TJSP. Mantida para cálculos cuja regra aplicável determine o critério anterior.",
        source_label="TJSP · Antiga Tabela Prática", source_url="arquivo fornecido pelo usuário", status="OFICIAL · HISTÓRICA",
        factor_code="TJSP_ANTIGA",
        note="A tabela informa somente correção monetária; juros moratórios devem ser configurados separadamente quando cabíveis.",
    ),
    CourtProfile(
        code="TJRJ_CIVEL_14905", tribunal="TJRJ", title="Cível — Lei 6.899/81 / Lei 14.905/2024",
        description="Fatores do relatório cível com IPCA mensal a partir da Lei 14.905/2024 e fator IPCA para cálculo da Taxa Legal.",
        source_label="TJRJ · Central de Cálculos · relatório 01/06/2026",
        source_url="https://www.tjrj.jus.br/web/cgj/servicos/fatores-correcao-monetaria", status="OFICIAL · ATÉ 06/2026",
        factor_code="TJRJ_CIVEL_14905",
        note="O perfil carrega a correção monetária da tabela. Juros/Taxa Legal permanecem parametrizáveis na grade de regras.",
    ),
    CourtProfile(
        code="TJMG_ICGJ", tribunal="TJMG", title="ICGJ — Justiça Estadual",
        description="Fatores ICGJ usados para feitos em curso na Justiça Estadual de Minas Gerais, com a série oficial disponibilizada pelo TJMG.",
        source_label="TJMG · CADEJ / ICGJ · tabela de referência 06/2026",
        source_url="https://www8.tjmg.jus.br/cadej/pages/web/consulta-indice/indicadoresEconomicos.xhtml", status="OFICIAL · ATÉ 06/2026",
        factor_code="TJMG_ICGJ",
        note="O perfil carrega a correção monetária da tabela ICGJ. Juros permanecem configuráveis separadamente.",
    ),
    CourtProfile(
        code="TJDFT_OFICIAL", tribunal="TJDFT", title="Índices oficiais TJDFT",
        description="Pré-carrega INPC até 31/08/2024 e IPCA a partir de 01/09/2024, mantendo todas as linhas editáveis.",
        source_label="TJDFT · JurisCalc", source_url="https://www.tjdft.jus.br/servicos/atualizacao-monetaria-1", status="OFICIAL · PRESET",
        note="O perfil pré-carrega somente a correção (INPC até 31/08/2024 e IPCA a partir de 01/09/2024). Juros são opcionais e podem ser ativados em bloco próprio, com método e data inicial definidos pelo usuário.",
    ),
    CourtProfile(
        code="TJDFT_INPC", tribunal="TJDFT", title="INPC durante todo o período",
        description="Alternativa disponibilizada pelo TJDFT para decisões que determinem expressamente INPC em todo o período.",
        source_label="TJDFT · JurisCalc", source_url="https://www.tjdft.jus.br/servicos/atualizacao-monetaria-1", status="OFICIAL · PRESET",
        note="O perfil define somente INPC como correção. Juros são opcionais e podem ser ativados separadamente conforme o título judicial.",
    ),
)


def get_profiles() -> dict[str, CourtProfile]:
    return {p.code: p for p in PROFILE_DEFINITIONS}


def load_factor_series(data_dir: str | Path) -> dict[str, FactorSeries]:
    base = Path(data_dir)
    specs = {
        "TJSP_NOVA": ("tjsp_nova.csv", "TJSP — Nova Tabela Prática"),
        "TJSP_ANTIGA": ("tjsp_antiga.csv", "TJSP — Antiga Tabela Prática"),
        "TJRJ_CIVEL_14905": ("tjrj_civel_lei14905_jun2026.csv", "TJRJ — Cível Lei 14.905/2024"),
        "TJMG_ICGJ": ("tjmg_icgj_jun2026.csv", "TJMG — ICGJ"),
    }
    result: dict[str, FactorSeries] = {}
    for code, (filename, name) in specs.items():
        path = base / filename
        if not path.exists():
            continue
        values: dict[str, Decimal] = {}
        source = ""
        observation = ""
        with path.open("r", encoding="utf-8-sig", newline="") as fh:
            for row in csv.DictReader(fh):
                competence = (row.get("competencia") or "").strip()
                raw = (row.get("fator") or "").strip().replace(",", ".")
                if not competence or not raw:
                    continue
                values[competence] = Decimal(raw)
                source = row.get("fonte") or source
                observation = row.get("observacao") or observation
        if values:
            result[code] = FactorSeries(code=code, name=name, values=values, source=source, observation=observation)
    return result


def _clamp(start: date, end: date, a: date, b: date) -> tuple[date, date] | None:
    lo, hi = max(start, a), min(end, b)
    return (lo, hi) if lo <= hi else None


def build_profile_rules(profile_code: str, start_date: date, end_date: date) -> list[UpdateRule]:
    if start_date >= end_date:
        raise ValueError("A data inicial precisa ser anterior à data-base para aplicar um perfil.")

    if profile_code in {"TJSP_NOVA", "TJSP_ANTIGA", "TJRJ_CIVEL_14905", "TJMG_ICGJ"}:
        profile = get_profiles()[profile_code]
        return [UpdateRule(
            order=1, start_date=start_date, end_date=end_date,
            description=profile.title,
            correction_index_code=profile.factor_code or profile_code,
            correction_method=CORRECTION_FACTOR_TABLE,
            index_lag_months=0,
            monthly_interest_rate=Decimal("0"), interest_method=INTEREST_NONE,
            interest_base=INTEREST_BASE_CORRECTED, event_order=ORDER_CORRECTION_INTEREST,
            competence_start_day=1,
        )]

    if profile_code == "TJDFT_INPC":
        return [UpdateRule(
            order=1, start_date=start_date, end_date=end_date,
            description="TJDFT — INPC durante todo o período",
            correction_index_code="INPC", correction_method=CORRECTION_MONTHLY_CLOSE,
            index_lag_months=0, monthly_interest_rate=Decimal("0"), interest_method=INTEREST_NONE,
            interest_base=INTEREST_BASE_CORRECTED, event_order=ORDER_CORRECTION_INTEREST,
            competence_start_day=1,
        )]

    if profile_code == "TJDFT_OFICIAL":
        result: list[UpdateRule] = []
        first = _clamp(start_date, end_date, date(1900, 1, 1), date(2024, 8, 31))
        second = _clamp(start_date, end_date, date(2024, 9, 1), date(2100, 12, 31))
        if first:
            result.append(UpdateRule(
                order=len(result) + 1, start_date=first[0], end_date=first[1],
                description="TJDFT — INPC até 31/08/2024",
                correction_index_code="INPC", correction_method=CORRECTION_MONTHLY_CLOSE,
                index_lag_months=0, monthly_interest_rate=Decimal("0"), interest_method=INTEREST_NONE,
                interest_base=INTEREST_BASE_CORRECTED, event_order=ORDER_CORRECTION_INTEREST, competence_start_day=1,
            ))
        if second:
            result.append(UpdateRule(
                order=len(result) + 1, start_date=second[0], end_date=second[1],
                description="TJDFT — IPCA a partir de 01/09/2024",
                correction_index_code="IPCA", correction_method=CORRECTION_MONTHLY_CLOSE,
                index_lag_months=0, monthly_interest_rate=Decimal("0"), interest_method=INTEREST_NONE,
                interest_base=INTEREST_BASE_CORRECTED, event_order=ORDER_CORRECTION_INTEREST, competence_start_day=1,
            ))
        return result

    return [UpdateRule(
        order=1, start_date=start_date, end_date=end_date,
        description="Regra personalizada",
        correction_index_code="INPC", correction_method=CORRECTION_MONTHLY_CLOSE,
        index_lag_months=0, monthly_interest_rate=Decimal("0"), interest_method=INTEREST_NONE,
        interest_base=INTEREST_BASE_CORRECTED, event_order=ORDER_CORRECTION_INTEREST, competence_start_day=1,
    )]
