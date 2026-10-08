"""Núcleo matemático do Motor de Cálculos — versão de validação."""

from .engines import calculate_contract_evolution
from .indices import IndexSeries, load_configuration

# O aplicativo importa build_opinion_pdf diretamente de core.pdf_reports.
# Mantemos o padrão visual aprovado e usamos a V17 para os ajustes editoriais
# dos pareceres Credplan Fixo e Variável.
from . import pdf_reports as _pdf_reports
from .pdf_reports_v17 import build_opinion_pdf as _build_opinion_pdf_v17

_pdf_reports.build_opinion_pdf = _build_opinion_pdf_v17

__all__ = ["calculate_contract_evolution", "IndexSeries", "load_configuration"]
