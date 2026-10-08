from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO

from core.indices import AppConfiguration, load_configuration
from core.court_profiles import load_factor_series
from .index_service import IndexDiscovery, IndexRepository, discover_index_workbook, load_index_workbook


@dataclass(frozen=True)
class RuntimeConfiguration:
    configuration: AppConfiguration
    indices: IndexRepository | None
    index_source: str
    index_error: str = ""


def load_runtime_configuration(
    *,
    app_dir: str | Path,
    configuration_source: str | Path | bytes | BinaryIO,
    preferred_index_path: str | Path,
    active_index_bytes: bytes | None = None,
) -> RuntimeConfiguration:
    """Carrega parâmetros gerais e injeta a base única de índices.

    Se ``indices_motor_calculos.xlsx`` estiver disponível e válido, ele passa a ser
    a fonte autoritativa de séries econômicas. Se ainda não estiver, a configuração
    antiga continua funcionando como fallback durante a migração.
    """
    cfg = load_configuration(configuration_source)
    repo: IndexRepository | None = None
    index_source = ""
    index_error = ""

    if active_index_bytes:
        try:
            repo = load_index_workbook(active_index_bytes)
            index_source = "Base de índices carregada nesta sessão"
        except Exception as exc:
            index_error = str(exc)
    else:
        discovery = discover_index_workbook(app_dir, preferred_index_path)
        repo = discovery.repository
        index_source = str(discovery.path)
        index_error = discovery.error

    if repo is not None:
        # A base única é autoritativa para índices percentuais.
        cfg.indices = dict(repo.percent_series)

    return RuntimeConfiguration(
        configuration=cfg,
        indices=repo,
        index_source=index_source,
        index_error=index_error,
    )


def factor_series_for_runtime(
    *,
    app_dir: str | Path,
    repository: IndexRepository | None,
) -> dict[str, object]:
    """Retorna fatores judiciais da base única, com fallback legado durante a migração."""
    factors: dict[str, object] = {}
    if repository is not None:
        factors.update(repository.factor_series)
    # Compatibilidade de uma versão: arquivos CSV antigos só preenchem séries ausentes.
    legacy_dir = Path(app_dir) / "data" / "court_profiles"
    try:
        legacy = load_factor_series(legacy_dir)
    except Exception:
        legacy = {}
    for code, factor in legacy.items():
        factors.setdefault(code, factor)
    return factors
