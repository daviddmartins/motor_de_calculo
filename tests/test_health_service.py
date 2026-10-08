from pathlib import Path

from services.health_service import build_health_snapshot, WARNING
from services.index_service import load_index_workbook


ROOT = Path(__file__).resolve().parents[1]


class _AccessRegistry:
    enabled = True
    active_entries = (object(), object())


def test_health_reports_catalog_validation_status_without_inventing_freshness():
    repo = load_index_workbook(ROOT / "config" / "indices_motor_calculos.xlsx")
    health = build_health_snapshot(
        configuration_path=ROOT / "config" / "configuracoes_validacao.xlsx",
        configuration_loaded=True,
        repository=repo,
        index_error="",
        access_registry=_AccessRegistry(),
        holidays_count=10,
    )
    inpc = next(x for x in health.checks if x.key == "INPC")
    assert inpc.status == WARNING  # catálogo da fonte registra A VALIDAR
    assert "Cobertura" in inpc.detail
