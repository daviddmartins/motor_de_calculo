from decimal import Decimal
from pathlib import Path

from services.index_service import load_index_workbook, repository_to_xlsx_bytes


ROOT = Path(__file__).resolve().parents[1]


def test_unified_index_base_has_required_series():
    repo = load_index_workbook(ROOT / "config" / "indices_motor_calculos.xlsx")
    required = {
        "INPC", "IPCA", "TAXA_LEGAL", "IGPM", "TR",
        "TJSP_NOVA", "TJSP_ANTIGA", "TJRJ_CIVEL_14905", "TJMG_ICGJ",
    }
    assert required.issubset(repo.active_definitions)
    assert repo.get("TJSP_NOVA").value_type == "FATOR"
    assert repo.get("INPC").value_type == "PERCENTUAL"


def test_percent_values_are_converted_only_at_runtime():
    repo = load_index_workbook(ROOT / "config" / "indices_motor_calculos.xlsx")
    item = repo.get("TAXA_LEGAL")
    comp = item.first_competence
    sheet_value = item.values[comp]
    runtime_value = repo.percent_series["TAXA_LEGAL"].values[comp]
    assert runtime_value == sheet_value / Decimal("100")


def test_unified_base_roundtrip_preserves_record_counts():
    repo = load_index_workbook(ROOT / "config" / "indices_motor_calculos.xlsx")
    rebuilt = load_index_workbook(repository_to_xlsx_bytes(repo))
    assert set(rebuilt.definitions) == set(repo.definitions)
    assert {
        code: item.record_count for code, item in rebuilt.definitions.items()
    } == {
        code: item.record_count for code, item in repo.definitions.items()
    }
