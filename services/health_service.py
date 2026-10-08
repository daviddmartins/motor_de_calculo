from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .index_service import IndexRepository


OK = "OK"
WARNING = "ATENÇÃO"
ERROR = "ERRO"
INFO = "INFO"


@dataclass(frozen=True)
class HealthCheck:
    key: str
    label: str
    status: str
    detail: str
    target: str

    @property
    def requires_action(self) -> bool:
        return self.status in {WARNING, ERROR}


@dataclass(frozen=True)
class HealthSnapshot:
    checks: tuple[HealthCheck, ...]

    @property
    def errors(self) -> int:
        return sum(x.status == ERROR for x in self.checks)

    @property
    def warnings(self) -> int:
        return sum(x.status == WARNING for x in self.checks)

    @property
    def overall(self) -> str:
        if self.errors:
            return ERROR
        if self.warnings:
            return WARNING
        return OK

    @property
    def action_items(self) -> tuple[HealthCheck, ...]:
        return tuple(x for x in self.checks if x.requires_action)

    def as_rows(self) -> list[dict[str, str]]:
        return [
            {"Componente": x.label, "Situação": x.status, "Detalhe": x.detail}
            for x in self.checks
        ]


def _status_from_catalog(status: str) -> str:
    upper = (status or "").upper()
    if any(token in upper for token in ("A VALIDAR", "REVIS", "ERRO", "INVÁL")):
        return WARNING
    if "PARCIAL" in upper:
        return WARNING
    return OK


def build_health_snapshot(
    *,
    configuration_path: str | Path,
    configuration_loaded: bool,
    repository: IndexRepository | None,
    index_error: str,
    access_registry,
    holidays_count: int,
    engine_versions: dict[str, str] | None = None,
) -> HealthSnapshot:
    checks: list[HealthCheck] = []
    config_path = Path(configuration_path)
    checks.append(HealthCheck(
        key="config",
        label="Configuração geral",
        status=OK if configuration_loaded else ERROR,
        detail=(f"{config_path.name} carregado." if configuration_loaded else f"Não foi possível carregar {config_path.name}."),
        target="Configuração",
    ))

    if access_registry.enabled:
        active = len(access_registry.active_entries)
        checks.append(HealthCheck(
            key="access",
            label="Controle de acesso",
            status=OK if active else WARNING,
            detail=f"{active} usuário(s) ativo(s) na planilha reconhecida.",
            target="Acesso ao sistema",
        ))
    else:
        checks.append(HealthCheck(
            key="access",
            label="Controle de acesso",
            status=WARNING,
            detail="Planilha localizada, mas o controle ainda não foi ativado com e-mails.",
            target="Acesso ao sistema",
        ))

    if repository is None:
        checks.append(HealthCheck(
            key="index_base",
            label="Base única de índices",
            status=ERROR,
            detail=index_error or "indices_motor_calculos.xlsx não foi localizado ou é inválido.",
            target="Central de índices",
        ))
        return HealthSnapshot(tuple(checks))

    checks.append(HealthCheck(
        key="index_base",
        label="Base única de índices",
        status=OK,
        detail=f"{len(repository.active_definitions)} série(s) ativa(s) e {sum(x.record_count for x in repository.active_definitions.values())} registro(s).",
        target="Central de índices",
    ))

    required = [
        ("INPC", "INPC"),
        ("IPCA", "IPCA"),
        ("TAXA_LEGAL", "Taxa Legal"),
        ("TJSP_NOVA", "TJSP — Tabela Prática Nova"),
        ("TJSP_ANTIGA", "TJSP — Tabela Prática Antiga"),
        ("TJRJ_CIVEL_14905", "TJRJ — Cível / Lei 14.905"),
        ("TJMG_ICGJ", "TJMG — ICGJ"),
    ]
    for code, label in required:
        item = repository.definitions.get(code)
        if item is None or not item.active:
            checks.append(HealthCheck(code, label, ERROR, "Série obrigatória ausente ou inativa.", "Central de índices"))
            continue
        if not item.values:
            checks.append(HealthCheck(code, label, ERROR, "Série sem valores cadastrados.", "Central de índices"))
            continue
        status = _status_from_catalog(item.status)
        detail = f"Cobertura {item.first_competence} a {item.last_competence} · {item.record_count} registro(s)"
        if item.status:
            detail += f" · {item.status}"
        checks.append(HealthCheck(code, label, status, detail, "Central de índices"))

    if engine_versions:
        detail = " · ".join(f"{name}: {version}" for name, version in engine_versions.items())
        checks.append(HealthCheck(
            key="engines",
            label="Versões dos motores",
            status=INFO,
            detail=detail,
            target="Configuração",
        ))

    checks.append(HealthCheck(
        key="holidays",
        label="Calendário de feriados",
        status=OK,
        detail=(f"{holidays_count} feriado(s) cadastrado(s)." if holidays_count else "Sem feriados adicionais cadastrados; fins de semana continuam tratados pelo motor."),
        target="Feriados",
    ))
    return HealthSnapshot(tuple(checks))
