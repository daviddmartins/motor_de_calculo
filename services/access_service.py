from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from core.access_control import AccessEntry, AccessRegistry, RequestIdentity, discover_access_registry, detect_request_identity


@dataclass(frozen=True)
class AccessContext:
    registry: AccessRegistry
    identity: RequestIdentity
    entry: AccessEntry | None
    allowed: bool
    user_name: str
    profile: str
    is_admin: bool
    validator_options: tuple[str, ...]


def build_access_context(*, app_dir: str | Path, preferred_path: str | Path, headers: dict | None) -> AccessContext:
    registry = discover_access_registry(app_dir, preferred_path)
    identity = detect_request_identity(headers or {})
    entry = registry.find(identity.email)
    if registry.enabled:
        allowed = bool(entry and entry.active)
        user_name = entry.name.strip() if entry and entry.name.strip() else identity.fallback_name
        profile = entry.profile.strip() if entry else ""
        is_admin = bool(entry and entry.active and entry.is_admin)
    else:
        # Evita bloqueio acidental enquanto o arquivo oficial ainda está sendo preparado.
        allowed = True
        user_name = identity.fallback_name
        profile = "Administrador (validação)"
        is_admin = True
    validators = tuple(["Sem validação"] + registry.active_names)
    return AccessContext(
        registry=registry,
        identity=identity,
        entry=entry,
        allowed=allowed,
        user_name=user_name,
        profile=profile,
        is_admin=is_admin,
        validator_options=validators,
    )
