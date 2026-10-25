"""Where the context NIP and the KSeF token live - never in the repo or the workbook.

    ~/.ksefiz/<env>/context[.<profile>].json   NIP, token reference, permissions (no secrets)
    ~/.ksefiz/<env>/seal[.<profile>].pem       self-signed TEST seal + key (TEST only)
    OS keyring "ksefiz" / "<env>:<nip>"         the KSeF token
    KSEF_NIP / KSEF_TOKEN env vars override, TEST default profile only
Profiles let TEST hold several contexts (e.g. "buyer"); "default" is the seller context.
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import keyring

from ksefiz.config import Env

KEYRING_SERVICE = "ksefiz"
DEFAULT_PROFILE = "default"


def ksefiz_home() -> Path:
    return Path(os.environ.get("KSEFIZ_HOME", Path.home() / ".ksefiz"))


def state_dir(env: Env) -> Path:
    return ksefiz_home() / env.value


def state_file(env: Env, stem: str, suffix: str, profile: str = DEFAULT_PROFILE) -> Path:
    name = f"{stem}{suffix}" if profile == DEFAULT_PROFILE else f"{stem}.{profile}{suffix}"
    return state_dir(env) / name


@dataclass
class Context:
    nip: str
    token_reference: str = ""
    permissions: tuple[str, ...] = ()


def load_context(env: Env, profile: str = DEFAULT_PROFILE) -> Context | None:
    if env == Env.TEST and profile == DEFAULT_PROFILE and os.environ.get("KSEF_NIP"):
        return Context(os.environ["KSEF_NIP"])
    p = state_file(env, "context", ".json", profile)
    if not p.exists():
        return None
    d: dict[str, Any] = json.loads(p.read_text(encoding="utf-8"))
    return Context(d["nip"], d.get("token_reference", ""), tuple(d.get("permissions", ())))


def save_context(env: Env, ctx: Context, profile: str = DEFAULT_PROFILE, **extra: Any) -> Path:
    p = state_file(env, "context", ".json", profile)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps({"nip": ctx.nip, "token_reference": ctx.token_reference,
                             "permissions": list(ctx.permissions), **extra}, indent=2),
                 encoding="utf-8")
    return p


def write_private(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "wb") as f:
        f.write(data)


def get_token(env: Env, nip: str) -> str | None:
    if env == Env.TEST and os.environ.get("KSEF_TOKEN"):
        return os.environ["KSEF_TOKEN"]
    return keyring.get_password(KEYRING_SERVICE, f"{env.value}:{nip}")


def set_token(env: Env, nip: str, token: str) -> None:
    keyring.set_password(KEYRING_SERVICE, f"{env.value}:{nip}", token)
