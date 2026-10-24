"""One-time bootstrap on the KSeF TEST environment (never DEMO/PROD):

1. random test NIP + self-signed company seal (organizationIdentifier VATPL-<nip>)
2. XAdES login with the seal -> owner session for that NIP
3. POST /tokens with CollectiveIdentifierManage only, wait until Active
4. token -> OS keyring, NIP + token reference -> ~/.ksefiz/test/context.json
"""

from __future__ import annotations

import random
import time
from collections.abc import Callable

from cryptography import x509
from cryptography.hazmat.primitives.asymmetric import rsa

from ksefiz import auth, xades
from ksefiz.client import KsefClient
from ksefiz.config import Env
from ksefiz.credentials import (
    DEFAULT_PROFILE,
    Context,
    load_context,
    save_context,
    set_token,
    state_file,
    write_private,
)
from ksefiz.synthetic import fake_nip

TOKEN_PERMISSIONS = ("CollectiveIdentifierManage",)
TOKEN_POLL_TIMEOUT_S = 120.0


def random_test_nip(rng: random.Random | None = None) -> str:
    """Matches the KSeF NIP pattern but fails the mod-11 checksum: cannot be a real company."""
    return fake_nip(rng or random.Random())


def bootstrap(
    client: KsefClient,
    nip: str,
    echo: Callable[[str], None] = print,
    sleep: Callable[[float], None] = time.sleep,
    profile: str = DEFAULT_PROFILE,
) -> Context:
    _require_test(client)
    cert, key = xades.self_signed_seal(nip)
    seal_path = state_file(Env.TEST, "seal", ".pem", profile)
    write_private(seal_path, xades.save_pem(cert, key))
    echo(f"self-signed seal for NIP {nip} saved to {seal_path}")

    session = _xades_login(client, nip, cert, key, sleep)
    echo(f"XAdES login OK (auth reference {session.reference_number})")
    try:
        gen = client.request("POST", "/tokens", bearer=session.access_token, json={
            "permissions": list(TOKEN_PERMISSIONS),
            "description": "ksefiz IZ tool - CollectiveIdentifierManage only (TEST)",
        })
        ref, token = gen["referenceNumber"], gen["token"]
        echo(f"KSeF token requested (reference {ref}), waiting for Active...")
        deadline = time.monotonic() + TOKEN_POLL_TIMEOUT_S
        while True:
            st = client.request("GET", f"/tokens/{ref}", bearer=session.access_token)
            status = st.get("status")
            if status == "Active":
                break
            if status in ("Failed", "Revoked", "Revoking"):
                raise RuntimeError(f"KSeF token ended in status {status}")
            if time.monotonic() > deadline:
                raise RuntimeError(f"KSeF token still {status} after {TOKEN_POLL_TIMEOUT_S:.0f}s")
            sleep(1.0)
        set_token(Env.TEST, nip, token)
        ctx = Context(nip, ref, tuple(st.get("requestedPermissions") or TOKEN_PERMISSIONS))
        path = save_context(Env.TEST, ctx, profile, created=st.get("dateCreated", ""))
        echo(f"token Active, stored in the OS keyring; context saved to {path}")
        return ctx
    finally:
        auth.close(client, session)


def _require_test(client: KsefClient) -> None:
    if client.env != Env.TEST:
        raise RuntimeError("TEST bootstrap refuses to run outside the TEST environment")


def _xades_login(client: KsefClient, nip: str, cert: x509.Certificate,
                 key: rsa.RSAPrivateKey, sleep: Callable[[float], None]) -> auth.Session:
    ch = auth.challenge(client)
    signed = xades.sign(xades.auth_token_request(ch["challenge"], nip), cert, key)
    return auth.login_with_xades(client, signed, sleep)


def seal_login(client: KsefClient, profile: str = DEFAULT_PROFILE,
               sleep: Callable[[float], None] = time.sleep) -> tuple[Context, auth.Session]:
    """Owner session for a bootstrapped TEST profile, via its stored self-signed seal."""
    _require_test(client)
    ctx = load_context(Env.TEST, profile)
    seal = state_file(Env.TEST, "seal", ".pem", profile)
    if ctx is None or not seal.exists():
        raise RuntimeError(f"TEST profile {profile!r} not set up; run "
                           f"`ksefiz testenv-setup --profile {profile}`")
    cert, key = xades.load_pem(seal.read_bytes())
    return ctx, _xades_login(client, ctx.nip, cert, key, sleep)
