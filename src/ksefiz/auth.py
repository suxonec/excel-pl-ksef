"""KSeF authentication: challenge -> (KSeF token | XAdES) -> poll -> redeem -> refresh -> close.

Token values are never logged or written anywhere; they live in memory only.
"""

from __future__ import annotations

import base64
import contextlib
import time
from collections.abc import Callable
from dataclasses import dataclass
from datetime import UTC, datetime
from typing import Any

from cryptography import x509
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding, rsa

from ksefiz.client import ApiError, KsefClient

STATUS_IN_PROGRESS = 100
STATUS_SUCCESS = 200
POLL_INTERVAL_S = 1.0
POLL_TIMEOUT_S = 120.0
REFRESH_MARGIN_S = 60.0


class AuthError(Exception):
    pass


@dataclass
class Session:
    access_token: str
    access_valid_until: datetime
    refresh_token: str
    refresh_valid_until: datetime
    reference_number: str

    def __repr__(self) -> str:  # keep tokens out of tracebacks and logs
        return (f"Session(reference_number={self.reference_number!r}, "
                f"access_valid_until={self.access_valid_until.isoformat()})")


def _ts(s: str) -> datetime:
    return datetime.fromisoformat(s.replace("Z", "+00:00"))


def mf_public_key(
    client: KsefClient, usage: str, now: datetime | None = None
) -> tuple[rsa.RSAPublicKey, str]:
    """Public key and publicKeyId of the currently valid MF certificate for `usage`
    ("KsefTokenEncryption" or "SymmetricKeyEncryption")."""
    now = now or datetime.now(UTC)
    certs: Any = client.request("GET", "/security/public-key-certificates")
    for c in certs:
        usages = c.get("usage") or []
        usages = usages if isinstance(usages, list) else [usages]
        if usage not in usages:
            continue
        if not _ts(c["validFrom"]) <= now <= _ts(c["validTo"]):
            continue
        cert = x509.load_der_x509_certificate(base64.b64decode(c["certificate"]))
        key = cert.public_key()
        if isinstance(key, rsa.RSAPublicKey):
            return key, str(c.get("publicKeyId") or "")
    raise AuthError(f"no valid {usage} certificate published by KSeF")


def token_encryption_key(
    client: KsefClient, now: datetime | None = None
) -> tuple[rsa.RSAPublicKey, str]:
    return mf_public_key(client, "KsefTokenEncryption", now)


def rsa_oaep(key: rsa.RSAPublicKey, data: bytes) -> bytes:
    oaep = padding.OAEP(mgf=padding.MGF1(hashes.SHA256()), algorithm=hashes.SHA256(), label=None)
    return key.encrypt(data, oaep)


def encrypt_token(key: rsa.RSAPublicKey, ksef_token: str, timestamp_ms: int) -> str:
    plain = f"{ksef_token}|{timestamp_ms}".encode()
    return base64.b64encode(rsa_oaep(key, plain)).decode("ascii")


def challenge(client: KsefClient) -> dict[str, Any]:
    result: dict[str, Any] = client.request("POST", "/auth/challenge")
    return result


def _wait_and_redeem(
    client: KsefClient, init: dict[str, Any], sleep: Callable[[float], None]
) -> Session:
    ref = init["referenceNumber"]
    auth_token = init["authenticationToken"]["token"]
    deadline = time.monotonic() + POLL_TIMEOUT_S
    while True:
        st = client.request("GET", f"/auth/{ref}", bearer=auth_token)
        code = int(st["status"]["code"])
        if code == STATUS_SUCCESS:
            break
        if code != STATUS_IN_PROGRESS:
            details = "; ".join(st["status"].get("details") or [])
            raise AuthError(f"authentication failed: {code} {st['status'].get('description')}"
                            + (f" ({details})" if details else ""))
        if time.monotonic() > deadline:
            raise AuthError(f"authentication still in progress after {POLL_TIMEOUT_S:.0f}s")
        sleep(POLL_INTERVAL_S)
    tokens = client.request("POST", "/auth/token/redeem", bearer=auth_token)
    return Session(
        access_token=tokens["accessToken"]["token"],
        access_valid_until=_ts(tokens["accessToken"]["validUntil"]),
        refresh_token=tokens["refreshToken"]["token"],
        refresh_valid_until=_ts(tokens["refreshToken"]["validUntil"]),
        reference_number=ref,
    )


def login_with_token(
    client: KsefClient, nip: str, ksef_token: str, sleep: Callable[[float], None] = time.sleep
) -> Session:
    """KSeF token route (the tool's normal login)."""
    ch = challenge(client)
    key, key_id = token_encryption_key(client)
    body: dict[str, Any] = {
        "challenge": ch["challenge"],
        "contextIdentifier": {"type": "Nip", "value": nip},
        "encryptedToken": encrypt_token(key, ksef_token, int(ch["timestampMs"])),
    }
    if key_id:
        body["publicKeyId"] = key_id
    return _wait_and_redeem(client, client.request("POST", "/auth/ksef-token", json=body), sleep)


def login_with_xades(
    client: KsefClient, signed_xml: bytes, sleep: Callable[[float], None] = time.sleep
) -> Session:
    """Certificate route: signed AuthTokenRequest (see xades.py). Used to bootstrap TEST."""
    init = client.request("POST", "/auth/xades-signature", content=signed_xml,
                          content_type="application/xml",
                          params={"verifyCertificateChain": "false"})
    return _wait_and_redeem(client, init, sleep)


def refresh(client: KsefClient, session: Session) -> Session:
    r = client.request("POST", "/auth/token/refresh", bearer=session.refresh_token)
    session.access_token = r["accessToken"]["token"]
    session.access_valid_until = _ts(r["accessToken"]["validUntil"])
    return session


def ensure_fresh(client: KsefClient, session: Session) -> Session:
    left = (session.access_valid_until - datetime.now(UTC)).total_seconds()
    return refresh(client, session) if left < REFRESH_MARGIN_S else session


def close(client: KsefClient, session: Session) -> None:
    """Revoke the session; never raises (best effort on exit)."""
    with contextlib.suppress(ApiError, OSError):
        client.request("DELETE", "/auth/sessions/current", bearer=session.access_token)
