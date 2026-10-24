import base64
import copy
import hashlib
import json
from datetime import UTC, datetime, timedelta

import httpx
import pytest
import respx
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding, rsa
from lxml import etree

from ksefiz import auth, testenv, xades
from ksefiz.client import ApiError, KsefClient
from ksefiz.config import BASE_URLS, Env

BASE = BASE_URLS[Env.TEST]
NIP = "6261819874"


def _oaep() -> padding.OAEP:
    return padding.OAEP(mgf=padding.MGF1(hashes.SHA256()), algorithm=hashes.SHA256(), label=None)


MfKey = tuple[rsa.RSAPrivateKey, list[dict[str, object]]]


@pytest.fixture
def mf_key() -> MfKey:
    """A fake MF key published three ways: wrong usage, expired, and the valid one."""
    cert, key = xades.self_signed_seal("1111111111", "fake MF")
    der = base64.b64encode(cert.public_bytes(serialization.Encoding.DER)).decode()
    now = datetime.now(UTC)
    iso = lambda d: d.isoformat()  # noqa: E731
    certs: list[dict[str, object]] = [
        {"certificate": der, "publicKeyId": "S" * 44, "usage": ["SymmetricKeyEncryption"],
         "validFrom": iso(now - timedelta(days=1)), "validTo": iso(now + timedelta(days=1)),
         "certificateId": "a"},
        {"certificate": der, "publicKeyId": "E" * 44, "usage": ["KsefTokenEncryption"],
         "validFrom": iso(now - timedelta(days=9)), "validTo": iso(now - timedelta(days=1)),
         "certificateId": "b"},
        {"certificate": der, "publicKeyId": "K" * 44, "usage": ["KsefTokenEncryption"],
         "validFrom": iso(now - timedelta(days=1)), "validTo": iso(now + timedelta(days=1)),
         "certificateId": "c"},
    ]
    return key, certs


def _client() -> KsefClient:
    return KsefClient(Env.TEST, sleep=lambda s: None)


def _tokens() -> dict[str, object]:
    until = (datetime.now(UTC) + timedelta(minutes=15)).isoformat()
    return {"accessToken": {"token": "ACCESS", "validUntil": until},
            "refreshToken": {"token": "REFRESH", "validUntil": until}}


def test_encrypt_token_round_trip() -> None:
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    enc = auth.encrypt_token(key.public_key(), "tok-123", 1761000000000)
    assert key.decrypt(base64.b64decode(enc), _oaep()) == b"tok-123|1761000000000"


@respx.mock
def test_login_with_token_full_flow(mf_key: MfKey) -> None:
    key, certs = mf_key
    respx.post(f"{BASE}/auth/challenge").respond(json={
        "challenge": "CH-1", "timestamp": "2026-10-24T00:00:00Z", "timestampMs": 1761264000000,
        "clientIp": "127.0.0.1"})
    respx.get(f"{BASE}/security/public-key-certificates").respond(json=certs)
    init = respx.post(f"{BASE}/auth/ksef-token").respond(202, json={
        "referenceNumber": "REF-1", "authenticationToken": {"token": "TMP", "validUntil": "x"}})
    respx.get(f"{BASE}/auth/REF-1").mock(side_effect=[
        httpx.Response(200, json={"status": {"code": 100, "description": "w toku"}}),
        httpx.Response(200, json={"status": {"code": 200, "description": "ok"}}),
    ])
    redeem = respx.post(f"{BASE}/auth/token/redeem").respond(json=_tokens())

    with _client() as c:
        s = auth.login_with_token(c, NIP, "SECRET", sleep=lambda x: None)

    body = json.loads(init.calls[0].request.content)
    assert body["contextIdentifier"] == {"type": "Nip", "value": NIP}
    assert body["publicKeyId"] == "K" * 44
    assert key.decrypt(base64.b64decode(body["encryptedToken"]), _oaep()) == (
        b"SECRET|1761264000000")
    assert init.calls[0].request.headers["X-Error-Format"] == "problem-details"
    assert redeem.calls[0].request.headers["Authorization"] == "Bearer TMP"
    assert s.access_token == "ACCESS" and s.reference_number == "REF-1"
    assert "ACCESS" not in repr(s) and "REFRESH" not in repr(s)


@respx.mock
def test_auth_failure_status_raises() -> None:
    respx.get(f"{BASE}/auth/REF-2").respond(json={
        "status": {"code": 450, "description": "Błędny token", "details": ["Token nieaktywny"]}})
    init = {"referenceNumber": "REF-2", "authenticationToken": {"token": "TMP"}}
    with _client() as c, pytest.raises(auth.AuthError, match="450.*Token nieaktywny"):
        auth._wait_and_redeem(c, init, sleep=lambda x: None)


@respx.mock
def test_429_retry_after_is_respected() -> None:
    route = respx.get(f"{BASE}/rate-limits").mock(side_effect=[
        httpx.Response(429, headers={"Retry-After": "3"}, json={"status": {"code": 429}}),
        httpx.Response(200, json={"ok": True}),
    ])
    slept: list[float] = []
    with KsefClient(Env.TEST, sleep=slept.append) as c:
        assert c.request("GET", "/rate-limits") == {"ok": True}
    assert route.call_count == 2 and slept == [3.0]


@respx.mock
def test_problem_details_error_codes() -> None:
    respx.post(f"{BASE}/collective-identifiers/query").respond(400, json={
        "title": "Bad Request", "status": 400, "detail": "Żądanie jest nieprawidłowe.",
        "errors": [{"code": 21418, "description": "Nieprawidłowy token kontynuacji"}],
        "traceId": "T-1", "instance": "/x", "timestamp": "2026-10-24T00:00:00Z"},
        headers={"Content-Type": "application/problem+json"})
    with _client() as c, pytest.raises(ApiError) as ei:
        c.request("POST", "/collective-identifiers/query", json={})
    assert ei.value.codes == [21418] and ei.value.trace_id == "T-1"
    assert "kontynuacji" in str(ei.value)


def test_xades_signature_verifies() -> None:
    cert, key = xades.self_signed_seal(NIP)
    out = xades.sign(xades.auth_token_request("CH-1", NIP), cert, key)
    doc = etree.fromstring(out)
    ns = {"ds": xades.DS, "xades": xades.XADES}
    sig = doc.find("ds:Signature", ns)
    si = sig.find("ds:SignedInfo", ns)

    def c14n(e: etree._Element) -> bytes:
        return etree.tostring(e, method="c14n", exclusive=True)

    def digest(b: bytes) -> str:
        return base64.b64encode(hashlib.sha256(b).digest()).decode()

    refs = si.findall("ds:Reference", ns)
    unsigned = copy.deepcopy(doc)
    unsigned.remove(unsigned.find("ds:Signature", ns))
    assert digest(c14n(unsigned)) == refs[0].find("ds:DigestValue", ns).text
    sp = doc.find(".//xades:SignedProperties", ns)
    assert digest(c14n(sp)) == refs[1].find("ds:DigestValue", ns).text
    cert.public_key().verify(base64.b64decode(sig.find("ds:SignatureValue", ns).text),
                             c14n(si), padding.PKCS1v15(), hashes.SHA256())
    subject = cert.subject.get_attributes_for_oid(xades.ORGANIZATION_IDENTIFIER)
    assert subject[0].value == f"VATPL-{NIP}"


@pytest.mark.parametrize("env", [Env.DEMO, Env.PROD])
def test_bootstrap_refuses_outside_test(env: Env) -> None:
    with KsefClient(env) as c, pytest.raises(RuntimeError, match="TEST"):
        testenv.bootstrap(c, NIP, echo=lambda s: None)


def test_random_test_nip_is_never_real() -> None:
    from ksefiz.synthetic import nip_checksum_ok

    assert not any(nip_checksum_ok(testenv.random_test_nip()) for _ in range(200))
