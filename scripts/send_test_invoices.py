"""Delivery step 4 helper (NOT part of the tool): send synthetic FA(3) invoices on KSeF TEST
so that real TEST KSeF numbers exist for IZ generate / decode / lookup tests.

    python scripts/send_test_invoices.py                    # 5 seller->buyer + 1 buyer->seller
    python scripts/send_test_invoices.py --batch --count 500   # 500 seller->buyer, batch session

Needs the TEST profiles "default" (seller) and "buyer" (`ksefiz testenv-setup --profile buyer`).
Logs in with each profile's self-signed seal (owner rights - the CollectiveIdentifierManage
token deliberately cannot send invoices). Output goes OUTSIDE the repo:
~/.ksefiz/test/invoices.csv. TEST only.
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import io
import os
import time
import uuid
import zipfile
from collections.abc import Callable
from dataclasses import dataclass
from datetime import UTC, date, datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal
from urllib.parse import urlsplit
from xml.sax.saxutils import escape

import httpx
from cryptography.hazmat.primitives import padding as sympad
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes

from ksefiz import auth, testenv
from ksefiz.audit import AuditLog
from ksefiz.client import KsefClient
from ksefiz.config import Env
from ksefiz.credentials import load_context, state_dir
from ksefiz.io_csv import read_rows, write_rows
from ksefiz.iz import is_valid_ksef

FORM_CODE = {"systemCode": "FA (3)", "schemaVersion": "1-0E", "value": "FA"}
POLL_TIMEOUT_S = 180.0
OUT_FIELDS = ["seller_nip", "buyer_nip", "invoice_number", "ksef_number", "net", "vat",
              "gross", "currency", "issue_date", "session_reference", "invoice_reference", "set"]
BATCH_POLL_TIMEOUT_S = 900.0
# Batch part uploads go to a pre-signed URL from KSeF; refuse anything outside MF hosts
ALLOWED_UPLOAD_SUFFIX = ".mf.gov.pl"


@dataclass(frozen=True)
class Spec:
    seller_nip: str
    buyer_nip: str
    net: Decimal
    description: str


def _money(d: Decimal) -> str:
    return str(d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def fa3_xml(spec: Spec, number: str, issue: date) -> bytes:
    """Minimal synthetic FA(3) invoice: one 23% line, transfer payment, no bank account."""
    net = spec.net.quantize(Decimal("0.01"))
    vat = (net * Decimal("0.23")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    created = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<Faktura xmlns:etd="http://crd.gov.pl/xml/schematy/dziedzinowe/mf/2022/01/05/eD/DefinicjeTypy/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://crd.gov.pl/wzor/2025/06/25/13775/">
  <Naglowek>
    <KodFormularza kodSystemowy="FA (3)" wersjaSchemy="1-0E">FA</KodFormularza>
    <WariantFormularza>3</WariantFormularza>
    <DataWytworzeniaFa>{created}</DataWytworzeniaFa>
    <SystemInfo>ksefiz test data</SystemInfo>
  </Naglowek>
  <Podmiot1>
    <DaneIdentyfikacyjne><NIP>{spec.seller_nip}</NIP><Nazwa>TEST SELLER ksefiz</Nazwa></DaneIdentyfikacyjne>
    <Adres><KodKraju>PL</KodKraju><AdresL1>ul. Testowa 1</AdresL1><AdresL2>00-001 Warszawa</AdresL2></Adres>
  </Podmiot1>
  <Podmiot2>
    <DaneIdentyfikacyjne><NIP>{spec.buyer_nip}</NIP><Nazwa>TEST BUYER ksefiz</Nazwa></DaneIdentyfikacyjne>
    <Adres><KodKraju>PL</KodKraju><AdresL1>ul. Probna 2</AdresL1><AdresL2>00-002 Warszawa</AdresL2></Adres>
    <JST>2</JST>
    <GV>2</GV>
  </Podmiot2>
  <Fa>
    <KodWaluty>PLN</KodWaluty>
    <P_1>{issue.isoformat()}</P_1>
    <P_1M>Warszawa</P_1M>
    <P_2>{escape(number)}</P_2>
    <P_13_1>{_money(net)}</P_13_1>
    <P_14_1>{_money(vat)}</P_14_1>
    <P_15>{_money(net + vat)}</P_15>
    <Adnotacje>
      <P_16>2</P_16><P_17>2</P_17><P_18>2</P_18><P_18A>2</P_18A>
      <Zwolnienie><P_19N>1</P_19N></Zwolnienie>
      <NoweSrodkiTransportu><P_22N>1</P_22N></NoweSrodkiTransportu>
      <P_23>2</P_23>
      <PMarzy><P_PMarzyN>1</P_PMarzyN></PMarzy>
    </Adnotacje>
    <RodzajFaktury>VAT</RodzajFaktury>
    <FaWiersz>
      <NrWierszaFa>1</NrWierszaFa>
      <P_7>{escape(spec.description)}</P_7>
      <P_8A>szt.</P_8A>
      <P_8B>1</P_8B>
      <P_9A>{_money(net)}</P_9A>
      <P_11>{_money(net)}</P_11>
      <P_12>23</P_12>
    </FaWiersz>
    <Platnosc>
      <TerminPlatnosci><Termin>{(issue + timedelta(days=30)).isoformat()}</Termin></TerminPlatnosci>
      <FormaPlatnosci>6</FormaPlatnosci>
    </Platnosc>
  </Fa>
</Faktura>
"""
    return xml.encode("utf-8")


def _b64(b: bytes) -> str:
    return base64.b64encode(b).decode("ascii")


def _aes_cbc(key: bytes, iv: bytes, data: bytes) -> bytes:
    padder = sympad.PKCS7(128).padder()
    enc = Cipher(algorithms.AES(key), modes.CBC(iv)).encryptor()
    return enc.update(padder.update(data) + padder.finalize()) + enc.finalize()


def send(client: KsefClient, session: auth.Session, specs: list[Spec], issue: date,
         echo: Callable[[str], None] = print) -> list[dict[str, str]]:
    sym_key, iv = os.urandom(32), os.urandom(16)
    mf_key, key_id = auth.mf_public_key(client, "SymmetricKeyEncryption")
    encryption = {"encryptedSymmetricKey": _b64(auth.rsa_oaep(mf_key, sym_key)),
                  "initializationVector": _b64(iv)}
    if key_id:
        encryption["publicKeyId"] = key_id
    opened = client.request("POST", "/sessions/online", bearer=session.access_token,
                            json={"formCode": FORM_CODE, "encryption": encryption})
    sref = opened["referenceNumber"]
    echo(f"online session {sref} opened")

    sent: list[tuple[Spec, str, str, bytes]] = []
    for spec in specs:
        number = f"KSEFIZ/{issue:%Y%m%d}/{uuid.uuid4().hex[:10].upper()}"
        xml = fa3_xml(spec, number, issue)
        enc = _aes_cbc(sym_key, iv, xml)
        r = client.request("POST", f"/sessions/online/{sref}/invoices", bearer=session.access_token,
                           json={"invoiceHash": _b64(hashlib.sha256(xml).digest()),
                                 "invoiceSize": len(xml),
                                 "encryptedInvoiceHash": _b64(hashlib.sha256(enc).digest()),
                                 "encryptedInvoiceSize": len(enc),
                                 "encryptedInvoiceContent": _b64(enc)})
        sent.append((spec, number, r["referenceNumber"], xml))
        echo(f"  sent {number} ({spec.seller_nip} -> {spec.buyer_nip}, net {spec.net})")
    client.request("POST", f"/sessions/online/{sref}/close", bearer=session.access_token)

    rows = []
    for spec, number, iref, _xml in sent:
        deadline = time.monotonic() + POLL_TIMEOUT_S
        while True:
            st = client.request("GET", f"/sessions/{sref}/invoices/{iref}",
                                bearer=session.access_token)
            code = int(st["status"]["code"])
            if code == 200:
                break
            if code >= 400:
                raise RuntimeError(f"invoice {number} rejected: {code} "
                                   f"{st['status'].get('description')} "
                                   f"{st['status'].get('details') or ''}")
            if time.monotonic() > deadline:
                raise RuntimeError(f"invoice {number} still {code} after {POLL_TIMEOUT_S:.0f}s")
            time.sleep(2.0)
        ksef = st["ksefNumber"]
        rows.append(_row(spec, number, ksef, issue, sref, iref, "online"))
        echo(f"  {number} -> {ksef}  (CRC-8 valid: {is_valid_ksef(ksef)})")
    return rows


def _row(spec: Spec, number: str, ksef: str, issue: date, sref: str, iref: str,
         label: str) -> dict[str, str]:
    net = spec.net.quantize(Decimal("0.01"))
    vat = (net * Decimal("0.23")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return {"seller_nip": spec.seller_nip, "buyer_nip": spec.buyer_nip, "invoice_number": number,
            "ksef_number": ksef, "net": _money(net), "vat": _money(vat),
            "gross": _money(net + vat), "currency": "PLN", "issue_date": issue.isoformat(),
            "session_reference": sref, "invoice_reference": iref, "set": label}


def send_batch(client: KsefClient, session: auth.Session, specs: list[Spec], issue: date,
               label: str, echo: Callable[[str], None] = print) -> list[dict[str, str]]:
    """Batch session: one encrypted ZIP part with all invoices (part uploads are not
    rate-limited, unlike online sending at 180 invoices/hour)."""
    bearer = session.access_token
    by_hash: dict[str, tuple[Spec, str]] = {}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, spec in enumerate(specs, start=1):
            number = f"KSEFIZ/{issue:%Y%m%d}/{uuid.uuid4().hex[:10].upper()}"
            xml = fa3_xml(spec, number, issue)
            by_hash[_b64(hashlib.sha256(xml).digest())] = (spec, number)
            zf.writestr(f"invoice_{i:04d}.xml", xml)
    zip_bytes = buf.getvalue()
    sym_key, iv = os.urandom(32), os.urandom(16)
    part = _aes_cbc(sym_key, iv, zip_bytes)
    mf_key, key_id = auth.mf_public_key(client, "SymmetricKeyEncryption")
    encryption = {"encryptedSymmetricKey": _b64(auth.rsa_oaep(mf_key, sym_key)),
                  "initializationVector": _b64(iv)}
    if key_id:
        encryption["publicKeyId"] = key_id
    opened = client.request("POST", "/sessions/batch", bearer=bearer, json={
        "formCode": FORM_CODE,
        "batchFile": {"fileSize": len(zip_bytes), "fileHash": _b64(hashlib.sha256(zip_bytes).digest()),
                      "fileParts": [{"ordinalNumber": 1, "fileSize": len(part),
                                     "fileHash": _b64(hashlib.sha256(part).digest())}]},
        "encryption": encryption,
    })
    sref = opened["referenceNumber"]
    echo(f"batch session {sref} opened ({len(specs)} invoices, ZIP {len(zip_bytes)} bytes)")
    for up in opened["partUploadRequests"]:
        host = urlsplit(up["url"]).hostname or ""
        if not host.endswith(ALLOWED_UPLOAD_SUFFIX):
            raise RuntimeError(f"refusing part upload to non-MF host {host!r}")
        r = httpx.request(up["method"], up["url"], content=part,
                          headers={k: v for k, v in (up.get("headers") or {}).items() if v},
                          timeout=120)
        if r.status_code != 201:
            raise RuntimeError(f"part upload failed: HTTP {r.status_code} {r.text[:300]}")
        echo(f"  part {up['ordinalNumber']} uploaded to {host}")
    client.request("POST", f"/sessions/batch/{sref}/close", bearer=bearer)

    deadline = time.monotonic() + BATCH_POLL_TIMEOUT_S
    while True:
        st = client.request("GET", f"/sessions/{sref}", bearer=bearer)
        code = int(st["status"]["code"])
        if code == 200:
            break
        if code >= 300:
            raise RuntimeError(f"batch session failed: {code} {st['status'].get('description')} "
                               f"{st['status'].get('details') or ''}")
        if time.monotonic() > deadline:
            raise RuntimeError(f"batch session still {code} after {BATCH_POLL_TIMEOUT_S:.0f}s")
        time.sleep(5.0)
    echo(f"  processed: {st.get('successfulInvoiceCount')} ok, "
         f"{st.get('failedInvoiceCount')} failed")

    rows: list[dict[str, str]] = []
    token: str | None = None
    while True:
        page = client.request("GET", f"/sessions/{sref}/invoices", bearer=bearer,
                              params={"pageSize": 1000},
                              extra_headers={"x-continuation-token": token} if token else None)
        for inv in page.get("invoices") or []:
            if int(inv["status"]["code"]) != 200 or not inv.get("ksefNumber"):
                continue
            spec, number = by_hash[inv["invoiceHash"]]
            rows.append(_row(spec, number, inv["ksefNumber"], issue, sref,
                             inv["referenceNumber"], label))
        token = page.get("continuationToken") or None
        if not token:
            break
    bad = [r["ksef_number"] for r in rows if not is_valid_ksef(r["ksef_number"])]
    echo(f"  {len(rows)} KSeF numbers collected, {len(bad)} failing the offline CRC check")
    if len(rows) != len(specs):
        raise RuntimeError(f"only {len(rows)} of {len(specs)} invoices got a KSeF number")
    return rows


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--count", type=int, default=5, help="seller->buyer invoices")
    ap.add_argument("--batch", action="store_true",
                    help="send via a batch session (use for large counts); no reverse invoice")
    args = ap.parse_args()

    buyer = load_context(Env.TEST, "buyer")
    if buyer is None:
        raise SystemExit("run `ksefiz testenv-setup --profile buyer` first")
    issue = date.today() - timedelta(days=1)  # P_1 must not be after KSeF acceptance (UTC)
    log = AuditLog("test", "send-test-invoices")
    rows: list[dict[str, str]] = []
    with KsefClient(Env.TEST, on_request=log.on_request) as client:
        seller, s_sess = testenv.seal_login(client, "default")
        try:
            nets = [Decimal("1000.00"), Decimal("250.50"), Decimal("12345.67"),
                    Decimal("99.99"), Decimal("5000.00"), Decimal("777.77"), Decimal("42.00")]
            specs = [Spec(seller.nip, buyer.nip, nets[i % len(nets)], f"Usluga testowa {i + 1}")
                     for i in range(args.count)]
            if args.batch:
                rows += send_batch(client, s_sess, specs, issue, f"batch{args.count}")
            else:
                rows += send(client, s_sess, specs, issue)
        finally:
            auth.close(client, s_sess)
        if args.batch:
            _finish(rows, log)
            return
        _, b_sess = testenv.seal_login(client, "buyer")
        try:
            rows += send(client, b_sess, [Spec(buyer.nip, seller.nip, Decimal("300.00"),
                                               "Usluga testowa odwrotna")], issue)
        finally:
            auth.close(client, b_sess)

    _finish(rows, log)


def _finish(rows: list[dict[str, str]], log: AuditLog) -> None:
    out = state_dir(Env.TEST) / "invoices.csv"
    new = len(rows)
    if out.exists():
        rows = read_rows(out) + rows
    write_rows(out, OUT_FIELDS, rows)
    log.write(result="ok", invoices=new)
    print(f"{new} new, {len(rows)} TEST invoices listed in {out}")


if __name__ == "__main__":
    main()
