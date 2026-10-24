import json
import random
from datetime import UTC, date, datetime, timedelta
from typing import Any

import httpx
import pytest
import respx

from ksefiz import operations as ops
from ksefiz.client import KsefClient
from ksefiz.config import BASE_URLS, Env
from ksefiz.ratelimit import RateLimiter
from ksefiz.synthetic import fake_nip, make_iz, make_ksef

BASE = BASE_URLS[Env.TEST]
rng = random.Random(7)
SELLER, OTHER = fake_nip(rng), fake_nip(rng)
K = [make_ksef(SELLER, "20261024", f"{i:012X}") for i in range(6)]
KO = make_ksef(OTHER, "20261024", "0" * 12)


def _row(k: str, group: str = "", amount: str = "", cur: str = "", desc: str = "",
         seller: str = "") -> dict[str, str]:
    return {"seller_nip": seller, "ksef_number": k, "amount": amount, "currency": cur,
            "description": desc, "group_key": group}


@pytest.fixture
def api() -> ops.Api:
    """An Api with a fake session and an unthrottled limiter (no login call)."""
    from ksefiz.auth import Session

    client = KsefClient(Env.TEST, sleep=lambda s: None)
    a = ops.Api(client, SELLER, "tok", sleep=lambda s: None)
    future = datetime.now(UTC) + timedelta(hours=1)
    a.session = Session("ACCESS", future, "REFRESH", future, "REF")
    a.limiter = RateLimiter({"perSecond": 10**6, "perMinute": 10**6, "perHour": 10**6})
    return a


# ------------------------------------------------------------ prevalidate


def test_prevalidate_groups_by_seller_by_default() -> None:
    ko2 = make_ksef(OTHER, "20261024", "1" * 12)
    groups = ops.prevalidate([_row(K[0]), _row(K[1]), _row(KO), _row(ko2)])
    assert [(g.key, len(g.invoices), g.ok) for g in groups] == [
        (SELLER, 2, True), (OTHER, 2, True)]


def test_prevalidate_rejects_what_the_api_would() -> None:
    rows = [_row(K[0], "MIX"), _row(KO, "MIX"),
            _row(K[1], "DUP"), _row(K[1].replace("-", ""), "DUP"),
            _row(K[2], "ONE"),
            _row(K[3], "BADPAY", amount="10"), _row(K[4], "BADPAY", amount="x", cur="PLN"),
            _row(K[0], "LONG", desc="x" * 513), _row(K[1], "LONG"),
            _row(K[0][:-2] + "00", "CRC"), _row(K[1], "CRC"),
            _row(K[0], "NIP", seller=OTHER), _row(K[1], "NIP")]
    errs = {g.key: " ".join(g.errors) for g in ops.prevalidate(rows)}
    assert "71004" in errs["MIX"]
    assert "71005" in errs["DUP"]
    assert "at least 2" in errs["ONE"]
    assert errs["BADPAY"].count("payment needs") == 2
    assert "512" in errs["LONG"]
    assert "crc" in errs["CRC"]
    assert "differs from KSeF number prefix" in errs["NIP"]


def test_prevalidate_limit_and_payload() -> None:
    rows = [_row(make_ksef(SELLER, "20261024", f"{i:012X}"), amount="1 234,50", cur="pln",
                 desc="ok") for i in range(4)]
    (g,) = ops.prevalidate(rows, max_invoices=3)
    assert "limit is 3" in g.errors[0]
    (g,) = ops.prevalidate(rows[:3], max_invoices=3)
    assert g.ok and g.invoices[0] == {"ksefNumber": rows[0]["ksef_number"],
                                      "payment": {"amount": 1234.5, "currency": "PLN"},
                                      "description": "ok"}


# ------------------------------------------------------------ generate


@respx.mock
def test_generate_skips_rejected_groups_and_reports_api_errors(api: ops.Api) -> None:
    route = respx.post(f"{BASE}/collective-identifiers").mock(side_effect=[
        httpx.Response(201, json={"collectiveIdentifierNumber": "IZ-1"}),
        httpx.Response(400, json={"title": "Bad Request", "detail": "x", "errors": [
            {"code": 71004, "description": "Różni sprzedawcy"}]}),
    ])
    groups = ops.prevalidate([_row(K[0], "A"), _row(K[1], "A"), _row(K[2], "B"),
                              _row(K[3], "C"), _row(K[4], "C")])
    rows = ops.generate(api, groups)
    assert [r["status"][:8] for r in rows] == ["OK", "rejected", "API erro"]
    assert rows[0]["iz"] == "IZ-1" and rows[2]["error_code"] == "71004"
    assert route.call_count == 2
    assert json.loads(route.calls[0].request.content)["invoices"][0] == {"ksefNumber": K[0]}


# ------------------------------------------------------------ decode


@respx.mock
def test_decode_batches_of_ten_with_paging(api: ops.Api) -> None:
    izs = [make_iz(SELLER, "202610", f"{i:012X}") for i in range(23)]
    calls: list[dict[str, Any]] = []

    def handler(req: httpx.Request) -> httpx.Response:
        body = json.loads(req.content)
        tok = req.headers.get("x-continuation-token")
        calls.append({"n": len(body["collectiveIdentifierNumbers"]), "tok": tok,
                      "page": req.url.params.get("pageSize")})
        batch = body["collectiveIdentifierNumbers"]
        half = len(batch) // 2
        part = batch[half:] if tok else batch[:half]
        items = [{"collectiveIdentifierNumber": iz, "ksefNumber": K[0], "detailsHidden": False,
                  "payment": {"amount": 10, "currency": "PLN"}} for iz in part if iz != izs[5]]
        return httpx.Response(200, json={"invoices": items,
                                         "continuationToken": None if tok else "NEXT"})

    respx.post(f"{BASE}/collective-identifiers/invoices").mock(side_effect=handler)
    bad = izs[0][:-2] + "00"
    rows = ops.decode(api, [*izs, izs[0].lower(), bad, ""])
    assert [c["n"] for c in calls] == [10, 10, 10, 10, 3, 3]
    assert [c["tok"] for c in calls] == [None, "NEXT"] * 3
    assert {c["page"] for c in calls} == {str(ops.PAGE_SIZE_DECODE)}
    by_status = {r["status"] for r in rows}
    assert {"OK", "invalid IZ (crc)", "no invoices returned"} <= by_status
    assert sum(r["status"] == "OK" for r in rows) == 22
    assert next(r for r in rows if r["iz"] == izs[5])["status"] == "no invoices returned"
    assert next(r for r in rows if r["status"] == "OK")["amount"] == "10.00"


@respx.mock
def test_decode_api_error_marks_whole_batch(api: ops.Api) -> None:
    respx.post(f"{BASE}/collective-identifiers/invoices").respond(
        400, json={"title": "Bad Request", "errors": [{"code": 21418, "description": "token"}]})
    izs = [make_iz(SELLER, "202610", f"{i:012X}") for i in range(3)]
    rows = ops.decode(api, izs)
    assert [r["error_code"] for r in rows] == ["21418"] * 3


# ------------------------------------------------------------ lookup / query


@respx.mock
def test_lookup_pages_and_marks_empty(api: ops.Api) -> None:
    def handler(req: httpx.Request) -> httpx.Response:
        if K[1] in str(req.url):
            return httpx.Response(200, json={"collectiveIdentifiers": []})
        tok = req.headers.get("x-continuation-token")
        n = "B" if tok else "A"
        return httpx.Response(200, json={
            "collectiveIdentifiers": [{"collectiveIdentifierNumber": n,
                                       "createdInCurrentContext": True,
                                       "dateCreated": "2026-10-24T00:00:00Z"}],
            "continuationToken": None if tok else "T"})

    respx.get(url__startswith=f"{BASE}/collective-identifiers/ksef/").mock(side_effect=handler)
    rows = ops.lookup(api, [K[0], K[1], "junk"])
    assert [(r["iz"], r["status"]) for r in rows] == [
        ("A", "OK"), ("B", "OK"), ("", "not in any IZ"), ("", "invalid KSeF number (format)")]


def test_windows_split_into_100_days() -> None:
    w = ops.windows(date(2026, 1, 1), date(2026, 12, 31))
    assert len(w) == 4
    assert all((b - a) <= timedelta(days=100) for a, b in w)
    assert w[0][0] == datetime(2026, 1, 1, tzinfo=UTC)
    assert w[-1][1] == datetime(2027, 1, 1, tzinfo=UTC) - timedelta(milliseconds=1)
    for (_, b), (a2, _) in zip(w, w[1:], strict=False):
        assert a2 - b == timedelta(milliseconds=1)
    with pytest.raises(ValueError):
        ops.windows(date(2026, 2, 1), date(2026, 1, 1))


@respx.mock
def test_query_dedupes_across_windows(api: ops.Api) -> None:
    route = respx.post(f"{BASE}/collective-identifiers/query").respond(json={
        "collectiveIdentifiers": [{"collectiveIdentifierNumber": "IZ-X", "dateCreated": "d",
                                   "invoiceCount": 2, "createdInCurrentContext": False}]})
    rows = ops.query(api, date(2026, 1, 1), date(2026, 6, 30), created_in_current_context=True)
    assert route.call_count == 2 and len(rows) == 1
    body = json.loads(route.calls[0].request.content)
    assert body["createdInCurrentContext"] is True
    assert body["dateCreatedFrom"].startswith("2026-01-01")
