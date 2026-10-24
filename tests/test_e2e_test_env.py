"""Runs against the real KSeF TEST environment. Needs `ksefiz testenv-setup` (default and
buyer profiles) and `scripts/send_test_invoices.py` first.

    .venv/bin/pytest -m e2e
    KSEFIZ_E2E_HEAVY=1 .venv/bin/pytest -m e2e   # + generate/decode 50 IZ (uses 20 of the
                                                # 132 IZ slots per test invoice each run)
    KSEFIZ_E2E_500=1 .venv/bin/pytest -m e2e     # + one IZ over 500 invoices (needs
                                                # send_test_invoices.py --batch --count 500)
"""

import csv
import itertools
import os
from collections import Counter
from collections.abc import Iterator

import pytest

from ksefiz import auth
from ksefiz import operations as ops
from ksefiz.client import ApiError, KsefClient
from ksefiz.config import Env
from ksefiz.credentials import get_token, load_context, state_dir

pytestmark = pytest.mark.e2e


def _creds(profile: str = "default") -> tuple[str, str]:
    ctx = load_context(Env.TEST, profile)
    token = get_token(Env.TEST, ctx.nip) if ctx else None
    if not ctx or not token:
        pytest.skip(f"no TEST profile {profile!r}; run `ksefiz testenv-setup`")
    return ctx.nip, token


@pytest.fixture(scope="module")
def invoices() -> list[dict[str, str]]:
    path = state_dir(Env.TEST) / "invoices.csv"
    if not path.exists():
        pytest.skip("no TEST invoices; run scripts/send_test_invoices.py")
    with path.open(encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


@pytest.fixture(scope="module")
def seller_api() -> Iterator[ops.Api]:
    nip, token = _creds()
    with KsefClient(Env.TEST) as c, ops.Api(c, nip, token,
                                            state_dir(Env.TEST) / f"ratelimit.{nip}.json") as a:
        yield a


def test_token_login_refresh_limits_close() -> None:
    nip, token = _creds()
    with KsefClient(Env.TEST) as c:
        s = auth.login_with_token(c, nip, token)
        auth.refresh(c, s)
        limits = c.request("GET", "/limits/context", bearer=s.access_token)
        assert limits["collectiveIdentifier"]["maxInvoices"] >= 2
        auth.close(c, s)


def test_api_itself_rejects_mixed_sellers_and_duplicates(
        seller_api: ops.Api, invoices: list[dict[str, str]]) -> None:
    """Confirms the error codes our pre-validation claims to prevent."""
    sellers = {r["seller_nip"]: r["ksef_number"] for r in invoices}
    assert len(sellers) >= 2, "need invoices from two sellers"
    mixed = [{"ksefNumber": k} for k in list(sellers.values())[:2]]
    with pytest.raises(ApiError) as e:
        seller_api.call("POST", "/collective-identifiers", json={"invoices": mixed})
    assert 71004 in e.value.codes
    k = invoices[0]["ksef_number"]
    with pytest.raises(ApiError) as e:
        seller_api.call("POST", "/collective-identifiers",
                        json={"invoices": [{"ksefNumber": k}, {"ksefNumber": k}]})
    assert 71005 in e.value.codes


def test_lookup_finds_every_iz_of_an_invoice(seller_api: ops.Api,
                                             invoices: list[dict[str, str]]) -> None:
    own = [r for r in invoices if r["seller_nip"] == seller_api._nip][:2]
    rows = ops.generate(seller_api, ops.prevalidate(
        [{"ksef_number": r["ksef_number"], "group_key": g} for g in ("L1", "L2") for r in own]))
    created = {r["iz"] for r in rows}
    assert all(r["status"] == ops.OK for r in rows)
    found = {r["iz"] for r in ops.lookup(seller_api, [own[0]["ksef_number"]])}
    assert created <= found


@pytest.mark.skipif(not os.environ.get("KSEFIZ_E2E_HEAVY"), reason="set KSEFIZ_E2E_HEAVY=1")
def test_generate_50_and_decode_with_paging(seller_api: ops.Api, invoices: list[dict[str, str]],
                                            monkeypatch: pytest.MonkeyPatch) -> None:
    own = [r for r in invoices if r["seller_nip"] == seller_api._nip]
    pairs = list(itertools.combinations(own[:5], 2))  # 10 distinct pairs
    rows_in = [{"ksef_number": r["ksef_number"], "amount": r["gross"], "currency": "PLN",
                "description": f"e2e {i}", "group_key": f"G{i:02d}"}
               for i in range(50) for r in pairs[i % len(pairs)]]
    generated = ops.generate(seller_api, ops.prevalidate(rows_in, seller_api.max_invoices))
    izs = [r["iz"] for r in generated]
    assert len(izs) == 50 and all(r["status"] == ops.OK for r in generated)

    monkeypatch.setattr(ops, "PAGE_SIZE_DECODE", 10)  # 10 IZ x 2 invoices -> 2 pages per batch
    decoded = ops.decode(seller_api, izs)
    assert all(r["status"] == ops.OK for r in decoded)
    assert Counter(r["iz"] for r in decoded) == dict.fromkeys(izs, 2)
    assert all(r["details_hidden"] == "false" and r["amount"] for r in decoded)
    assert seller_api.limiter.remaining().per_hour >= 0


@pytest.mark.skipif(not os.environ.get("KSEFIZ_E2E_500"), reason="set KSEFIZ_E2E_500=1")
def test_generate_and_decode_500_invoice_iz(seller_api: ops.Api,
                                            invoices: list[dict[str, str]]) -> None:
    big = [r for r in invoices if r.get("set") == "batch500"]
    if len(big) < 500:
        pytest.skip("run scripts/send_test_invoices.py --batch --count 500")
    rows = [{"ksef_number": r["ksef_number"], "amount": r["gross"], "currency": "PLN",
             "group_key": "BIG"} for r in big[:500]]
    over = rows + [{**rows[0], "ksef_number": next(
        r["ksef_number"] for r in invoices if r.get("set") != "batch500"
        and r["seller_nip"] == seller_api._nip)}]
    (rejected,) = ops.prevalidate(over, seller_api.max_invoices)
    assert not rejected.ok and "limit is 500" in " ".join(rejected.errors)

    (gen,) = ops.generate(seller_api, ops.prevalidate(rows, seller_api.max_invoices))
    assert gen["status"] == ops.OK and gen["invoice_count"] == "500"
    decoded = ops.decode(seller_api, [gen["iz"]])
    assert len(decoded) == 500 and all(r["status"] == ops.OK for r in decoded)
    assert {r["ksef_number"] for r in decoded} == {r["ksef_number"] for r in rows}
