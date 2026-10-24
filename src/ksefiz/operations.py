"""M3 decode, M5 generate, M6 reverse lookup, M7 period query - with paging, the shared
`collectiveIdentifier` rate-limit budget, token refresh and offline pre-validation.

Every function returns plain row dicts whose keys are the *_FIELDS lists below; the same
lists define the Excel sheet columns (scripts/build_workbook.py).
"""

from __future__ import annotations

import time
from collections import Counter, defaultdict
from collections.abc import Callable, Iterator, Sequence
from dataclasses import dataclass, field
from datetime import UTC, date, datetime, timedelta
from datetime import time as dtime
from pathlib import Path
from typing import Any
from urllib.parse import quote

from ksefiz import auth
from ksefiz.client import ApiError, KsefClient
from ksefiz.config import DECODE_BATCH_SIZE, DEFAULT_MAX_INVOICES_PER_IZ, QUERY_MAX_DAYS
from ksefiz.io_csv import Row, parse_amount
from ksefiz.iz import KIND_IZ, KIND_KSEF, check
from ksefiz.ratelimit import COLLECTIVE_DEFAULTS, RateLimiter

# pageSize maxima differ per endpoint (open-api.json)
PAGE_SIZE_DECODE = 500
PAGE_SIZE_LIST = 200
MAX_DESCRIPTION = 512

DECODE_FIELDS = ["iz", "ksef_number", "amount", "currency", "description", "details_hidden",
                 "status", "error_code", "timestamp"]
GENERATE_IN = ["seller_nip", "ksef_number", "amount", "currency", "description", "group_key"]
GENERATE_FIELDS = ["group_key", "iz", "invoice_count", "status", "error_code"]
LOOKUP_FIELDS = ["ksef_number", "iz", "created_in_current_context", "date_created", "status",
                 "error_code"]
QUERY_FIELDS = ["iz", "date_created", "invoice_count", "created_in_current_context"]

OK = "OK"


class Monitor:
    """Progress reporting and cooperative cancellation for long runs (Excel buttons).
    Operations call tick() after each unit of work and stop early once stopped() is true,
    returning the rows collected so far."""

    def tick(self, done: int, total: int) -> None:
        pass

    def stopped(self) -> bool:
        return False


_NO_MONITOR = Monitor()


def _now() -> str:
    return datetime.now(UTC).isoformat(timespec="seconds")


def _flag(v: object) -> str:
    return "true" if v else "false"


def _err(e: ApiError) -> str:
    return ",".join(str(c) for c in e.codes) or str(e.status)


# ---------------------------------------------------------------- session


class Api:
    """Authenticated, throttled access to the collective-identifier endpoints."""

    def __init__(self, client: KsefClient, nip: str, ksef_token: str,
                 limiter_state: Path | None = None,
                 sleep: Callable[[float], None] = time.sleep) -> None:
        self.client = client
        self._nip = nip
        self._token = ksef_token
        self._sleep = sleep
        self._limiter_state = limiter_state
        self.session: auth.Session | None = None
        self.limiter = RateLimiter(COLLECTIVE_DEFAULTS, state_path=limiter_state, sleep=sleep)
        self.max_invoices = DEFAULT_MAX_INVOICES_PER_IZ

    def __enter__(self) -> Api:
        self.session = auth.login_with_token(self.client, self._nip, self._token, self._sleep)
        bearer = self.session.access_token
        try:  # effective values for this context; keep defaults if not readable
            rl = self.client.request("GET", "/rate-limits", bearer=bearer).get(
                "collectiveIdentifier")
            if rl:
                self.limiter = RateLimiter(rl, state_path=self._limiter_state, sleep=self._sleep)
            lim = self.client.request("GET", "/limits/context", bearer=bearer)
            self.max_invoices = int(lim["collectiveIdentifier"]["maxInvoices"])
        except (ApiError, KeyError, TypeError, ValueError):
            pass
        return self

    def __exit__(self, *exc: object) -> None:
        if self.session:
            auth.close(self.client, self.session)

    def call(self, method: str, path: str, *, json: Any = None,
             continuation: str | None = None, page_size: int | None = None) -> Any:
        assert self.session is not None, "use Api as a context manager"
        self.session = auth.ensure_fresh(self.client, self.session)
        self.limiter.acquire()
        return self.client.request(
            method, path, json=json, bearer=self.session.access_token,
            params={"pageSize": page_size} if page_size else None,
            extra_headers={"x-continuation-token": continuation} if continuation else None)

    def pages(self, method: str, path: str, key: str, json: Any = None,
              page_size: int = PAGE_SIZE_LIST) -> Iterator[dict[str, Any]]:
        """All items under `key` across continuation pages."""
        token: str | None = None
        while True:
            body = self.call(method, path, json=json, continuation=token, page_size=page_size)
            yield from body.get(key) or []
            token = body.get("continuationToken") or None
            if not token:
                return


# ---------------------------------------------------------------- M5 generate


@dataclass
class GenGroup:
    key: str
    seller_nip: str
    invoices: list[dict[str, Any]] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.errors


def prevalidate(rows: Sequence[Row], max_invoices: int = DEFAULT_MAX_INVOICES_PER_IZ
                ) -> list[GenGroup]:
    """Group invoices by group_key (default: seller NIP) and run every check the API would
    otherwise fail on: one seller per IZ (71004), no duplicates (71005), 2..max invoices,
    valid KSeF numbers, payment = amount + currency together, description length."""
    groups: dict[str, GenGroup] = {}
    seen_sellers: dict[str, set[str]] = defaultdict(set)
    numbers: dict[str, list[str]] = defaultdict(list)
    for i, r in enumerate(rows, start=2):
        c = check(r.get("ksef_number", ""))
        seller = (r.get("seller_nip", "").strip() or (c.normalised[:10] if c.normalised else ""))
        key = r.get("group_key", "").strip() or seller
        g = groups.setdefault(key, GenGroup(key, seller))
        where = f"row {i}"
        if not c.valid or c.kind != KIND_KSEF:
            why = "is an IZ" if c.kind == KIND_IZ else c.reason
            g.errors.append(f"{where}: invalid KSeF number ({why})")
            continue
        seen_sellers[key].add(seller)
        if c.normalised[:10] != seller:
            g.errors.append(f"{where}: seller_nip {seller} differs from KSeF number prefix")
        numbers[key].append(c.normalised)
        inv: dict[str, Any] = {"ksefNumber": c.normalised}
        amount_raw, cur = r.get("amount", "").strip(), r.get("currency", "").strip().upper()
        if amount_raw or cur:
            amount = parse_amount(amount_raw)
            if amount is None or len(cur) != 3 or not cur.isalpha():
                g.errors.append(f"{where}: payment needs a numeric amount and a 3-letter currency")
            else:
                inv["payment"] = {"amount": amount, "currency": cur}
        desc = r.get("description", "").strip()
        if len(desc) > MAX_DESCRIPTION:
            g.errors.append(f"{where}: description longer than {MAX_DESCRIPTION} characters")
        elif desc:
            inv["description"] = desc
        g.invoices.append(inv)

    for key, g in groups.items():
        if len(seen_sellers[key]) > 1:
            g.errors.append(f"mixed sellers {sorted(seen_sellers[key])} (KSeF error 71004)")
        dups = [n for n, k in Counter(numbers[key]).items() if k > 1]
        if dups:
            g.errors.append(f"duplicate KSeF numbers {dups} (KSeF error 71005)")
        n = len(g.invoices)
        if n < 2:
            g.errors.append(f"{n} invoice(s); an IZ needs at least 2")
        elif n > max_invoices:
            g.errors.append(f"{n} invoices; the context limit is {max_invoices} per IZ")
    return list(groups.values())


def generate(api: Api, groups: Sequence[GenGroup], monitor: Monitor = _NO_MONITOR) -> list[Row]:
    out: list[Row] = []
    for n, g in enumerate(groups, start=1):
        if monitor.stopped():
            break
        row = {"group_key": g.key, "iz": "", "invoice_count": str(len(g.invoices)),
               "status": "", "error_code": ""}
        if not g.ok:
            row["status"] = "rejected before API: " + "; ".join(g.errors)
        else:
            try:
                r = api.call("POST", "/collective-identifiers", json={"invoices": g.invoices})
                row["iz"], row["status"] = r["collectiveIdentifierNumber"], OK
            except ApiError as e:
                row["status"], row["error_code"] = f"API error: {e.detail or e.title}", _err(e)
        out.append(row)
        monitor.tick(n, len(groups))
    return out


# ---------------------------------------------------------------- M3 decode


def decode(api: Api, izs: Sequence[str], monitor: Monitor = _NO_MONITOR) -> list[Row]:
    """One row per invoice; IZ that fail offline validation never reach the API."""
    out: list[Row] = []
    valid: list[str] = []
    for raw in izs:
        c = check(raw)
        if c.valid and c.kind == KIND_IZ:
            if c.normalised not in valid:
                valid.append(c.normalised)
        elif raw.strip():
            out.append(_decode_row(raw.strip(), status=f"invalid IZ ({c.reason})"))
    starts = range(0, len(valid), DECODE_BATCH_SIZE)
    for n, i in enumerate(starts, start=1):
        if monitor.stopped():
            break
        batch = valid[i:i + DECODE_BATCH_SIZE]
        try:
            items = list(api.pages("POST", "/collective-identifiers/invoices", "invoices",
                                   json={"collectiveIdentifierNumbers": batch},
                                   page_size=PAGE_SIZE_DECODE))
        except ApiError as e:
            out += [_decode_row(iz, status=f"API error: {e.detail or e.title}",
                                error_code=_err(e)) for iz in batch]
            monitor.tick(n, len(starts))
            continue
        found = set()
        for it in items:
            found.add(it["collectiveIdentifierNumber"])
            pay = it.get("payment") or {}
            out.append(_decode_row(
                it["collectiveIdentifierNumber"], it["ksefNumber"],
                "" if pay.get("amount") is None else f"{float(pay['amount']):.2f}",
                pay.get("currency") or "", it.get("description") or "",
                "true" if it.get("detailsHidden") else "false", OK))
        out += [_decode_row(iz, status="no invoices returned") for iz in batch if iz not in found]
        monitor.tick(n, len(starts))
    return out


def _decode_row(iz: str, ksef: str = "", amount: str = "", currency: str = "",
                description: str = "", hidden: str = "", status: str = "",
                error_code: str = "") -> Row:
    return {"iz": iz, "ksef_number": ksef, "amount": amount, "currency": currency,
            "description": description, "details_hidden": hidden, "status": status,
            "error_code": error_code, "timestamp": _now()}


# ---------------------------------------------------------------- M6 lookup


def lookup(api: Api, ksef_numbers: Sequence[str], monitor: Monitor = _NO_MONITOR) -> list[Row]:
    out: list[Row] = []
    for n, raw in enumerate(ksef_numbers, start=1):
        if monitor.stopped():
            break
        monitor.tick(n - 1, len(ksef_numbers))
        c = check(raw)
        base = {"ksef_number": c.normalised or raw.strip(), "iz": "",
                "created_in_current_context": "", "date_created": "", "error_code": ""}
        if not raw.strip():
            continue
        if not c.valid or c.kind != KIND_KSEF:
            out.append({**base, "status": f"invalid KSeF number ({c.reason})"})
            continue
        try:
            items = list(api.pages("GET", f"/collective-identifiers/ksef/{quote(c.normalised)}",
                                   "collectiveIdentifiers"))
        except ApiError as e:
            out.append({**base, "status": f"API error: {e.detail or e.title}",
                        "error_code": _err(e)})
            continue
        if not items:
            out.append({**base, "status": "not in any IZ"})
        for it in items:
            out.append({**base, "iz": it["collectiveIdentifierNumber"],
                        "created_in_current_context": _flag(it["createdInCurrentContext"]),
                        "date_created": it["dateCreated"], "status": OK})
    return out


# ---------------------------------------------------------------- M7 query


def windows(date_from: date, date_to: date, days: int = QUERY_MAX_DAYS
            ) -> list[tuple[datetime, datetime]]:
    """Split [date_from, date_to] (whole days, UTC) into API windows of at most `days`."""
    if date_to < date_from:
        raise ValueError("date_to is before date_from")
    start = datetime.combine(date_from, dtime.min, UTC)
    end = datetime.combine(date_to + timedelta(days=1), dtime.min, UTC)
    out = []
    while start < end:
        stop = min(start + timedelta(days=days), end)
        out.append((start, stop - timedelta(milliseconds=1)))
        start = stop
    return out


def query(api: Api, date_from: date, date_to: date, *,
          created_in_current_context: bool | None = None,
          monitor: Monitor = _NO_MONITOR) -> list[Row]:
    out: list[Row] = []
    seen: set[str] = set()
    wins = windows(date_from, date_to)
    for n, (a, b) in enumerate(wins, start=1):
        if monitor.stopped():
            break
        monitor.tick(n - 1, len(wins))
        body: dict[str, Any] = {"dateCreatedFrom": a.isoformat(timespec="milliseconds"),
                                "dateCreatedTo": b.isoformat(timespec="milliseconds")}
        if created_in_current_context is not None:
            body["createdInCurrentContext"] = created_in_current_context
        for it in api.pages("POST", "/collective-identifiers/query", "collectiveIdentifiers",
                            json=body):
            n = it["collectiveIdentifierNumber"]
            if n in seen:
                continue
            seen.add(n)
            out.append({"iz": n, "date_created": it["dateCreated"],
                        "invoice_count": str(it["invoiceCount"]),
                        "created_in_current_context": _flag(it["createdInCurrentContext"])})
    return out
