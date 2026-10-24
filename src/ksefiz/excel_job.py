"""File-based job protocol between the Excel buttons (excel/vba/modHelper.bas) and ksefiz.

The macro creates a job folder and starts `ksefiz excel-run --job DIR --op OP ...` in the
background. Inside DIR:

    in.xlsx        written by Excel: header row + values (decode/generate/lookup)
    progress.txt   written by ksefiz: "done total message" (ASCII), updated as it runs
    cancel         created by Excel when the user presses Esc; checked between API calls
    out.xlsx       written by ksefiz: sheet "out" (result rows) + sheet "_result" (key/value)
    exit           written by the launcher shell after ksefiz exits (exit code)

Results are written even on errors or cancellation, so the macro always has something to show.
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

from ksefiz import auth, operations
from ksefiz.audit import AuditLog
from ksefiz.client import ApiError, KsefClient
from ksefiz.config import Env
from ksefiz.credentials import get_token, load_context, state_dir
from ksefiz.io_csv import Row, read_xlsx_rows, write_xlsx

OPS = ("decode", "generate", "check", "lookup", "query")
RESULT_FIELDS = ["key", "value"]


class JobMonitor(operations.Monitor):
    def __init__(self, job: Path, label: str) -> None:
        self.job = job
        self.label = label

    def tick(self, done: int, total: int) -> None:
        progress(self.job, done, total, self.label)

    def stopped(self) -> bool:
        return (self.job / "cancel").exists()


def progress(job: Path, done: int, total: int, message: str) -> None:
    text = f"{done} {total} {message}".encode("ascii", "replace")
    tmp = job / "progress.tmp"
    tmp.write_bytes(text)
    tmp.replace(job / "progress.txt")


@dataclass
class JobResult:
    status: str = "ok"  # ok | cancelled | error
    message: str = ""
    error_code: str = ""
    trace_id: str = ""
    requests: int = 0
    input_hash: str = ""
    budget_minute: str = ""
    budget_hour: str = ""
    nip: str = ""

    def rows(self) -> list[dict[str, Any]]:
        return [{"key": k, "value": v} for k, v in vars(self).items()]


def _column(rows: list[Row], name: str) -> list[str]:
    return [r.get(name, "") if name in r else next(iter(r.values()), "") for r in rows]


def default_log_dir() -> Path:
    return state_dir(Env.TEST).parent / "logs"


def run(job: Path, op: str, env: Env, profile: str, *, date_from: str = "", date_to: str = "",
        own_only: bool = False, log_dir: Path | None = None) -> JobResult:
    """Run one Excel job; never raises for API/auth problems (they go into out.xlsx)."""
    res = JobResult()
    fields: list[str] = []
    rows: list[Row] = []
    in_path = job / "in.xlsx"
    log = AuditLog(env.value, f"excel-{op}", (log_dir or default_log_dir()) / "audit.jsonl")
    try:
        progress(job, 0, 0, f"{op}: starting")
        rows_in = read_xlsx_rows(in_path) if in_path.exists() else []
        if in_path.exists():
            res.input_hash = hashlib.sha256(in_path.read_bytes()).hexdigest()[:16]
        monitor = JobMonitor(job, op)

        if op == "check":
            fields = operations.GENERATE_FIELDS
            rows = [{"group_key": g.key, "iz": "", "invoice_count": str(len(g.invoices)),
                     "status": "ready" if g.ok else "rejected before API: " + "; ".join(g.errors),
                     "error_code": ""} for g in operations.prevalidate(rows_in)]
        else:
            ctx = load_context(env, profile)
            token = get_token(env, ctx.nip) if ctx else None
            if not ctx or not token:
                raise RuntimeError(f"no {env.value.upper()} context/token for profile {profile!r}")
            res.nip = ctx.nip
            with KsefClient(env, on_request=log.on_request) as client, operations.Api(
                    client, ctx.nip, token, state_dir(env) / f"ratelimit.{ctx.nip}.json") as api:
                progress(job, 0, 0, f"{op}: logged in")
                if op == "decode":
                    fields = operations.DECODE_FIELDS
                    rows = operations.decode(api, _column(rows_in, "iz"), monitor)
                elif op == "generate":
                    fields = operations.GENERATE_FIELDS
                    groups = operations.prevalidate(rows_in, api.max_invoices)
                    rows = operations.generate(api, groups, monitor)
                elif op == "lookup":
                    fields = operations.LOOKUP_FIELDS
                    rows = operations.lookup(api, _column(rows_in, "ksef_number"), monitor)
                elif op == "query":
                    fields = operations.QUERY_FIELDS
                    rows = operations.query(api, date.fromisoformat(date_from),
                                            date.fromisoformat(date_to),
                                            created_in_current_context=True if own_only else None,
                                            monitor=monitor)
                else:
                    raise ValueError(f"unknown operation {op!r}")
                b = api.limiter.remaining()
                res.budget_minute, res.budget_hour = str(b.per_minute), str(b.per_hour)
            if monitor.stopped():
                res.status, res.message = "cancelled", f"cancelled after {len(rows)} rows"
        if not res.message:
            res.message = f"{len(rows)} rows"
    except ApiError as e:
        res.status, res.message = "error", str(e)
        res.error_code = ",".join(map(str, e.codes)) or str(e.status)
        res.trace_id = e.trace_id
    except (auth.AuthError, RuntimeError, ValueError, OSError) as e:
        res.status, res.message = "error", str(e)
    res.requests = log.requests
    log.write(result=res.status, message=res.message, requests=res.requests,
              input_hash=res.input_hash, error_code=res.error_code, trace_id=res.trace_id)
    write_xlsx(job / "out.xlsx", {"out": (fields, rows), "_result": (RESULT_FIELDS, res.rows())})
    progress(job, 1, 1, f"{op}: {res.status}")
    return res
