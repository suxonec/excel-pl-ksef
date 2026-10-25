"""ksefiz command line. detect / validate / match are offline; the rest call KSeF."""

from __future__ import annotations

import getpass
import hashlib
import sys
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from datetime import date
from pathlib import Path
from typing import Annotated

import typer

from ksefiz import __version__, auth, excel_job, operations, synthetic, testenv, userconfig
from ksefiz.audit import AuditLog
from ksefiz.client import ApiError, KsefClient
from ksefiz.config import DEFAULT_ENV, Env
from ksefiz.credentials import DEFAULT_PROFILE, get_token, load_context, state_dir
from ksefiz.io_csv import read_rows, write_rows
from ksefiz.iz import check, iter_detect
from ksefiz.match import MATCH_FIELDS, match

app = typer.Typer(no_args_is_help=True, add_completion=False,
                  help="KSeF collective identifier (IZ) bulk helper.")

InPath = Annotated[Path, typer.Option("--in", exists=True, dir_okay=False, help="Input CSV")]
OutPath = Annotated[Path, typer.Option("--out", dir_okay=False, help="Output CSV")]
Delim = Annotated[str, typer.Option("--delimiter", help="Output delimiter")]
EnvOpt = Annotated[Env, typer.Option("--env", help="KSeF environment (PROD asks to confirm)")]
ProfileOpt = Annotated[str, typer.Option(help="Stored context profile, e.g. 'buyer' on TEST")]

DETECT_FIELDS = ["line_id", "kind", "raw", "normalised", "crc_ok", "valid", "reason"]
VALIDATE_FIELDS = ["input", "kind", "normalised", "crc_ok", "valid", "reason"]


def _yn(b: bool) -> str:
    return "Y" if b else "N"


@app.command()
def version() -> None:
    """Print the tool version."""
    typer.echo(__version__)


@app.command("detect")
def detect_cmd(
    in_path: InPath,
    out_path: OutPath,
    text_col: Annotated[str, typer.Option(help="Column to scan; blank = all columns")] = (
        "free_text"),
    delimiter: Delim = ",",
) -> None:
    """M1: find IZ and KSeF numbers in bank statement lines."""
    rows = read_rows(in_path)
    out = []
    without_valid = 0  # lines the workbook highlights red
    for i, r in enumerate(rows, start=1):
        text = r.get(text_col, "") if text_col and text_col in r else " ".join(r.values())
        found = list(iter_detect(text))
        without_valid += not any(d.check.valid for d in found)
        for d in found:
            out.append({
                "line_id": r.get("line_id", str(i)), "kind": d.check.kind,
                "raw": d.raw.replace("\r", " ").replace("\n", " "),
                "normalised": d.check.normalised, "crc_ok": _yn(d.check.crc_ok),
                "valid": _yn(d.check.valid), "reason": d.check.reason,
            })
    write_rows(out_path, DETECT_FIELDS, out, delimiter)
    valid = sum(o["valid"] == "Y" for o in out)
    typer.echo(f"{len(rows)} lines, {len(out)} candidates, {valid} valid, "
               f"{without_valid} lines without a valid ID -> {out_path}")


@app.command("validate")
def validate_cmd(
    in_path: InPath,
    out_path: OutPath,
    col: Annotated[str, typer.Option(help="Column with IDs; blank = first column")] = "",
    delimiter: Delim = ",",
) -> None:
    """M2: format and CRC-8 check of a list of IZ / KSeF numbers."""
    rows = read_rows(in_path)
    out = []
    for r in rows:
        value = r.get(col, "") if col else next(iter(r.values()), "")
        c = check(value)
        out.append({"input": value, "kind": c.kind, "normalised": c.normalised,
                    "crc_ok": _yn(c.crc_ok), "valid": _yn(c.valid), "reason": c.reason})
    write_rows(out_path, VALIDATE_FIELDS, out, delimiter)
    bad = sum(o["valid"] == "N" for o in out)
    typer.echo(f"{len(out)} checked, {bad} invalid -> {out_path}")


@app.command("match")
def match_cmd(
    decoded: Annotated[Path, typer.Option(
        exists=True, dir_okay=False, help="Decoded rows: iz, ksef_number, amount, currency")],
    open_items: Annotated[Path, typer.Option(
        exists=True, dir_okay=False, help="ERP extract: doc_no, customer, ksef_number, amount")],
    out_path: OutPath,
    delimiter: Delim = ",",
) -> None:
    """M4: join decoded KSeF numbers to ERP open items."""
    rows = match(read_rows(decoded), read_rows(open_items))
    write_rows(out_path, MATCH_FIELDS, (r.as_dict() for r in rows), delimiter)
    counts: dict[str, int] = {}
    for r in rows:
        counts[r.status] = counts.get(r.status, 0) + 1
    typer.echo(", ".join(f"{k}: {v}" for k, v in sorted(counts.items())) + f" -> {out_path}")


@app.command("gen-testdata")
def gen_testdata(
    out_dir: Annotated[Path, typer.Option(file_okay=False)] = Path("testdata"),
    seed: int = 20261024,
    lines: int = 1000,
) -> None:
    """Write synthetic statements, open items and decode results (no real data)."""
    synthetic.write(synthetic.generate(seed=seed, lines=lines), out_dir)
    typer.echo(f"synthetic data written to {out_dir}/")


def _client(env: Env, operation: str) -> tuple[KsefClient, AuditLog]:
    if env == Env.PROD:
        typer.confirm("PROD: real data with legal effect. Continue?", abort=True)
    log = AuditLog(env.value, operation)
    return KsefClient(env, on_request=log.on_request), log


def _fail(e: Exception) -> typer.Exit:
    typer.secho(f"error: {e}", fg=typer.colors.RED, err=True)
    return typer.Exit(1)


@app.command("testenv-setup")
def testenv_setup(
    nip: Annotated[str, typer.Option(help="Test NIP; default: random, checksum-invalid")] = "",
    force: Annotated[bool, typer.Option(help="Create a new context even if one exists")] = False,
    profile: ProfileOpt = DEFAULT_PROFILE,
) -> None:
    """Bootstrap KSeF TEST: self-signed seal, XAdES login, CollectiveIdentifierManage token."""
    existing = load_context(Env.TEST, profile)
    if existing and not force:
        typer.echo(f"TEST profile {profile!r} already set up for NIP {existing.nip} "
                   f"(token {existing.token_reference}); use --force for a new one")
        return
    client, log = _client(Env.TEST, "testenv-setup")
    with client:
        try:
            ctx = testenv.bootstrap(client, nip or testenv.random_test_nip(), echo=typer.echo,
                                    profile=profile)
        except (ApiError, auth.AuthError, RuntimeError) as e:
            log.write(result="error", error=str(e))
            raise _fail(e) from e
    log.write(result="ok", nip=ctx.nip, token_reference=ctx.token_reference)
    typer.echo(f"done: NIP {ctx.nip}, permissions {', '.join(ctx.permissions)}")


@app.command("auth-test")
def auth_test(env: EnvOpt = DEFAULT_ENV, profile: ProfileOpt = DEFAULT_PROFILE) -> None:
    """Log in with the stored KSeF token, refresh once, read context limits, log out."""
    ctx = load_context(env, profile)
    token = get_token(env, ctx.nip) if ctx else None
    if not ctx or not token:
        raise _fail(RuntimeError(f"no {env.value.upper()} context/token; "
                                 "run `ksefiz testenv-setup` first (TEST)"))
    client, log = _client(env, "auth-test")
    with client:
        try:
            session = auth.login_with_token(client, ctx.nip, token)
            typer.echo(f"token login OK for NIP {ctx.nip} on {env.value.upper()} "
                       f"(access token valid until {session.access_valid_until:%H:%M:%S} UTC)")
            auth.refresh(client, session)
            typer.echo("refresh OK")
            limits = client.request("GET", "/limits/context", bearer=session.access_token)
            iz = {k: v for k, v in limits.items() if "collective" in k.lower()}
            typer.echo(f"context limits: {iz or limits}")
            auth.close(client, session)
            typer.echo("session closed")
        except (ApiError, auth.AuthError) as e:
            log.write(result="error", error=str(e))
            raise _fail(e) from e
    log.write(result="ok", requests=log.requests)


def _first_col(rows: list[dict[str, str]], name: str) -> list[str]:
    return [r.get(name, "") if name in r else next(iter(r.values()), "") for r in rows]


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()[:16]


@contextmanager
def _api(env: Env, profile: str, operation: str, in_path: Path | None
         ) -> Iterator[tuple[operations.Api, AuditLog]]:
    ctx = load_context(env, profile)
    token = get_token(env, ctx.nip) if ctx else None
    if not ctx or not token:
        raise _fail(RuntimeError(f"no {env.value.upper()} context/token for profile "
                                 f"{profile!r}; run `ksefiz testenv-setup` (TEST)"))
    client, log = _client(env, operation)
    input_hash = _sha256(in_path) if in_path else ""
    try:
        with client, operations.Api(client, ctx.nip, token,
                                    state_dir(env) / f"ratelimit.{ctx.nip}.json") as api:
            yield api, log
            b = api.limiter.remaining()
            typer.echo(f"collectiveIdentifier budget left: {b.per_minute}/min, "
                       f"{b.per_hour}/hour")
    except (ApiError, auth.AuthError) as e:
        log.write(result="error", error=str(e), input_hash=input_hash)
        raise _fail(e) from e
    log.write(result="ok", requests=log.requests, input_hash=input_hash, nip=ctx.nip)


def _summary(rows: list[dict[str, str]], out_path: Path, ok: Callable[[dict[str, str]], bool]
             ) -> None:
    good = sum(1 for r in rows if ok(r))
    typer.echo(f"{len(rows)} rows ({good} OK, {len(rows) - good} not) -> {out_path}")


@app.command("decode")
def decode_cmd(in_path: InPath, out_path: OutPath, env: EnvOpt = DEFAULT_ENV,
               profile: ProfileOpt = DEFAULT_PROFILE,
               col: Annotated[str, typer.Option(help="Column with IZ")] = "iz",
               delimiter: Delim = ",") -> None:
    """M3: IZ list -> one row per invoice (KSeF number, payment where visible)."""
    izs = _first_col(read_rows(in_path), col)
    with _api(env, profile, "decode", in_path) as (api, _):
        rows = operations.decode(api, izs)
    write_rows(out_path, operations.DECODE_FIELDS, rows, delimiter)
    _summary(rows, out_path, lambda r: r["status"] == operations.OK)


@app.command("generate")
def generate_cmd(in_path: InPath, out_path: OutPath, env: EnvOpt = DEFAULT_ENV,
                 profile: ProfileOpt = DEFAULT_PROFILE,
                 dry_run: Annotated[bool, typer.Option(help="Pre-checks only, no API")] = False,
                 delimiter: Delim = ",") -> None:
    """M5: group invoices per seller / group_key, pre-validate, generate one IZ per group."""
    rows_in = read_rows(in_path)
    if dry_run:
        groups = operations.prevalidate(rows_in)
        rows = [{"group_key": g.key, "iz": "", "invoice_count": str(len(g.invoices)),
                 "status": "ready" if g.ok else "rejected before API: " + "; ".join(g.errors),
                 "error_code": ""} for g in groups]
    else:
        with _api(env, profile, "generate", in_path) as (api, _):
            rows = operations.generate(api, operations.prevalidate(rows_in, api.max_invoices))
    write_rows(out_path, operations.GENERATE_FIELDS, rows, delimiter)
    _summary(rows, out_path, lambda r: r["status"] in (operations.OK, "ready"))


@app.command("lookup")
def lookup_cmd(in_path: InPath, out_path: OutPath, env: EnvOpt = DEFAULT_ENV,
               profile: ProfileOpt = DEFAULT_PROFILE,
               col: Annotated[str, typer.Option(help="Column with KSeF numbers")] = "ksef_number",
               delimiter: Delim = ",") -> None:
    """M6: KSeF number -> every IZ that contains it."""
    ksefs = _first_col(read_rows(in_path), col)
    with _api(env, profile, "lookup", in_path) as (api, _):
        rows = operations.lookup(api, ksefs)
    write_rows(out_path, operations.LOOKUP_FIELDS, rows, delimiter)
    _summary(rows, out_path, lambda r: r["status"] == operations.OK)


@app.command("query")
def query_cmd(date_from: Annotated[str, typer.Option("--from", help="YYYY-MM-DD")],
              date_to: Annotated[str, typer.Option("--to", help="YYYY-MM-DD")],
              out_path: OutPath, env: EnvOpt = DEFAULT_ENV,
              profile: ProfileOpt = DEFAULT_PROFILE,
              own_only: Annotated[bool, typer.Option(help="Only IZ made in this context")] = False,
              delimiter: Delim = ",") -> None:
    """M7: IZ created in (or involving) this context in a date range (any length)."""
    d1, d2 = date.fromisoformat(date_from), date.fromisoformat(date_to)
    with _api(env, profile, "query", None) as (api, _):
        rows = operations.query(api, d1, d2, created_in_current_context=True if own_only else None)
    write_rows(out_path, operations.QUERY_FIELDS, rows, delimiter)
    typer.echo(f"{len(rows)} IZ in {len(operations.windows(d1, d2))} window(s) -> {out_path}")


@app.command("token-set")
def token_set(
    nip: Annotated[str, typer.Option(help="Company NIP the token belongs to")],
    env: EnvOpt = DEFAULT_ENV,
    profile: ProfileOpt = DEFAULT_PROFILE,
    stdin: Annotated[bool, typer.Option("--stdin", help="Read the token from stdin")] = False,
    test: Annotated[bool, typer.Option(help="Log in once to check the token")] = True,
) -> None:
    """Store a KSeF token in the OS credential store (input hidden) and select it for Excel."""
    nip = nip.strip().replace("-", "").replace(" ", "")
    err = userconfig.validate_nip(nip, env)
    if err:
        raise _fail(ValueError(err))
    if env == Env.PROD:
        typer.confirm("PROD: real data with legal effect. Store a PROD token?", abort=True)
    token = sys.stdin.readline().strip() if stdin else getpass.getpass("KSeF token (hidden): ")
    if not userconfig.looks_like_token(token.strip()):
        raise _fail(ValueError("that does not look like a KSeF token (20-200 characters, "
                               "no spaces)"))
    token = token.strip()
    if test:
        res = userconfig.test_connection(env, nip, token)
        if not res.ok:
            raise _fail(RuntimeError(res.message + " Token NOT stored."))
        typer.echo(res.message)
    ini = userconfig.save_setup(env, nip, token, profile)
    typer.echo(f"token stored in the OS credential store; workbook settings: {ini}")


@app.command("doctor")
def doctor() -> None:
    """Show what this computer is set up with (for support; prints no secrets)."""
    import keyring

    typer.echo(f"ksefiz {__version__} ({'packaged' if getattr(sys, 'frozen', False) else 'python'}"
               f", {sys.platform})")
    typer.echo(f"executable:        {sys.executable}")
    typer.echo(f"credential store:  {keyring.get_keyring().__class__.__module__}")
    ini = userconfig.read_excel_ini()
    typer.echo(f"workbook settings: {userconfig.excel_ini_path()} "
               f"{ini if ini else '(not set - run the setup assistant)'}")
    for env in Env:
        ctx = load_context(env)
        if ctx:
            has = "token stored" if get_token(env, ctx.nip) else "NO token"
            typer.echo(f"{env.value.upper():5} context:     NIP {ctx.nip}, {has}")
    wb = userconfig.workbook_path()
    typer.echo(f"workbook:          {wb} ({'present' if wb.exists() else 'missing'})")
    typer.echo(f"trusted location:  {userconfig.trusted_location_status()}")


@app.command("setup")
def setup_cmd() -> None:
    """Open the setup assistant (environment, NIP, token, connection test)."""
    from ksefiz import setup_gui

    setup_gui.main()


@app.command("excel-run", hidden=True)
def excel_run(
    job: Annotated[Path, typer.Option(file_okay=False, exists=True, help="Job folder")],
    op: Annotated[str, typer.Option(help="decode | generate | check | lookup | query")],
    env: EnvOpt = DEFAULT_ENV,
    profile: ProfileOpt = DEFAULT_PROFILE,
    date_from: Annotated[str, typer.Option("--from")] = "",
    date_to: Annotated[str, typer.Option("--to")] = "",
    own_only: bool = False,
    log_dir: Annotated[Path | None, typer.Option(file_okay=False)] = None,
    confirm_prod: Annotated[bool, typer.Option(help="Excel already asked the user")] = False,
) -> None:
    """Back end of the Excel buttons (file protocol in excel_job.py)."""
    if env == Env.PROD and not confirm_prod:
        raise _fail(RuntimeError("PROD from Excel needs --confirm-prod (asked by the macro)"))
    if op not in excel_job.OPS:
        raise _fail(ValueError(f"unknown --op {op!r}"))
    res = excel_job.run(job, op, env, profile, date_from=date_from, date_to=date_to,
                        own_only=own_only, log_dir=log_dir)
    typer.echo(f"{op}: {res.status} - {res.message}")
    if res.status == "error":
        raise typer.Exit(2)


if __name__ == "__main__":
    app()
