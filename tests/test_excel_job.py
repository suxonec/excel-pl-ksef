import random
from pathlib import Path
from typing import Any

import pytest

from ksefiz import excel_job
from ksefiz import operations as ops
from ksefiz.config import Env
from ksefiz.io_csv import read_xlsx_rows, write_xlsx
from ksefiz.synthetic import fake_nip, make_iz, make_ksef

rng = random.Random(11)
SELLER = fake_nip(rng)


def _result(job: Path) -> dict[str, str]:
    return {r["key"]: r.get("value", "") for r in read_xlsx_rows(job / "out.xlsx", "_result")}


def test_xlsx_round_trip_keeps_text_and_unicode(tmp_path: Path) -> None:
    rows = [{"a": "0001", "b": "Zapłata żółć", "c": ""}, {"a": "12", "b": "", "c": "x"}]
    write_xlsx(tmp_path / "t.xlsx", {"s": (["a", "b", "c"], rows)})
    assert read_xlsx_rows(tmp_path / "t.xlsx") == [{"a": "0001", "b": "Zapłata żółć", "c": ""},
                                                   {"a": "12", "b": "", "c": "x"}]


def test_check_job_is_offline(tmp_path: Path) -> None:
    k = [make_ksef(SELLER, "20261024", f"{i:012X}") for i in range(3)]
    fields = ops.GENERATE_IN
    rows = [dict.fromkeys(fields, "") | {"ksef_number": x, "group_key": g}
            for x, g in ((k[0], "A"), (k[1], "A"), (k[2], "B"))]
    write_xlsx(tmp_path / "in.xlsx", {"in": (fields, rows)})
    res = excel_job.run(tmp_path, "check", Env.TEST, "default", log_dir=tmp_path / "logs")
    assert res.status == "ok" and res.requests == 0 and res.input_hash
    out = read_xlsx_rows(tmp_path / "out.xlsx")
    assert [(r["group_key"], r["status"][:8]) for r in out] == [("A", "ready"), ("B", "rejected")]
    assert (tmp_path / "progress.txt").read_text() == "1 1 check: ok"
    assert (tmp_path / "logs" / "audit.jsonl").exists()


class FakeApi:
    """Stands in for operations.Api inside excel_job.run (no network)."""

    def __init__(self, *a: Any, **k: Any) -> None:
        from ksefiz.ratelimit import RateLimiter

        self.limiter = RateLimiter()
        self.max_invoices = 500

    def __enter__(self) -> "FakeApi":
        return self

    def __exit__(self, *exc: object) -> None:
        pass


@pytest.fixture
def fake_env(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    from ksefiz.credentials import Context

    monkeypatch.setattr(excel_job, "load_context", lambda env, profile: Context(SELLER))
    monkeypatch.setattr(excel_job, "get_token", lambda env, nip: "tok")
    monkeypatch.setattr(excel_job, "state_dir", lambda env: tmp_path / "state")
    monkeypatch.setattr(excel_job.operations, "Api", FakeApi)


def test_decode_job_writes_rows_progress_and_budget(
        tmp_path: Path, fake_env: None, monkeypatch: pytest.MonkeyPatch) -> None:
    izs = [make_iz(SELLER, "202610", f"{i:012X}") for i in range(25)]
    write_xlsx(tmp_path / "in.xlsx", {"in": (["iz"], [{"iz": i} for i in izs])})
    ticks: list[tuple[int, int]] = []

    def fake_decode(api: Any, values: list[str], monitor: ops.Monitor) -> list[dict[str, str]]:
        assert values == izs
        for n in range(1, 4):
            monitor.tick(n, 3)
            ticks.append((n, 3))
        return [ops._decode_row(i, "K", "1.00", "PLN", "", "false", ops.OK) for i in values]

    monkeypatch.setattr(excel_job.operations, "decode", fake_decode)
    res = excel_job.run(tmp_path, "decode", Env.TEST, "default", log_dir=tmp_path / "logs")
    assert res.status == "ok" and res.message == "25 rows" and res.budget_hour == "240"
    assert ticks == [(1, 3), (2, 3), (3, 3)]
    out = read_xlsx_rows(tmp_path / "out.xlsx")
    assert len(out) == 25 and list(out[0]) == ops.DECODE_FIELDS
    assert _result(tmp_path)["status"] == "ok"


def test_cancel_file_stops_and_reports_partial(
        tmp_path: Path, fake_env: None, monkeypatch: pytest.MonkeyPatch) -> None:
    write_xlsx(tmp_path / "in.xlsx", {"in": (["ksef_number"], [{"ksef_number": "x"}])})

    def fake_lookup(api: Any, values: list[str], monitor: ops.Monitor) -> list[dict[str, str]]:
        (tmp_path / "cancel").write_text("cancel")
        assert monitor.stopped()
        return [{"ksef_number": "x", "status": "OK"}]

    monkeypatch.setattr(excel_job.operations, "lookup", fake_lookup)
    res = excel_job.run(tmp_path, "lookup", Env.TEST, "default", log_dir=tmp_path / "logs")
    assert res.status == "cancelled" and "after 1 rows" in res.message
    assert len(read_xlsx_rows(tmp_path / "out.xlsx")) == 1


def test_errors_land_in_result_sheet(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(excel_job, "load_context", lambda env, profile: None)
    res = excel_job.run(tmp_path, "query", Env.TEST, "nobody", date_from="2026-10-01",
                        date_to="2026-10-31", log_dir=tmp_path / "logs")
    assert res.status == "error" and "profile 'nobody'" in res.message
    assert _result(tmp_path)["status"] == "error"


def test_operations_honour_monitor_stop() -> None:
    class StopAfterOne(ops.Monitor):
        def __init__(self) -> None:
            self.n = 0

        def tick(self, done: int, total: int) -> None:
            self.n = done

        def stopped(self) -> bool:
            return self.n >= 1

    class NoCallApi:
        def call(self, *a: Any, **k: Any) -> dict[str, str]:
            return {"collectiveIdentifierNumber": "IZ"}

    groups = ops.prevalidate([{"ksef_number": make_ksef(SELLER, "20261024", f"{i:012X}"),
                               "group_key": f"G{i // 2}"} for i in range(6)])
    rows = ops.generate(NoCallApi(), groups, StopAfterOne())  # type: ignore[arg-type]
    assert len(rows) == 1


def test_excel_run_refuses_prod_without_confirmation(tmp_path: Path) -> None:
    from typer.testing import CliRunner

    from ksefiz.cli import app

    r = CliRunner().invoke(app, ["excel-run", "--job", str(tmp_path), "--op", "check",
                                 "--env", "prod"])
    assert r.exit_code == 1 and not (tmp_path / "out.xlsx").exists()
