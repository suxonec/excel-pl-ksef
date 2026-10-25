from pathlib import Path

from typer.testing import CliRunner

from ksefiz import match as m
from ksefiz import synthetic
from ksefiz.cli import app
from ksefiz.io_csv import parse_amount, read_rows

K1 = "5265877635-20250826-0100001AF629-AF"
K2 = "7340009461-20261130-00FF00FF00FF-77"
IZ = "1111111111-IZ202607-65ED02180000-E7"


def _oi(k: str, amount: str, cur: str = "PLN", doc: str = "D1") -> dict[str, str]:
    return {"doc_no": doc, "customer": "C1", "ksef_number": k, "amount": amount, "currency": cur}


def _dec(k: str, amount: str, cur: str = "PLN", iz: str = IZ) -> dict[str, str]:
    return {"iz": iz, "ksef_number": k, "amount": amount, "currency": cur}


def _status(decoded: list[dict[str, str]], open_items: list[dict[str, str]]) -> list[str]:
    return [r.status for r in m.match(decoded, open_items)]


def test_statuses() -> None:
    oi = [_oi(K1, "100.00")]
    assert _status([_dec(K1, "100,00")], oi) == [m.STATUS_MATCHED]
    assert _status([_dec(K1, "90.00")], oi) == [m.STATUS_AMOUNT_DIFF]
    assert _status([_dec(K1, "")], oi) == [m.STATUS_NO_AMOUNT]
    assert _status([_dec(K1, "100.00", "EUR")], oi) == [m.STATUS_CURRENCY_DIFF]
    assert _status([_dec(K2, "100.00")], oi) == [m.STATUS_UNMATCHED]
    assert _status([_dec(K1, "100.00")], oi + [_oi(K1, "100.00", doc="D2")]) == [
        m.STATUS_DUPLICATE_ERP]


def test_ksef_number_keys_are_normalised() -> None:
    oi = [_oi(K1.replace("-", "").lower(), "100.00")]
    assert _status([_dec(K1.replace("-", " "), "100.00")], oi) == [m.STATUS_MATCHED]


def test_invoice_in_several_iz_judged_on_total() -> None:
    oi = [_oi(K1, "100.00")]
    split = [_dec(K1, "40.00", iz="A"), _dec(K1, "60.00", iz="B")]
    rows = m.match(split, oi)
    assert [r.status for r in rows] == [m.STATUS_MATCHED, m.STATUS_MATCHED]
    assert all(r.times_referenced == 2 and r.total_paid == 100.0 for r in rows)
    twice = [_dec(K1, "100.00", iz="A"), _dec(K1, "100.00", iz="B")]
    assert _status(twice, oi) == [m.STATUS_AMOUNT_DIFF] * 2


def test_synthetic_dataset_statuses() -> None:
    ds = synthetic.generate()
    assert _status(ds.decoded, ds.open_items) == [e["status"] for e in ds.expected_match]


def test_parse_amount() -> None:
    assert parse_amount("1 234,56") == 1234.56
    assert parse_amount("1.234,56") == 1234.56
    assert parse_amount("1,234.56") == 1234.56
    assert parse_amount("12") == 12.0
    assert parse_amount("") is None
    assert parse_amount("abc") is None


def test_cli_end_to_end(tmp_path: Path) -> None:
    runner = CliRunner()
    data = tmp_path / "data"
    assert runner.invoke(app, ["gen-testdata", "--out-dir", str(data)]).exit_code == 0

    found = tmp_path / "found.csv"
    r = runner.invoke(app, ["detect", "--in", str(data / "statements_sample.csv"),
                            "--out", str(found)])
    assert r.exit_code == 0, r.output
    expected = read_rows(data / "statements_expected.csv")
    with_valid = {e["line_id"] for e in expected if e["valid"] == "Y"}
    assert f"{1000 - len(with_valid)} lines without a valid ID" in r.output
    got = read_rows(found)
    assert [(g["line_id"], g["normalised"], g["valid"]) for g in got] == [
        (e["line_id"], e["normalised"], e["valid"]) for e in expected]

    out = tmp_path / "match.csv"
    r = runner.invoke(app, ["match", "--decoded", str(data / "decoded_sample.csv"),
                            "--open-items", str(data / "open_items_sample.csv"),
                            "--out", str(out), "--delimiter", ";"])
    assert r.exit_code == 0, r.output
    assert [x["status"] for x in read_rows(out)] == [
        e["status"] for e in read_rows(data / "match_expected.csv")]

    ids = tmp_path / "ids.csv"
    ids.write_text("id\n" + IZ + "\n" + IZ[:-2] + "00\n", encoding="utf-8")
    res = tmp_path / "res.csv"
    r = runner.invoke(app, ["validate", "--in", str(ids), "--out", str(res)])
    assert r.exit_code == 0, r.output
    assert [x["valid"] for x in read_rows(res)] == ["Y", "N"]
