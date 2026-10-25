from pathlib import Path

import pytest
from typer.testing import CliRunner

from ksefiz import credentials, userconfig
from ksefiz.cli import app
from ksefiz.config import Env

MF_NIP = "5265877635"  # MF documentation example, checksum-valid
TOKEN = "20261025-EC-0000000000-0000000000-00|nip-5265877635|" + "a" * 64


@pytest.fixture(autouse=True)
def isolated(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> dict[str, str]:
    """Own KSEFIZ_HOME and an in-memory credential store for every test."""
    monkeypatch.setenv("KSEFIZ_HOME", str(tmp_path / "home"))
    monkeypatch.delenv("KSEF_NIP", raising=False)
    monkeypatch.delenv("KSEF_TOKEN", raising=False)
    store: dict[str, str] = {}
    monkeypatch.setattr(credentials.keyring, "set_password",
                        lambda svc, user, pw: store.__setitem__(f"{svc}/{user}", pw))
    monkeypatch.setattr(credentials.keyring, "get_password",
                        lambda svc, user: store.get(f"{svc}/{user}"))
    return store


def test_excel_ini_round_trip_is_ascii_crlf() -> None:
    path = userconfig.write_excel_ini(Env.PROD, MF_NIP)
    raw = path.read_bytes()
    assert raw.isascii() and b"\r\n" in raw
    assert userconfig.read_excel_ini() == {"environment": "PROD", "context_nip": MF_NIP,
                                           "profile": ""}
    userconfig.write_excel_ini(Env.TEST, "1111111111", "buyer")
    assert userconfig.read_excel_ini()["profile"] == "buyer"
    with pytest.raises(ValueError):
        userconfig.write_excel_ini(Env.TEST, "1111111111", "bad profile")


@pytest.mark.parametrize(("nip", "env", "ok"), [
    (MF_NIP, Env.PROD, True),
    ("526-587-76-35", Env.PROD, True),       # separators tolerated
    ("5265877636", Env.PROD, False),         # wrong check digit
    ("5265877636", Env.TEST, True),          # TEST NIPs are random
    ("0265877635", Env.TEST, False),         # pattern
    ("12345", Env.TEST, False),
])
def test_validate_nip(nip: str, env: Env, ok: bool) -> None:
    assert (userconfig.validate_nip(nip, env) is None) is ok


def test_looks_like_token() -> None:
    assert userconfig.looks_like_token(TOKEN)
    assert not userconfig.looks_like_token("short")
    assert not userconfig.looks_like_token(TOKEN[:30] + " " + TOKEN[30:])


def test_save_setup_keeps_context_when_token_kept(isolated: dict[str, str]) -> None:
    userconfig.save_setup(Env.DEMO, MF_NIP, TOKEN)
    assert isolated == {f"ksefiz/demo:{MF_NIP}": TOKEN}
    credentials.save_context(Env.DEMO, credentials.Context(MF_NIP, "REF-1", ("X",)))
    userconfig.save_setup(Env.DEMO, MF_NIP, None)  # "keep the stored token"
    ctx = credentials.load_context(Env.DEMO)
    assert ctx is not None and ctx.token_reference == "REF-1"
    assert userconfig.read_excel_ini()["environment"] == "DEMO"


def test_token_set_reads_stdin_and_never_prints_the_token(isolated: dict[str, str]) -> None:
    r = CliRunner().invoke(app, ["token-set", "--env", "demo", "--nip", MF_NIP, "--stdin",
                                 "--no-test"], input=TOKEN + "\n")
    assert r.exit_code == 0, r.output
    assert TOKEN not in r.output
    assert isolated[f"ksefiz/demo:{MF_NIP}"] == TOKEN
    assert userconfig.read_excel_ini()["context_nip"] == MF_NIP


def test_token_set_rejects_bad_input(isolated: dict[str, str]) -> None:
    runner = CliRunner()
    r = runner.invoke(app, ["token-set", "--env", "demo", "--nip", "5265877636", "--stdin",
                            "--no-test"], input=TOKEN + "\n")
    assert r.exit_code == 1 and "check digit" in r.output
    r = runner.invoke(app, ["token-set", "--env", "demo", "--nip", MF_NIP, "--stdin",
                            "--no-test"], input="oops\n")
    assert r.exit_code == 1 and not isolated


def test_token_set_does_not_store_a_token_that_fails_login(
        isolated: dict[str, str], monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(userconfig, "test_connection",
                        lambda env, nip, tok: userconfig.TestResult(False, "rejected (450)"))
    r = CliRunner().invoke(app, ["token-set", "--env", "test", "--nip", MF_NIP, "--stdin"],
                           input=TOKEN + "\n")
    assert r.exit_code == 1 and "NOT stored" in r.output and not isolated


def test_doctor_prints_no_secrets(isolated: dict[str, str]) -> None:
    userconfig.save_setup(Env.TEST, MF_NIP, TOKEN)
    r = CliRunner().invoke(app, ["doctor"])
    assert r.exit_code == 0 and "token stored" in r.output and TOKEN not in r.output


def test_explain_maps_common_failures() -> None:
    from ksefiz import auth
    from ksefiz.client import ApiError

    assert "450" in userconfig.explain(auth.AuthError("authentication failed: 450 Bad (x)"))
    assert "no permission" in userconfig.explain(auth.AuthError("authentication failed: 415 x"))
    assert "429" in userconfig.explain(ApiError(429, "Too Many", ""))
    assert "Cannot reach KSeF" in userconfig.explain(ConnectionError("dns"))
