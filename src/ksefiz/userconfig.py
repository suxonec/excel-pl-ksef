"""Per-user setup shared by the setup assistant, `ksefiz token-set` / `doctor`, the Windows
installer (packaging/windows/ksefiz.iss) and the workbook (modSetup.AutoConfigure):

    %LOCALAPPDATA%\\Programs\\KSeF IZ Tool\\ksefiz.exe   helper (installer)
    %USERPROFILE%\\KSeF IZ Tool\\KSeF_IZ_Tool.xlsm        workbook, Excel Trusted Location
                                                       (deliberately not in OneDrive-synced
                                                       Documents: Excel ignores Trusted
                                                       Locations for files opened from OneDrive)
    ~/.ksefiz/excel.ini                                environment / NIP / profile the workbook
                                                       reads on open (ASCII, no secrets)
    OS keyring                                         the KSeF token (credentials.py)
"""

from __future__ import annotations

import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path

from ksefiz import auth
from ksefiz.client import ApiError, KsefClient
from ksefiz.config import Env
from ksefiz.credentials import (
    DEFAULT_PROFILE,
    Context,
    ksefiz_home,
    load_context,
    save_context,
    set_token,
)
from ksefiz.iz import NIP_RE, nip_checksum_ok

APP_NAME = "KSeF IZ Tool"
WORKBOOK_NAME = "KSeF_IZ_Tool.xlsm"
TRUSTED_LOCATION_KEY = "KSeFIZTool"
_OFFICE_KEY = r"Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations"
_POLICY_KEY = r"Software\Policies\Microsoft\Office\16.0\Excel\Security\Trusted Locations"
_INI_KEYS = ("environment", "context_nip", "profile")


def excel_ini_path() -> Path:
    return ksefiz_home() / "excel.ini"


def workbook_dir() -> Path:
    return Path.home() / APP_NAME


def workbook_path() -> Path:
    return workbook_dir() / WORKBOOK_NAME


def install_dir() -> Path:
    local = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    return local / "Programs" / APP_NAME


def write_excel_ini(env: Env, nip: str, profile: str = DEFAULT_PROFILE) -> Path:
    """key=value lines, ASCII only, so VBA can read them with Line Input."""
    if not (nip.isascii() and nip.isdigit()) or not re.fullmatch(r"[A-Za-z0-9_-]+", profile):
        raise ValueError("NIP must be digits and profile [A-Za-z0-9_-]")
    path = excel_ini_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        "; written by the KSeF IZ Tool setup assistant - read by the workbook on open",
        f"environment={env.value.upper()}",
        f"context_nip={nip}",
        f"profile={'' if profile == DEFAULT_PROFILE else profile}",
    ]
    path.write_text("\r\n".join(lines) + "\r\n", encoding="ascii")
    return path


def read_excel_ini() -> dict[str, str]:
    path = excel_ini_path()
    if not path.exists():
        return {}
    out: dict[str, str] = {}
    for line in path.read_text(encoding="ascii", errors="replace").splitlines():
        key, sep, value = line.partition("=")
        if sep and key.strip() in _INI_KEYS:
            out[key.strip()] = value.strip()
    return out


def validate_nip(nip: str, env: Env) -> str | None:
    """Error message, or None. DEMO/PROD need a real (checksum-valid) NIP."""
    nip = nip.strip().replace("-", "").replace(" ", "")
    if not (len(nip) == 10 and nip.isascii() and nip.isdigit()):
        return "A NIP has exactly 10 digits."
    if not NIP_RE.match(nip):
        return "This is not a valid NIP format."
    if env != Env.TEST and not nip_checksum_ok(nip):
        return "The NIP check digit is wrong - please check for a typo."
    return None


def looks_like_token(token: str) -> bool:
    return 20 <= len(token) <= 200 and token.isascii() and not any(c.isspace() for c in token)


# ---------------------------------------------------------------- Trusted Location


def trusted_location_status() -> str:
    """'registered', 'missing', 'blocked by policy' or 'n/a' (not Windows)."""
    if sys.platform != "win32":
        return "n/a"
    import winreg

    def _dword(root: int, key: str, name: str) -> int | None:
        try:
            with winreg.OpenKey(root, key) as k:
                return int(winreg.QueryValueEx(k, name)[0])
        except OSError:
            return None

    for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        if _dword(root, _POLICY_KEY, "AllowUserLocations") == 0:
            return "blocked by policy"
        if _dword(root, _POLICY_KEY, "AllLocationsDisabled") == 1:
            return "blocked by policy"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            _OFFICE_KEY + "\\" + TRUSTED_LOCATION_KEY) as k:
            path = str(winreg.QueryValueEx(k, "Path")[0])
    except OSError:
        return "missing"
    return "registered" if Path(path.rstrip("\\")) == workbook_dir() else "missing"


# ---------------------------------------------------------------- connection test + save


@dataclass
class TestResult:
    ok: bool
    message: str
    max_invoices: int | None = None


def explain(e: Exception) -> str:
    """Plain-language message for the assistant."""
    text = str(e)
    if isinstance(e, auth.AuthError):
        m = re.search(r"failed: (\d{3})", text)
        code = m.group(1) if m else ""
        if code == "415":
            return "KSeF accepted the token but it has no permission for this company (415)."
        if code == "450":
            return ("KSeF rejected the token (450): it is wrong, revoked, not yet active, or "
                    "belongs to another company/NIP.")
        return f"Login failed: {text}"
    if isinstance(e, ApiError):
        if e.status == 429:
            return "KSeF is limiting requests right now (429). Wait a minute and try again."
        if e.status in (400, 401, 403):
            return f"KSeF refused the request: {e.detail or e.title} (HTTP {e.status})."
        return f"KSeF error: {text}"
    if isinstance(e, OSError) or e.__class__.__module__.startswith("httpx"):
        return ("Cannot reach KSeF. Check the internet connection or company proxy "
                f"({e.__class__.__name__}).")
    return f"Unexpected error: {text}"


def test_connection(env: Env, nip: str, token: str) -> TestResult:
    try:
        with KsefClient(env, timeout=30) as client:
            session = auth.login_with_token(client, nip, token)
            try:
                lim = client.request("GET", "/limits/context", bearer=session.access_token)
                max_inv = int(lim["collectiveIdentifier"]["maxInvoices"])
            except (ApiError, KeyError, TypeError, ValueError):
                max_inv = None
            auth.close(client, session)
    except Exception as e:  # noqa: BLE001 - every failure becomes a message for the user
        return TestResult(False, explain(e))
    extra = f" Up to {max_inv} invoices per IZ." if max_inv else ""
    return TestResult(True, f"Connected to KSeF {env.value.upper()} as NIP {nip}.{extra}", max_inv)


def save_setup(env: Env, nip: str, token: str | None, profile: str = DEFAULT_PROFILE) -> Path:
    """Store the token (if given) in the OS keyring, the context, and excel.ini."""
    if token is not None:
        set_token(env, nip, token)
    existing = load_context(env, profile)
    ctx = existing if existing and existing.nip == nip and token is None else Context(nip)
    save_context(env, ctx, profile)
    return write_excel_ini(env, nip, profile)
