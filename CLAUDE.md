# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

An Excel workbook plus a Python helper CLI (`ksefiz`) for bulk KSeF 2.0 collective identifier (IZ) work: detect, validate, decode, generate, reverse lookup, period query. Full brief: `docs/BRIEF.md`. Decisions and open questions: `docs/decisions.md` (record new ones there as ADRs). Verified API facts: `docs/api-notes.md`.

**Status:** steps 1-6 are done. Offline M1/M2/M4 exist in Python and VBA (verified in Excel for Mac: self-test 42/42, `check_workbook.py` parity). All API modules work live on TEST, from the CLI and from Excel buttons (`modHelper.bas`/`modApi.bas` -> `ksefiz excel-run`; the Mac launcher is tested via osascript). Open questions 1-3 are answered in `docs/decisions.md`. The user tested the buttons. The repo is pushed to GitHub (`origin`, public; CI on push), so committed files must stay free of personal details: no local paths, user names, or the TEST context NIPs / KSeF numbers. Use MF documentation examples or `synthetic.fake_nip()`. Machine-specific workbooks come from `build_workbook.py --local` into git-ignored `output/`. Step 8 hand-over pack is in `docs/USER_GUIDE.md`, `docs/SECURITY_NOTE.md`, `docs/HANDOVER.md`: keep them in sync when behaviour changes. Remaining: step 7 (DEMO, needs credentials and explicit approval), a Windows Excel run, and the 500-invoice IZ (see SECURITY_NOTE section 9). Two TEST profiles exist: `default` = seller and `buyer` (`~/.ksefiz/test/context[.buyer].json`, seals next to them, tokens in keyring `ksefiz`/`test:<nip>`). Their TEST KSeF numbers are in `~/.ksefiz/test/invoices.csv`. Reuse these instead of re-running `testenv-setup --force`. Never call DEMO or PROD unless explicitly told to. When the spec and the brief disagree, `open-api.json` in https://github.com/CIRFMF/ksef-api wins.

## Commands

The venv lives outside the synced Documents folder (macOS re-flags `.venv` hidden, and Python 3.12+ then skips the editable-install `.pth`). `.venv` is a symlink to `~/.local/share/venvs/ksefiz`. Install into it with `VIRTUAL_ENV=~/.local/share/venvs/ksefiz uv pip install -e ".[dev]"`.

```bash
.venv/bin/pytest -q                              # all unit tests (-m e2e deselected by default)
.venv/bin/pytest tests/test_iz.py::test_crc8_spec_examples
.venv/bin/ruff check . && .venv/bin/mypy         # mypy is strict, over src/
.venv/bin/ksefiz gen-testdata | detect | validate | match   # see README for flags
.venv/bin/python scripts/build_workbook.py [--demo|--local] # template / generic demo (committed) / personal copy in output/
.venv/bin/python scripts/check_workbook.py <saved.xlsm>     # VBA vs Python parity on a saved workbook
.venv/bin/ksefiz auth-test --env test                       # live TEST login check
.venv/bin/pytest -m e2e                                     # live tests (TEST only)
```

## Architecture

- `src/ksefiz/iz.py` is the core. `DETECT_RE` is built programmatically from per-character classes, an in-segment line break `_BR`, and a between-part `_GAP`. Boundaries: no digit before the match, no hex character after it. `check()` normalises to the 35-char canonical form, then validates in the order format → NIP pattern → date → CRC; the first failure becomes the `reason`. Uppercasing and `\d` are ASCII-only on purpose, for VBA parity.
- **Every offline rule exists twice**: Python (`iz.py`, `match.py`, `io_csv.parse_amount`) and VBA (`excel/vba/modIZ.bas` uses a hand-written character scanner, not RegExp; `modMatch.bas` uses `Collection`, not Dictionary, so it also runs on Mac). Change both together. `tests/vectors.csv` and `tests/vectors_detect.csv` are the shared contract. pytest runs them, and `build_workbook.py` embeds them in the hidden `_Vectors` sheet for the `RunSelfTest` macro. Add a vector whenever a rule changes.
- Sheet column layouts are defined twice too: the field lists in Python (`synthetic.*_FIELDS`, `cli.DETECT_FIELDS`, `match.MATCH_FIELDS`) used by `build_workbook.py`, and the hard-coded column letters in the `.bas` headers. Keep them aligned.
- `synthetic.py` generates all test data from a seed, with expected detect and match results built in by construction. Generated NIPs deliberately fail the mod-11 checksum, so they can't be real. Keep every test datum synthetic. The only checksum-valid NIPs allowed are the two MF documentation examples.
- `.bas` files are CRLF and ASCII-only (`.gitattributes`), because the Windows VBE needs that on import. The `.xlsm` is never committed. It is built from the template plus the `.bas` files (manually on Mac, `scripts/build_xlsm.ps1` on Windows). Macros write results as values, never formulas.
- API layer: `client.py` (httpx, `X-Error-Format: problem-details`, `ApiError.codes`, 429/`Retry-After` retries, an `on_request` hook feeding `audit.AuditLog`) → `auth.py` (token route: challenge → MF `KsefTokenEncryption` key → RSA-OAEP-SHA256 `token|timestampMs` → poll → redeem/refresh/close; `Session.__repr__` hides tokens) and `xades.py` (TEST self-signed seal + XAdES-BES signer, bootstrap only, via `testenv.py`). `credentials.py` decides where secrets live: the keyring, plus the `KSEF_NIP`/`KSEF_TOKEN` env vars on TEST only. Mock HTTP with `respx` (`tests/test_auth_mock.py`).
- `operations.py` holds M3/M5/M6/M7. `Api` wraps login, `ensure_fresh`, the persisted `RateLimiter` (`ratelimit.py`, one shared `collectiveIdentifier` budget) and paging (`x-continuation-token`; `pageSize` max 500 for decode, 200 for lookup/query). Every function returns row dicts keyed by `*_FIELDS`, which also define the Excel columns. `prevalidate()` must reject anything the API would (71004/71005, etc.) before any call.
- Excel ↔ helper: `excel_job.py` defines the job-folder protocol (in.xlsx → progress.txt / cancel → out.xlsx + `_result` sheet, and `exit` written by the launcher). `excel/mac/ksefiz.applescript` (installed by `scripts/install_mac_helper.sh`) is the only way Mac VBA can start a process. Keep its guards (executable name, container path, quoting). Operations take an `operations.Monitor` for progress and cooperative cancel.
- Windows distribution: `packaging/windows/` (PyInstaller spec with two exes: `ksefiz.exe` + windowed `KSeF IZ Setup.exe`; Inno Setup `ksefiz.iss` with the per-user install, the HKCU Trusted Location for `%USERPROFILE%\KSeF IZ Tool\`, and backup-on-upgrade), CI `.github/workflows/windows-installer.yml` (build + silent-install smoke test). `userconfig.py` is the single source of truth for paths, `excel.ini` (read by `modSetup.AutoConfigure` on open) and the Trusted Location check. `setup_gui.py` is UI only. The release workbook `excel/dist/KSeF_IZ_Tool.xlsm` can only be built on Windows with Excel (`scripts/build_xlsm.ps1`). After any VBA change it must be rebuilt, and `scripts/verify_xlsm_vba.py` must pass (CI enforces it).
- `config.py` holds the environment URLs (default TEST) and the API limits (decode 10 per call, 500 per IZ by default, 132 IZ per invoice, 100-day query window). Rate limits for the collective endpoints are 20/s, 120/min and 240/h, all enforced together.
