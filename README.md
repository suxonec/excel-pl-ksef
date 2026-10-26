# KSeF collective ID (IZ) tool

From 2027-01-01, Polish B2B transfers must quote a KSeF number or a collective identifier (IZ). This tool works with those identifiers in bulk, as an Excel workbook backed by a Python helper CLI (`ksefiz`).
It is a bridge for cash application and AP teams, not the long-term ERP integration. The full brief is in [docs/BRIEF.md](docs/BRIEF.md).

**Hand-over pack:** [user guide](docs/USER_GUIDE.md), [security and compliance note](docs/SECURITY_NOTE.md), [hand-over summary and acceptance evidence](docs/HANDOVER.md).

**Status: steps 1-6 done.** The offline modules work without network access or credentials. All API modules work end to end on KSeF TEST, from the CLI and from Excel buttons:

| Module | Excel | CLI |
|---|---|---|
| M1 Detect IZ / KSeF numbers in statement text (lines without a valid ID turn red) | `1_Detect` + **Detect IDs** | `ksefiz detect` |
| M2 Validate format + CRC-8 | worksheet functions | `ksefiz validate` |
| M4 Match decoded rows to ERP open items | `3_Match` + **Match open items** | `ksefiz match` |
| M3 Decode IZ to invoices | `2_Decode` + button | `ksefiz decode` |
| M5 Generate IZ per seller / group | `4_Generate` + button | `ksefiz generate [--dry-run]` |
| M6 Reverse lookup | `5_Lookup` + button | `ksefiz lookup` |
| M7 Period query | `6_Query` + button | `ksefiz query --from --to` |
| M8 Audit | `Log` | `logs/audit.jsonl` |

```bash
.venv/bin/ksefiz generate --in invoices.csv --out iz.csv --dry-run   # pre-checks only, no API
.venv/bin/ksefiz generate --in invoices.csv --out iz.csv             # seller_nip, ksef_number, amount,
                                                                      # currency, description, group_key
.venv/bin/ksefiz decode   --in iz.csv --out decoded.csv              # column `iz`
.venv/bin/ksefiz lookup   --in ksef.csv --out lookup.csv             # column `ksef_number`
.venv/bin/ksefiz query    --from 2026-01-01 --to 2026-12-31 --out q.csv [--own-only]
```

All API commands take `--env` (default `test`) and `--profile`. They print the remaining `collectiveIdentifier` budget, which is shared across runs.

## Windows install (end users)

Download `KSeF-IZ-Tool-Setup-<version>.exe` (CI artifact of the `windows-installer` workflow) and run it. It installs per user (no admin rights), registers the workbook folder as an Excel Trusted Location, and starts the setup assistant (environment, NIP, KSeF token, connection test). Details are in [docs/USER_GUIDE.md](docs/USER_GUIDE.md#installing-on-windows).

How it's built:
- `packaging/windows/ksefiz.spec` (PyInstaller) builds `ksefiz.exe` plus the windowed `KSeF IZ Setup.exe`.
- `packaging/windows/ksefiz.iss` (Inno Setup) builds the installer.
- The bundled workbook is `excel/dist/KSeF_IZ_Tool.xlsm`, built once per VBA change on a Windows PC with Excel: `powershell -ExecutionPolicy Bypass -File scripts\build_xlsm.ps1`.
- Before committing it, run `scripts/sanitize_xlsm.py` on it. Excel stores the user name and the save folder in every saved file.
- CI checks with `scripts/verify_xlsm_vba.py` that the workbook contains exactly the code in `excel/vba/` and no user or machine details.

The workbook has a **Load sample data** button (README sheet). It copies synthetic examples from a hidden `_Sample` sheet into the input sheets, so Detect, Check (no API) and Match can be tried right away.

Support commands: `ksefiz doctor` (what this PC is set up with, no secrets), `ksefiz token-set --env <env> --nip <NIP>` (store a token, hidden input), `ksefiz setup` (the assistant).

## KSeF TEST environment

```bash
.venv/bin/ksefiz testenv-setup        # one-time: random test NIP, self-signed seal, XAdES login,
                                      # KSeF token with CollectiveIdentifierManage only -> OS keyring
.venv/bin/ksefiz testenv-setup --profile buyer   # second test company (buyer)
.venv/bin/ksefiz auth-test --env test # token login, refresh, context limits, logout
.venv/bin/python scripts/send_test_invoices.py   # 5 seller->buyer + 1 buyer->seller FA(3)
                                      # invoices -> ~/.ksefiz/test/invoices.csv
.venv/bin/pytest -m e2e               # live tests against TEST
KSEFIZ_E2E_HEAVY=1 .venv/bin/pytest -m e2e   # + generate/decode 50 IZ with paging
```

No credentials are needed from anyone, because TEST accepts self-signed certificates. The token is kept in the OS keyring (service `ksefiz`), never in files. The NIP and token reference go in `~/.ksefiz/test/context.json`. Every HTTP call is logged in `logs/audit.jsonl` without bodies or tokens. DEMO and PROD need real certificates, and `--env prod` asks for confirmation. TEST has daily maintenance from 16:00 to 18:00 Polish time.

## Python helper

```bash
uv venv -p 3.12 ~/.local/share/venvs/ksefiz && ln -s ~/.local/share/venvs/ksefiz .venv
VIRTUAL_ENV=~/.local/share/venvs/ksefiz uv pip install -e ".[dev]"

.venv/bin/ksefiz gen-testdata                      # synthetic data -> testdata/
.venv/bin/ksefiz detect   --in testdata/statements_sample.csv --out output/found.csv
.venv/bin/ksefiz validate --in ids.csv --out output/validated.csv [--col id]
.venv/bin/ksefiz match    --decoded testdata/decoded_sample.csv \
                          --open-items testdata/open_items_sample.csv --out output/match.csv
```

CSV files are read and written as UTF-8 with a BOM. On read, `;`, `,` or tab is detected automatically. Pass `--delimiter ";"` for Polish-locale Excel.

On macOS, keep the venv outside a synced `Documents` folder. macOS can mark the venv's files hidden, and Python 3.12+ then skips the `.pth` files that make the editable install importable.

## Excel workbook

1. Build a workbook with `.venv/bin/python scripts/build_workbook.py`, choosing one of three flavours:
   - no flag: `excel/KSeF_IZ_Tool_template.xlsx`, empty sheets (committed);
   - `--demo`: `excel/KSeF_IZ_Tool_demo.xlsx`, prefilled with synthetic statements, decoded rows and open items, with nothing machine-specific (committed);
   - `--local`: `output/KSeF_IZ_Tool_local.xlsx`, the demo plus this machine's `helper_path` and the KSeF TEST invoices from `~/.ksefiz/test/invoices.csv`. It also writes **`output/KSeF_IZ_Tool_local.xlsm`**: the same data in a copy of the release workbook, with all macros, ready to use without importing VBA. It's for your own testing and is **never committed** (`output/` is git-ignored).
2. Open it in Excel. Open the VBA editor (Alt+F11, or Tools > Macro > Visual Basic Editor on Mac). Use File > Import File for every `excel/vba/*.bas` (`modIZ`, `modDetect`, `modMatch`, `modSelfTest`, `modSetup`, `modHelper`, `modApi`).
3. Run the macro `SetupWorkbook`. It adds the buttons.
4. Save as **Excel Macro-Enabled Workbook** (`excel/KSeF_IZ_Tool.xlsm`). The `.xlsm` is git-ignored.
5. Click **Run self-test**. It replays the Python test vectors (`tests/vectors*.csv`) against the VBA code.

On Windows, `scripts/build_xlsm.ps1` does steps 2-4 automatically. It needs "Trust access to the VBA project object model".

### API buttons (decode, generate, lookup, query)

The buttons start `ksefiz` in the background. They show progress in the status bar, Esc cancels after the current request, and results are written back as values. Each run adds a row to the `Log` sheet (request count, input hash, error code, trace id), and `Config > budget_left` shows the remaining `collectiveIdentifier` budget.

- **Config:** set `helper_path` to the full path of the executable. That's `<repo>/.venv/bin/ksefiz` on Mac or `<repo>\.venv\Scripts\ksefiz.exe` on Windows (`--local` fills it in). `environment` defaults to TEST, and PROD asks for confirmation on every run. `profile` picks a stored context (blank = default). `log_folder` is an absolute folder, or blank for `~/.ksefiz/logs`.
- **Mac, once:** run `scripts/install_mac_helper.sh`. Excel's sandbox only lets VBA start programs through an AppleScript in `~/Library/Application Scripts/com.microsoft.Excel/`. The installed `ksefiz.scpt` will only start an executable named `ksefiz`, for a job folder inside Excel's container, with every argument quoted.
- **Sheets:**
  - `2_Decode`: IZ in A (**Take IZ from 1_Detect**, then **Decode**). Results go to C:K, which `3_Match` reads.
  - `4_Generate`: invoices in A:F. **Check (no API)** runs the pre-checks only, then **Generate IZ**.
  - `5_Lookup`: KSeF numbers in A.
  - `6_Query`: from / to in B1 / B2, `own_only` in B3.
- **How it works:** the job files exchanged between the macro and the helper (`in.xlsx`, `progress.txt`, `cancel`, `out.xlsx`) are described in `src/ksefiz/excel_job.py`.

Worksheet functions: `IZ_VALID`, `KSEF_VALID`, `IZ_NORMALISE`, `IZ_CRC8`, `ID_REASON` (`OK` / `format` / `nip` / `date` / `crc`) and `ID_KIND`.

To check the VBA against Python on any data: run Detect and Match in the workbook, save it, then run
`.venv/bin/python scripts/check_workbook.py <file.xlsm>`.

After editing VBA inside Excel, export the changed module (VBE > File > Export File) back to `excel/vba/`, so the repo stays the source of truth.

## Development

```bash
.venv/bin/pytest -q                      # unit tests (e2e against KSeF TEST deselected)
.venv/bin/pytest tests/test_iz.py -k crc # single test
.venv/bin/ruff check . && .venv/bin/mypy
```

All data in this repo is synthetic. Generated NIPs deliberately fail the mod-11 NIP checksum, so they can never match a real company. Design decisions and open questions are in [docs/decisions.md](docs/decisions.md). API facts are in [docs/api-notes.md](docs/api-notes.md).
