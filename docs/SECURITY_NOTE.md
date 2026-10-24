# KSeF IZ Tool: security and compliance note

For the security review and the tax/compliance review before any use beyond KSeF TEST. Version 0.1, commit history on `main`. Status: proof of concept, verified on KSeF **TEST** only.

## 1. Summary for reviewers

- The tool works with KSeF collective identifiers (IZ) only. It **cannot** send, read or download invoices. It is designed for, and tested with, a KSeF token that carries **only** `CollectiveIdentifierManage`.
- The only network destinations are the three MF KSeF API hosts, over HTTPS with certificate verification. There is no telemetry, no other service, and no update check.
- Secrets: the KSeF token is held in the OS credential store (Windows Credential Manager / macOS Keychain). Short-lived access tokens exist in process memory only. Neither the token nor an access token is ever written to the workbook, logs, CSV or job files. This was verified by scanning all outputs, logs, state files and the git history.
- The default environment is TEST. PROD needs an explicit choice **and** a confirmation on every run.
- Two local execution paths need a policy decision: **VBA macros** (an `.xlsm` workbook) and, on Mac, an **AppleScript launcher** (on Windows, `WScript.Shell`). See section 5.
- Items that must be closed before PROD are listed in section 9.

## 2. Architecture and data flow

```
 Excel workbook (.xlsm, VBA)                      ksefiz helper (Python 3.11+)              KSeF API 2.0
 ─────────────────────────                        ───────────────────────────               ────────────
 paste data ─► sheet ─► job folder/in.xlsx ──────► excel-run: validate offline ──HTTPS──►  api-test / api-demo /
 values  ◄──── sheet ◄─ job folder/out.xlsx ◄────── results, progress, audit  ◄──────────  api.ksef.mf.gov.pl
                         (deleted after run)        token from OS keyring
```

| Component | Runs where | Network |
|---|---|---|
| Offline modules (detect, checksum validation, match) | VBA in the workbook, and Python | none |
| `ksefiz` CLI / `excel-run` | user's machine, user's privileges | KSeF hosts only |
| Mac launcher `ksefiz.scpt` | macOS, started by Excel via `AppleScriptTask` | none itself |

Hosts (hard-coded in `src/ksefiz/config.py`):
- `https://api-test.ksef.mf.gov.pl/v2`
- `https://api-demo.ksef.mf.gov.pl/v2`
- `https://api.ksef.mf.gov.pl/v2`

The HTTP client (`httpx`) verifies TLS certificates and does **not** follow redirects.

## 3. Data processed and where it is stored

| Data | Source | Sent to KSeF | Stored locally |
|---|---|---|---|
| IZ numbers, KSeF invoice numbers | statements, ERP extract, KSeF | yes (they are the subject of every call) | workbook; audit log (lookup paths contain the KSeF number) |
| Payment amount, currency, description | AP input / KSeF decode | only on generate (optional fields) | workbook only |
| Bank statement free text, payer, ERP customer, document no. | user paste | **no**, processed offline only | workbook only |
| Context NIP | config | yes (login context) | `~/.ksefiz/<env>/context*.json`; audit result lines |
| OS user name | OS | no | audit log, `Log` sheet |
| KSeF token | KSeF (generated once) | encrypted with the MF public key at each login | OS credential store only |
| Access / refresh tokens | KSeF | as Bearer header | **process memory only** |

- **Job folders** (`in.xlsx`, `out.xlsx`, progress, log) are created per button run in Excel's sandbox container (Mac) or `%TEMP%` (Windows), and are **deleted after the run**. The only exception: a job that stops reporting progress for 15 minutes is kept for diagnosis.
- **Audit log** `~/.ksefiz/logs/audit.jsonl` (or the CLI working folder's `logs/`) gets one JSON line per HTTP call (method, path, status, duration) and one result line per run (result, request count, input hash, error code, trace id, NIP). It contains **no request or response bodies** and **no token values**.
- **Personal data (GDPR):** for sole traders (JDG), a NIP and an invoice can relate to a natural person. The workbook, the audit log and the OS user name in the log are therefore personal data. Retention is **not** automated (see section 9).

## 4. Authentication and secrets

**Normal route: KSeF token** (`src/ksefiz/auth.py`)
1. `POST /auth/challenge`.
2. Fetch the MF public-key certificate whose usage is `KsefTokenEncryption` and that is valid now.
3. Encrypt `token|timestampMs` with RSA-OAEP (SHA-256 / MGF1-SHA-256).
4. `POST /auth/ksef-token`, then poll `GET /auth/{ref}` until success, then redeem.
5. The access token lasts about 15 minutes and is refreshed automatically before each call when it is close to expiry. The refresh token lasts up to 7 days per MF.
6. `DELETE /auth/sessions/current` when the run ends.

**Least privilege.** `CollectiveIdentifierManage` allows generating, decoding, reverse lookup and query of IZ only. Invoice sending, invoice download and permission management are **not** possible with it. That was verified on TEST: the same context needed a separate owner login to send test invoices.

**Token storage.**
- The token is held in `keyring`, service `ksefiz`, account `<env>:<nip>`. Anyone who can run code as the same OS user can read it. That is the normal trust boundary of a per-user credential store.
- `KSEF_TOKEN` / `KSEF_NIP` environment variables are honoured **on TEST only**.
- `Session.__repr__` hides tokens, so they cannot leak through tracebacks.

**TEST bootstrap** (`src/ksefiz/testenv.py`, `xades.py`). A self-signed company seal plus a XAdES signature is used to create TEST contexts and tokens. The code refuses to run against DEMO/PROD, which don't accept self-signed certificates anyway. The seal's private key is stored as `~/.ksefiz/test/seal*.pem` with mode `0600`.

**Environment guard.**
- The default is TEST.
- The CLI asks for interactive confirmation for `--env prod`.
- From Excel, the macro asks on **every** PROD run and only then passes `--confirm-prod`. `excel-run` refuses PROD without that flag.
- The `README` sheet banner shows the environment in colour.

## 5. Local code execution surface

**VBA macros** (`excel/vba/*.bas`, plain text in git, reviewable):
- Read and write cells in the workbook.
- Create a temporary workbook for the job input.
- Start the helper (next two bullets).
- Read the job result.

They make **no network calls** and use no `Declare` or Win32 API calls. The macros are **not code-signed** yet.

**Mac launcher** (`excel/mac/ksefiz.applescript`, installed per user in `~/Library/Application Scripts/com.microsoft.Excel/`). This is the only mechanism Excel for Mac offers to start a process. It is restricted to:
- an executable whose path ends in `/ksefiz` and is executable,
- a job folder inside Excel's own container,
- every argument passed through `quoted form of`.

Tests: starting `/bin/sh`, a job folder outside the container, and an argument containing `; touch /tmp/pwned` were all refused or inert. Residual risk: any macro-enabled workbook can call this launcher. The worst it can do is start `ksefiz` with chosen arguments, which on PROD still needs the explicit `--confirm-prod`.

**Windows** uses `WScript.Shell.Run` with `cmd /v:on`. VBA allows only `[A-Za-z0-9-_.:/\ ]` in arguments and quotes each one. Some endpoint-protection setups (Defender ASR rule "Block Office applications from creating child processes") block this. That needs an IT decision or an exclusion for the helper.

## 6. Integrity and correctness controls

- **Offline validation before any call:**
  - Checksum (CRC-8) and format checks on every IZ and KSeF number. Invalid IDs never reach KSeF.
  - Generate pre-checks that mirror KSeF's rules: one seller (71004), no duplicates (71005), 2…`maxInvoices` per IZ (read from `GET /limits/context`), payment amount + currency together, description ≤ 512.
  - A live test confirmed that KSeF returns exactly 71004/71005 for the cases the pre-checks block.
- **Rate limits.** A client-side throttle for the shared `collectiveIdentifier` budget (20/s, 120/min, 240/h, effective values read from `GET /rate-limits`). It is persisted across runs and honours `Retry-After` on 429. Repeated limit breaches would otherwise extend KSeF's lockout.
- **Parity of the two implementations.** Offline logic exists in Python and VBA. 42 shared test vectors run in pytest and in the workbook (`Run self-test`). `scripts/check_workbook.py` compares VBA and Python output row by row (verified on Excel for Mac: 1,021 detect rows and 160 match statuses identical).
- **Testing.**
  - CI: ruff, strict mypy and 91 unit tests on Windows, macOS and Linux with Python 3.11/3.12.
  - Live end-to-end tests on KSeF TEST.
  - All test data is synthetic. Generated NIPs deliberately fail the NIP checksum, so they can't be real companies.

## 7. Supply chain

Runtime dependencies (`pyproject.toml`): `httpx`, `cryptography`, `lxml`, `openpyxl`, `keyring` and `typer`, all widely used and maintained.

Current state: **minimum versions only, no lockfile, no hash pinning.** The tool is installed from source into a virtual environment. The brief's single-exe packaging is not done yet. See section 9.

## 8. Tax / compliance points to confirm

- **IZ semantics, observed on TEST:**
  - The IZ prefix is the **seller's** NIP even when the buyer generates the IZ, as the API spec says. The MF manual's "generating context" reading did not hold.
  - Both seller and buyer see payment amount and description (`detailsHidden=false`). Only third parties see them hidden.
- **Effect of generating:** a generated IZ is a permanent KSeF record (legal effect on PROD). The tool does not delete or modify IZ, because KSeF offers no such operation.
- **Payment data in the IZ is optional.** The tool sends amount/currency/description only when the user fills them in. Confirm whether the company wants to disclose payment amounts in IZ by default.
- **Partial payments and credit notes.** One invoice can be in several IZ (max 132 per context). The matcher judges the **total** paid across them. Confirm this reflects the intended clearing rule.
- **Art. 108g referencing duty:** the tool helps produce the reference. The legal assessment of which reference to use (KSeF number vs IZ, direct debit cases) stays with tax.
- **System of record:** if a vendor e-invoicing platform also generates IZ, decide which register is authoritative (open question 6).

## 9. Open items before PROD

| # | Item | Owner | Status |
|---|---|---|---|
| 1 | Validation on **DEMO** with a real certificate / token (delivery step 7) | IT + tax | not started; needs credentials |
| 2 | **Windows Excel 365** run of the buttons (VBA Windows branch and `build_xlsm.ps1` are untested; Python is CI-tested on Windows) | IT | open |
| 3 | Generate an IZ with **500 invoices** on TEST (acceptance criterion; needs 500 test invoices) | dev | open |
| 4 | Policy: macro-enabled workbook allowed? AppleScript launcher (Mac) / `WScript.Shell` + ASR rules (Windows)? Code-sign the VBA project? | security | open (Q5) |
| 5 | Production auth route: KSeF token vs XAdES certificate (Q4); who generates and rotates tokens; revocation procedure (`DELETE /tokens/{ref}`; no CLI command yet) | security | open |
| 6 | Dependency pinning (lockfile with hashes) and packaging (signed single executable) | dev + security | open |
| 7 | Retention and access rules for the audit log, the workbook copies and the `Log` sheet (GDPR for sole-trader data) | compliance | open |
| 8 | Behaviour and procedure when KSeF is unavailable (fallback to individual KSeF numbers) | finance ops | documented in the user guide; confirm (Q7) |
| 9 | System of record for generated IZ (Q6) | finance + tax | open |
| 10 | Publicly visible repository: content is synthetic / TEST only. Confirm that public visibility is intended. | owner | open |

References:
- `docs/decisions.md` (ADR-001 … ADR-011 and the open questions)
- `docs/api-notes.md` (verified API behaviour)
- `docs/USER_GUIDE.md`
- `docs/HANDOVER.md` (acceptance evidence)
