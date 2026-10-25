# Decisions (ADR log)

## ADR-001 Python helper CLI behind Excel (2026-10-24)
KSeF token auth needs RSA-OAEP/SHA-256, which VBA cannot do (and Mac Excel not at all).
The `ksefiz` CLI handles auth, HTTP, rate limits and audit; Excel exchanges CSV files with it.
Offline modules (M1 detect, M2 validate, M4 match) exist in both Python and VBA.

## ADR-002 VBA `.xlsm` front end, built from a template (2026-10-24)
The macro-enabled workbook is the front end. Mac Excel cannot have VBA injected by script,
so `scripts/build_workbook.py` writes `excel/KSeF_IZ_Tool_template.xlsx` (sheets, headers,
self-test vectors), and the `.bas` modules in `excel/vba/` are imported once by hand.
`scripts/build_xlsm.ps1` automates that on Windows. The `.xlsm` itself is git-ignored because
working copies hold data. Depends on open question 5 (macro policy).

## ADR-003 Detection and normalisation rules (2026-10-24)
- Tolerated mangling: lowercase, missing hyphens, up to 3 of `- space tab NBSP` between ID parts,
  and line breaks (CR/LF plus indentation) anywhere inside or between parts.
- Boundaries: not preceded by a digit, not followed by a hex character. This allows glued
  prefixes like `REF:` or `FV`, but `…-E7A` is deliberately not detected (ambiguous).
- Candidates are reported even when the CRC fails, so cash application sees typos (`crc_ok=N`).
- Validation order: format -> NIP pattern -> date -> CRC; the first failure is the reason.
- Both IZ and KSeF numbers are normalised to the 35-char canonical form. A 36-char KSeF 1.0
  number is converted to 35 chars before its CRC is checked (assumption; KSeF 2.0 numbers are 35-char only).

## ADR-004 One hand-written scanner in VBA (2026-10-24)
The plan had `VBScript.RegExp` on Windows with a fallback scanner on Mac. Changed to a single
character scanner on both platforms, so there is one VBA code path to keep in step with the
Python regex. `Collection` is used instead of `Scripting.Dictionary` for the same reason.
Parity is enforced by `tests/vectors.csv` and `tests/vectors_detect.csv`, which pytest runs
and which the workbook's `RunSelfTest` macro runs from the hidden `_Vectors` sheet.

## ADR-005 Synthetic NIPs fail the mod-11 checksum (2026-10-24)
Generated NIPs match the KSeF NIP pattern but deliberately fail the official checksum, so
they can never be real companies. The only NIPs that pass are the two MF documentation examples.
Validation does **not** enforce the mod-11 checksum, because TEST context NIPs are random.

## ADR-006 M4 match semantics (2026-10-24)
One output row per decoded row. An invoice referenced several times (partial payments,
several IZ) is judged on the total paid across those rows. Statuses: `Matched`,
`Matched (no amount)` (details hidden), `Amount diff`, `Currency diff`, `Unmatched`,
`Duplicate in ERP`. Unpaid open items are not listed (the plan's "Unmatched (ERP)" status
was dropped, since in practice it would be most of the ledger).

## ADR-007 TEST bootstrap without user-supplied credentials (2026-10-24)
`ksefiz testenv-setup` creates its own TEST context. It draws a random NIP (checksum-invalid,
see ADR-005, and TEST accepts it), builds a self-signed company seal
(`organizationIdentifier=VATPL-<nip>`), logs in via XAdES (`verifyCertificateChain=false`),
then requests a KSeF token with **only** `CollectiveIdentifierManage` and waits for `Active`.
`/testdata/subject` is not needed, since a seal for a NIP gives owner rights on TEST.
The token goes to the OS keyring, the NIP and token reference to `~/.ksefiz/test/context.json`,
and the seal + key to `~/.ksefiz/test/seal.pem` (0600). Nothing lands in the repo.
The XAdES signer is hand-written (lxml + cryptography) after the MF C# `SignatureService`:
RSA-SHA256, exclusive C14N, enveloped document reference + SignedProperties reference.
It refuses to run outside TEST.

## ADR-008 Audit log format (2026-10-24)
`logs/audit.jsonl`, one JSON line per HTTP call (method, path, status, ms) plus one result
line per operation. No bodies, no token values. The Excel `Log` sheet gets its rows from the
helper once the buttons are wired up (step 6).

## ADR-009 TEST invoices for IZ tests (2026-10-24)
`scripts/send_test_invoices.py` (helper, not part of the tool) sends minimal synthetic FA(3)
invoices through an online session: AES-256-CBC/PKCS#7 with a fresh key per session, and the
key RSA-OAEP-encrypted to the MF `SymmetricKeyEncryption` certificate. It logs in with the
profile's self-signed seal (owner rights), because the `CollectiveIdentifierManage` token
cannot send invoices, by design. There are two TEST profiles: `default` = seller
(random test NIP), `buyer` (second random test NIP); the NIPs live only in `~/.ksefiz/test/`. Five seller->buyer invoices are for generate /
decode / lookup. One buyer->seller invoice is the second seller needed to test the 71004
mixed-sellers rejection. The resulting KSeF numbers are kept outside the repo in
`~/.ksefiz/test/invoices.csv` (TEST data is shared between integrators). The KSeF number
date segment is the KSeF acceptance date (Polish time), not `P_1`.

## ADR-010 API operations, paging and throttling (2026-10-24)
- `operations.Api` = token login, `ensure_fresh` before every call, the shared `collectiveIdentifier`
  rate-limit group (one budget for all four endpoints, effective values from `GET /rate-limits`),
  and `maxInvoices` from `GET /limits/context`. The session closes on exit.
- The limiter keeps sliding 1 s / 60 s / 3600 s windows, all enforced together. Timestamps are persisted per
  context in `~/.ksefiz/<env>/ratelimit.<nip>.json`, so consecutive CLI runs share the hourly budget.
  429 / `Retry-After` handling in `client.py` remains the safety net.
- Paging: `x-continuation-token` request header, `continuationToken` in the body. `pageSize` maxima
  differ: decode 500, lookup/query 200 (the API answers 21405 above them).
- Generate pre-validation mirrors every API rule we know: one seller (71004), no duplicates (71005),
  2..maxInvoices, CRC-valid KSeF numbers, seller_nip = KSeF number prefix, payment = amount + 3-letter
  currency, description <= 512. Groups default to the seller NIP, and `group_key` splits them further.
  A rejected group never reaches the API. The live e2e test confirms the API returns 71004/71005 for
  exactly these cases.
- Decode / lookup validate their input offline first: invalid IDs get a status row and no API call.
  Duplicate (normalised) IZ are decoded once.
- Period query splits any date range into <= 100-day UTC windows and dedupes across them.

## ADR-011 Excel buttons -> helper via job folders (2026-10-24)
- Each button run gets a job folder: Excel's sandbox container on Mac (`$HOME/ksefiz-jobs`), `%TEMP%\ksefiz-jobs` on Windows.
  Data goes both ways as `.xlsx` (Excel `SaveAs` / `Workbooks.Open`, openpyxl in Python), so there are no
  CSV encoding, delimiter or locale problems, and Polish text survives on both platforms.
- The helper runs in the background (`ksefiz excel-run`). It writes `progress.txt` (ASCII) after each unit
  of work, and checks for a `cancel` file between API calls (Esc in Excel). `out.xlsx` (rows + `_result`
  key/value sheet) is written even on error or cancellation. The launcher shell writes `exit` when
  the helper ends, so a crash is detected too, and the macro shows `log.txt`.
- Mac: VBA can only start processes through `AppleScriptTask`. `excel/mac/ksefiz.applescript` is the
  narrowest launcher that works. It only runs an executable named `ksefiz`, only for job folders inside
  Excel's container, and quotes every argument (tested against injection). VBA also restricts arguments
  to `[A-Za-z0-9-_.:/\ ]`.
- Windows: `WScript.Shell.Run` with `cmd /v:on` and every argument quoted.
- PROD from Excel: the macro asks every time and then passes `--confirm-prod`. `excel-run` refuses PROD without it.
- `2_Decode` layout changed: input IZ in column A, results in C:K (`3_Match` reads C:H). Input and output no
  longer overwrite each other.
- The job folder (with its input data) is deleted after each run, except when the helper stalls (kept for diagnosis).
- Audit from Excel runs goes to `~/.ksefiz/logs/audit.jsonl` by default (Config `log_folder` overrides).

## ADR-012 Batch test invoices and committed-file hygiene (2026-10-24)
- Large TEST invoice sets use a batch session (`send_test_invoices.py --batch --count N`): one
  encrypted ZIP part, not rate-limited, unlike online sending at 180/hour. The part upload goes to
  the pre-signed URL from KSeF and is refused for hosts outside `*.mf.gov.pl`.
- Workbooks come in three flavours: `template` and `demo` (synthetic, committed) and `local`
  (helper path + TEST invoices, `output/`, git-ignored). Workbook metadata is fixed
  (creator `ksefiz`, constant dates), so builds are reproducible and carry no user name.
- Committed files must not contain local paths, user names, TEST-context NIPs, TEST KSeF numbers
  or checksum-valid NIPs (only the MF documentation examples `1111111111` and `5265877635`).
  Synthetic ERP document numbers use an `FI` prefix, so they don't look like NIPs.

## ADR-013 Windows installer, setup assistant and self-configuring workbook (2026-10-25)
- **Distribution:** PyInstaller one-folder build with two executables sharing one runtime, `ksefiz.exe` (console, used by the
  buttons) and `KSeF IZ Setup.exe` (windowed Tkinter assistant). They're wrapped in a per-user Inno Setup installer with
  no admin rights. Silent install (`/VERYSILENT`) works for later IT deployment.
- **Locations** (`userconfig.py`):
  - helper: `%LOCALAPPDATA%\Programs\KSeF IZ Tool\`
  - workbook: `%USERPROFILE%\KSeF IZ Tool\`, outside Documents, because OneDrive Known Folder Move makes Excel open files
    via URL, where Trusted Locations don't apply
  - settings: `~/.ksefiz/excel.ini`
- **Trusted Location:** HKCU, that one folder, no subfolders, removed on uninstall. The user chose this over IT-managed
  macro trust. It needs security sign-off (SECURITY_NOTE item 11).
- **Workbook:** `ThisWorkbook.Workbook_Open` → `modSetup.AutoConfigure`. On Windows it copies `excel.ini` into Config and
  defaults `helper_path` to the installed exe. It adds the buttons if missing and runs the self-test quietly (README!A5),
  without leaving the workbook dirty. Mac (sandboxed) keeps manual Config.
- **Release workbook:** `excel/dist/KSeF_IZ_Tool.xlsm` is committed (exception in `.gitignore`), built by
  `build_xlsm.ps1` on Windows with Excel, which injects `excel/vba/ThisWorkbook.vba`. GitHub runners have no Excel.
  CI proves with oletools that the binary's VBA equals the text sources, so the opaque file can't hide code.
- **Token entry:** `ksefiz token-set` and the assistant store a token only after a successful test login.
- **Upgrades:** an existing workbook is renamed to a dated backup, never overwritten. The uninstaller never deletes workbooks.
- **TLS:** the OS trust store via `truststore`, so the tool works behind company TLS inspection.

## Open questions (from BRIEF section 10)
1. ~~IZ prefix NIP~~ **Answered 2026-10-24 on TEST:** an IZ generated by the *buyer* test context over the seller's invoices starts with the seller test NIP (`<seller-nip>-IZ...`), i.e. the **seller** NIP, as the spec says. The MF manual's "generating context" reading is wrong. Matching logic still does not depend on it.
2. ~~Effective max invoices per IZ~~ **Answered 2026-10-24:** `GET /limits/context` on the TEST context returns `collectiveIdentifier.maxInvoices = 500`. The tool reads this at run time rather than hard-coding it.
3. ~~Does decode return the payment amount when the buyer generated the IZ?~~ **Answered 2026-10-24 on TEST: yes.** The seller sees amount, currency and description of a buyer-generated IZ (`detailsHidden=false`), because the seller is a party on the invoice. Details are hidden only from contexts that are neither the creator nor a party. The seller's period query also lists buyer-generated IZ (`createdInCurrentContext=false`).
4. Token route vs certificate (XAdES) route for production: which does the security policy allow?
5. Is a macro-enabled workbook allowed in the target environment? (Also: is an AppleScript launcher acceptable on Mac clients, and is `WScript.Shell` allowed on Windows, i.e. not blocked by ASR rules?)
6. System of record for generated IZ if a vendor e-invoicing platform also generates them.
7. Behaviour when KSeF is unavailable: fail clearly and let users fall back to individual KSeF numbers.
