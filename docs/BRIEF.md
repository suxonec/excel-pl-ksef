# KSeF Collective ID Excel Tool - build brief for Claude Code

*Version 0.1 - 2026-10-24. Source: official MF integrator guide and OpenAPI spec https://github.com/CIRFMF/ksef-api (API 2.9.0, PROD since 2026-10-15)*

## Summary

Build an Excel-based tool, with a small helper process behind it, that runs the four KSeF 2.0 collective identifier (IZ, "identyfikator zbiorczy") operations in bulk: **detect, validate, decode, generate, reverse lookup**. It is a bridge for finance operations teams (cash application and accounts payable) from 2027-01-01, when Polish B2B payments must quote a KSeF number or a collective ID (art. 108g VAT Act). It replaces slow manual entry in a vendor screen (about 5 minutes per IZ) and is not the long-term ERP-integrated solution.

First milestone: a working proof of concept against the **KSeF TEST environment** with self-signed certificates and random test NIPs - no real company data.

## 1. Business context

- From 2027-01-01, buyers paying a seller by bank transfer must quote either the invoice's KSeF number (35 characters) or a collective ID that groups 2+ invoices from the **same seller**
- As **seller (AR / cash application)**: customers will pay with IZ references. The IZ must be decoded into KSeF numbers, then matched to open items in the ERP (the ERP already stores the KSeF number on FI documents)
- As **buyer (AP)**: when paying several invoices of one supplier in one transfer, an IZ must be generated and placed in the payment reference
- As **direct debit issuer**: the invoice issuer carries the referencing duty
- Partial payments and credit notes: one invoice can be in several IZ, so a reverse lookup is needed

## 2. Scope

### In scope (MVP)

| # | Module | API needed | Priority |
|---|---|---|---|
| M1 | **IZ detector** - extract IZ patterns from pasted bank statement / remittance text | No | Must |
| M2 | **IZ validator** - format check and CRC-8 checksum, offline | No | Must |
| M3 | **Bulk decoder** - IZ list to KSeF numbers (+ amount, currency, description where visible) | Yes | Must |
| M4 | **Open item matcher** - join decoded KSeF numbers to a pasted ERP open items extract | No | Must |
| M5 | **Bulk generator** - group invoices per seller, pre-validate, generate IZ | Yes | Should |
| M6 | **Reverse lookup** - KSeF number to the IZ that contain it | Yes | Should |
| M7 | **Period query** - list IZ created in the context in a date range | Yes | Could |
| M8 | **Audit log** - every API call written to a log sheet and a local file | No | Must |

### Out of scope

- Sending invoices to KSeF, interactive or batch sessions, UPO, invoice download
- Any direct ERP (SAP) integration - input and output is copy-paste or CSV only
- Production use until security and tax sign-off (see section 9)

## 3. Recommended architecture

Pure VBA is **not** recommended: KSeF token authentication needs RSA-OAEP with SHA-256 encryption against the MF public key certificate, which VBA cannot do cleanly (and Mac Excel cannot at all).

```
Excel workbook (.xlsm or .xlsx + Python in Excel / xlwings)
   |  input sheets, buttons, result sheets
   v
Helper CLI (Python 3.11+, packaged as a single exe if needed)
   |  auth, crypto, HTTP, retries, rate limiting, logging
   v
KSeF API 2.0  (TEST -> DEMO -> PROD)
```

- Preferred: **Python helper CLI** (`httpx`, `cryptography`, `openpyxl`) called from Excel via xlwings or a VBA `Shell` call that exchanges CSV/JSON files
- Alternative: PowerShell 7 helper, or a .NET helper built on the official MF client `CIRFMF/ksef-client-csharp` (Java: `CIRFMF/ksef-client-java`)
- The helper must also run standalone from the command line (`ksefiz decode --in iz.csv --out result.csv`) so it can be tested without Excel
- Modules M1, M2 and M4 must work fully offline in Excel (formulas or VBA), so the tool is still useful when the API is unavailable or not yet approved

## 4. KSeF API facts the build depends on

### Environments

| Code | Base URL | Notes |
|---|---|---|
| TEST | `https://api-test.ksef.mf.gov.pl/v2` | Release candidates, self-signed certificates allowed, data shared between integrators - use random NIPs only. Maintenance daily 16:00-18:00 |
| DEMO | `https://api-demo.ksef.mf.gov.pl/v2` | Production-like configuration, final validation |
| PROD | `https://api.ksef.mf.gov.pl/v2` | Real data, legal effect |

Environment must be a config setting, default TEST. PROD must need an explicit flag and a confirmation prompt.

### Authentication (KSeF token route)

1. `POST /auth/challenge` - returns `challenge`, `timestamp`, `timestampMs`
2. `GET /security/public-key-certificates` - pick the certificate meant for KSeF token encryption that is valid now; keep its `publicKeyId`
3. Build the string `"{ksefToken}|{timestampMs}"`, UTF-8 encode, encrypt with **RSA-OAEP, SHA-256, MGF1-SHA-256** using the public key from the DER certificate, Base64 encode
4. `POST /auth/ksef-token` with `challenge`, `contextIdentifier` (`{"type": "Nip", "value": "<10-digit NIP>"}`), `encryptedToken`, optional `publicKeyId` - returns `referenceNumber` and a temporary `authenticationToken`
5. Poll `GET /auth/{referenceNumber}` with `Authorization: Bearer <authenticationToken>` until the status shows success (handle failure codes)
6. `POST /auth/token/redeem` (bearer = authenticationToken) - returns `accessToken` and `refreshToken`, each with `validUntil`
7. `POST /auth/token/refresh` with the refresh token before the access token expires
8. On exit, `DELETE /auth/sessions/current`

Alternative route (later): XAdES-signed `AuthTokenRequest` via `POST /auth/xades-signature` with a qualified or KSeF certificate. Read `uwierzytelnianie.md` and `certyfikaty-KSeF.md` in the repo for details and C#/Java samples.

**Permissions:** any of `InvoiceRead`, `InvoiceWrite`, `CollectiveIdentifierManage`. The tool must be designed for the narrow **`CollectiveIdentifierManage`** permission (least privilege, no invoice content access).

Send header `X-Error-Format: problem-details` on all calls and parse `application/problem+json` errors.

### Collective identifier endpoints

| Operation | Endpoint | Request | Response |
|---|---|---|---|
| Generate | `POST /collective-identifiers` | `{"invoices": [{"ksefNumber": "...", "payment": {"amount": 123.45, "currency": "PLN"}, "description": "max 512 chars"}]}` - min 2 items, `payment` and `description` optional | `201` `{"collectiveIdentifierNumber": "1111111111-IZ202607-65ED02180000-E7"}` |
| Decode | `POST /collective-identifiers/invoices` | `{"collectiveIdentifierNumbers": ["..."]}` - **max 10 per call** | `invoices[]` with `ksefNumber`, `collectiveIdentifierNumber`, `payment`, `description`, `detailsHidden`; `continuationToken` for paging |
| Reverse lookup | `GET /collective-identifiers/ksef/{ksefNumber}` | - | `collectiveIdentifiers[]` with `collectiveIdentifierNumber`, `createdInCurrentContext`, `dateCreated`; `continuationToken` |
| Period query | `POST /collective-identifiers/query` | `dateCreatedFrom`, `dateCreatedTo` (required, **max 100 days**), optional `collectiveIdentifierNumber`, `invoiceCountFrom/To`, `createdInCurrentContext` | `collectiveIdentifiers[]` with number, `dateCreated`, `invoiceCount`, `createdInCurrentContext`; `continuationToken` |

Rules to enforce before calling the API:

- All invoices in one IZ must have the **same seller NIP** (error `71004` otherwise)
- No duplicate KSeF numbers in one request (error `71005`)
- Default max **500 invoices per IZ** (context override up to 5,000 per the spec; the MF manual says 10,000 - treat 500 as the safe default, read the effective value from the limits endpoint)
- One invoice can be in at most **132 IZ per context** (spec)
- Payment amount, currency and description are only returned to the IZ creator or a party on the invoice; otherwise `detailsHidden = true`
- Continuation token errors: `21418` (invalid token)

Read `limity/limity.md`, `limity/limity-api.md` and `api-changelog.md` for current limits and error codes; `GET /limits/context` and `GET /rate-limits` return effective values at run time.

### Rate limits (collective endpoints)

**20 per second, 120 per minute, 240 per hour** each. `POST /auth/ksef-token`: 60 per second.

- Throttle client-side (token bucket for all three windows)
- Respect `429` and `Retry-After`; exponential backoff with jitter
- Batch decode in groups of 10 - about 2,400 IZ per hour
- Show the remaining hourly budget in the Excel status bar or a status cell

### IZ format and offline validation

Format: `NNNNNNNNNN-IZYYYYMM-HHHHHHHHHHHH-CC` (35 characters)

- `NNNNNNNNNN` - 10-digit NIP (the spec says seller NIP; the MF manual says the NIP of the generating context - log both readings as an open question, do not hard-code either into matching logic)
- `IZ` - fixed prefix
- `YYYYMM` - year and month created
- `HHHHHHHHHHHH` - 12 uppercase hex characters
- `CC` - **CRC-8** over the first 32 characters (hyphens included), polynomial `0x07`, initial value `0x00`, no reflection, no final XOR, 2 uppercase hex characters

Detection regex (case-insensitive search, then normalise to uppercase):

```
\b\d{10}-IZ\d{6}-[0-9A-F]{12}-[0-9A-F]{2}\b
```

Also detect common bank-statement mangling: missing hyphens, spaces instead of hyphens, line breaks inside the ID. Normalise, then validate the CRC.

Reference CRC-8 (verified against the spec example `1111111111-IZ202607-65ED02180000-E7`):

```python
def crc8(data: bytes) -> str:
    c = 0
    for x in data:
        c ^= x
        for _ in range(8):
            c = ((c << 1) ^ 0x07) & 0xFF if c & 0x80 else (c << 1) & 0xFF
    return f"{c:02X}"

assert crc8(b"1111111111-IZ202607-65ED02180000") == "E7"
```

The same must exist as a VBA function (`IZ_CRC8`) so validation works in Excel without the helper. KSeF invoice numbers (35 characters) also carry a CRC-8 - check `kody-qr.md` / the spec for their format and add a `KSEF_VALID` check if the same algorithm applies.

## 5. Excel workbook design

| Sheet | Purpose |
|---|---|
| `README` | How to use, version, environment banner (TEST / DEMO / PROD in colour) |
| `Config` | Environment, context NIP, path to token file, log folder, limits. **No secrets in the workbook** |
| `1_Detect` | Paste statement lines (columns: line ID, date, amount, currency, payer, free text). Output: found IZ, normalised IZ, CRC valid Y/N, also any plain KSeF numbers found |
| `2_Decode` | IZ list (from 1_Detect or pasted). Button "Decode". Output one row per invoice: IZ, KSeF number, amount, currency, description, detailsHidden, status, error code, timestamp |
| `3_Match` | Paste ERP open items extract (document no., customer, KSeF number, amount, currency, due date). Join on KSeF number. Output: matched / unmatched / amount difference |
| `4_Generate` | Invoice list (seller NIP, KSeF number, amount, currency, description, group key). Pre-checks: same seller, duplicates, count within limit, CRC of each KSeF number. Button "Generate". Output: group key, IZ, count, status |
| `5_Lookup` | KSeF numbers in, IZ list out |
| `6_Query` | Date range in (split automatically into 100-day windows), IZ list out |
| `Log` | Append-only: timestamp, user, environment, operation, request count, input hash, result, error code, correlation ID |

Buttons call the helper and refresh results; long runs show progress and can be cancelled. Results must be written as values, never as live formulas pointing at the API.

## 6. Security and compliance requirements

- KSeF token stored **outside** the workbook: Windows Credential Manager / macOS Keychain via `keyring`, or an environment variable for TEST. Never logged, never written to CSV
- Access token kept in memory only; session closed on exit
- Least-privilege token with `CollectiveIdentifierManage` only
- PROD requires explicit `--env prod` plus an interactive confirmation
- Audit log for every call (no token values, no full invoice data beyond KSeF number and IZ)
- No telemetry, no external calls other than the KSeF API hosts
- Repo must contain **no client names, real NIPs, real KSeF numbers or real statement data**. Use generated test data only
- `.gitignore`: config with secrets, logs, `*.xlsm` working copies with data, `output/`

## 7. Proposed repository structure

```
ksef-iz-tool/
  README.md
  LICENSE
  pyproject.toml
  .gitignore
  .env.example                 # KSEF_ENV, KSEF_NIP, KSEF_TOKEN (TEST only)
  docs/
    BRIEF.md                   # this file
    api-notes.md               # endpoints, limits, error codes used
    decisions.md               # ADR log
  src/ksefiz/
    __init__.py
    cli.py                     # typer CLI: auth-test, detect, validate, decode, generate, lookup, query
    config.py
    auth.py                    # challenge, RSA-OAEP, redeem, refresh, close
    client.py                  # httpx client, problem-details parsing, retries
    ratelimit.py               # 20/s, 120/min, 240/h buckets
    iz.py                      # regex, normalise, crc8, validate
    operations.py              # decode, generate, lookup, query (+ paging)
    io_csv.py                  # CSV/XLSX in and out
    audit.py
  excel/
    KSeF_IZ_Tool.xlsm          # template, no data
    vba/                       # exported .bas modules, IZ_CRC8, buttons
  tests/
    test_iz.py                 # crc8, regex, normalisation (incl. spec example)
    test_ratelimit.py
    test_operations_mock.py    # respx / recorded responses
    test_e2e_test_env.py       # marked, runs only with TEST credentials
  testdata/
    statements_sample.csv      # synthetic
    open_items_sample.csv      # synthetic
  .github/workflows/ci.yml     # lint (ruff), type check (mypy), unit tests
```

## 8. Delivery plan

| Step | Output | Done when |
|---|---|---|
| 1 | Repo scaffold, CI, `iz.py` with CRC-8, regex, normalisation + tests | Spec example validates; mangled samples normalise |
| 2 | Excel offline modules M1, M2, M4 with VBA `IZ_CRC8` | Works with no network |
| 3 | TEST environment setup: create test subject (`POST /testdata/subject`), self-signed certificate auth, generate a KSeF token with `CollectiveIdentifierManage` | `ksefiz auth-test --env test` succeeds |
| 4 | Send 3-5 test invoices (FA(3)) to TEST so real KSeF numbers exist (helper script only, not part of the tool) | KSeF numbers available for IZ tests |
| 5 | M5 generate, M3 decode, M6 lookup, M7 query with paging, rate limiting, audit | End-to-end on TEST: generate IZ, decode it, reverse look up each invoice |
| 6 | Excel buttons wired to helper, progress, error display | Demo-ready workbook |
| 7 | DEMO environment validation | Same tests pass on DEMO |
| 8 | Hand-over pack: user guide, security note for review | Ready for security and tax review |

## 9. Acceptance criteria (proof of concept)

- Detects and validates IZ in a 1,000-line synthetic statement in under 5 seconds, zero false positives on random 35-character strings
- Generates an IZ on TEST for 2 and for 500 invoices; rejects mixed sellers and duplicates **before** calling the API
- Decodes 50 IZ on TEST within rate limits, with paging handled and every row logged
- Reverse lookup returns all IZ that contain a test invoice used in several IZ
- No secret appears in the workbook, logs, CSV output or git history
- Runs on Windows with Excel 365; CLI also runs on macOS

## 10. Open questions (carry into docs/decisions.md)

1. IZ prefix NIP: seller NIP (spec) or generating context NIP (MF manual)? Test on TEST by generating as a buyer context
2. Effective max invoices per IZ for our context: 500 default vs 10,000 in the manual
3. Does decode return payment amount when the IZ was generated by the customer (buyer)? Expect `detailsHidden = true` - confirm on TEST
4. Token route vs certificate (XAdES) route for production - which does the organisation's security policy allow?
5. Is a macro-enabled workbook allowed in the target environment, or must the front end be `.xlsx` + Python in Excel / a small web UI?
6. System of record for generated IZ if a vendor e-invoicing platform also generates them - avoid two registers
7. Behaviour when KSeF is unavailable: tool must fail clearly and let users fall back to individual KSeF numbers

## 11. Reference files in the MF repo

- `open-api.json` - full spec (source of truth for request and response schemas)
- `uwierzytelnianie.md` - authentication flow with C# and Java samples
- `tokeny-ksef.md` - generating and managing KSeF tokens
- `uprawnienia.md` - permissions incl. `CollectiveIdentifierManage`
- `limity/limity.md`, `limity/limity-api.md` - limits per endpoint
- `srodowiska.md` - environments
- `dane-testowe-scenariusze.md` - test data setup on TEST
- `api-changelog.md` - recent changes (collective ID error codes 71001, 71004, 71005, 21418)
- Official clients: https://github.com/CIRFMF/ksef-client-csharp, https://github.com/CIRFMF/ksef-client-java

## Instructions for Claude Code

- Read this brief and the referenced MF repo files before writing code; treat `open-api.json` as the source of truth if anything here conflicts
- Start with steps 1-2 (offline, no credentials), then stop and ask for TEST credentials before step 3
- Never call DEMO or PROD unless explicitly told to
- Keep all test data synthetic
- Record design decisions in `docs/decisions.md` as you go
