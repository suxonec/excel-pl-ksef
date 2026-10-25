# KSeF IZ Tool: hand-over summary

Proof-of-concept status as of 2026-10-24. Everything below was verified on the KSeF **TEST** environment with synthetic data.

| Document | For |
|---|---|
| [USER_GUIDE.md](USER_GUIDE.md) | cash application and AP users, key users |
| [SECURITY_NOTE.md](SECURITY_NOTE.md) | security review, tax/compliance review |
| [decisions.md](decisions.md) | design decisions (ADR-001 … ADR-011), open questions |
| [api-notes.md](api-notes.md) | KSeF API behaviour, verified live |
| [BRIEF.md](BRIEF.md) | original build brief |

## Delivery plan status (brief section 8)

| Step | Output | Status |
|---|---|---|
| 1 | Repo, CI, IZ/KSeF parsing and CRC-8 | done |
| 2 | Offline Excel modules M1, M2, M4 with VBA | done; Excel for Mac verified (self-test 42/42, row-level parity with Python) |
| 3 | TEST setup, token with `CollectiveIdentifierManage` | done (`ksefiz testenv-setup`, `auth-test`) |
| 4 | TEST invoices (FA(3)) so real KSeF numbers exist | done (helper script: 6 invoices via online session, 500 via one batch session; 2 test companies) |
| 5 | M5 generate, M3 decode, M6 lookup, M7 query, paging, rate limiting, audit | done, end to end on TEST |
| 6 | Excel buttons wired to the helper, progress, cancel, errors | done; tested on Excel for Mac, and offline buttons on Windows. Plus a Windows installer, setup assistant and sample data (ADR-013, ADR-014). |
| 7 | DEMO validation | **not started**: needs a real certificate/token and approval |
| 8 | Hand-over pack | this document set |

## Acceptance criteria (brief section 9)

| Criterion | Result | Evidence |
|---|---|---|
| Detect + validate IZ in a 1,000-line statement in < 5 s, zero false positives on random 35-char strings | **met**: 0.08 s; 0 valid hits on 10,000 random strings and 5,000 wrong-checksum IDs | `tests/test_iz.py` |
| Generate an IZ on TEST for 2 and for 500 invoices | **met**: 2, 5, 50 × 2 and **one IZ over 500 invoices** (3.5 s, one request). A 501-invoice group was rejected before the API. The 500-row decode returned exactly the sent KSeF numbers and amounts in 2 s. | `tests/test_e2e_test_env.py` (`KSEFIZ_E2E_500=1`), run logs |
| Reject mixed sellers and duplicates **before** calling the API | **met**, and the API itself was confirmed to return 71004 / 71005 for these cases | `tests/test_operations_mock.py`, e2e |
| Decode 50 IZ on TEST within rate limits, paging handled, every row logged | **met**: 50 IZ / 100 rows with forced paging in ~10 s. Every API call is in the audit log; the result rows go to the output sheet/CSV. | `KSEFIZ_E2E_HEAVY=1 pytest -m e2e` |
| Reverse lookup returns all IZ containing an invoice used in several IZ | **met**: 3 IZ for one invoice, including a buyer-created IZ | e2e, run logs |
| No secret in workbook, logs, CSV output or git history | **met**: scanned all repo files, logs, state files and the pushed commit for both live TEST tokens | scan before commit `06a68c8` |
| Runs on Windows with Excel 365; CLI also on macOS | **mostly**: CLI on macOS ✔; Python on Windows ✔ (CI); Excel for Mac ✔; **Windows 11 + Excel 365** ✔ for installer, setup assistant, Trusted Location, self-test and offline buttons. The KSeF buttons are not yet run on Windows (they need TEST invoices for the new test company). | owner test 2026-10-25, CI |

## Answers found during the build

1. **Whose NIP is in the IZ prefix?** The seller's, even when the buyer generates the IZ.
2. **Max invoices per IZ for our context:** 500 (`GET /limits/context`). The tool reads it at run time.
3. **Does decode show the amount when the buyer created the IZ?** Yes, to both parties on the invoice. It is hidden only from third parties.

Also learned:
- All four collective endpoints share **one** 240/hour budget.
- The lookup and query page size is capped at 200, unlike decode's 500.
- The date in a KSeF number is the KSeF acceptance date, not the invoice date.

## Still open

- **Code signing (release blocker):** Smart App Control blocks unsigned builds on Windows 11.
- **KSeF buttons on Windows:** Part B (the assistant creates TEST invoices for the new test company).

See SECURITY_NOTE.md section 9 for owners. In short:
- DEMO validation
- an Excel 365 run on Windows
- macro and launcher policy, and code signing
- production auth route and token rotation
- dependency pinning and packaging
- log retention
- the IZ system of record
- confirming the public repository visibility

## How to run it (developer)

```bash
.venv/bin/pytest -q                       # 91 unit tests
.venv/bin/pytest -m e2e                   # live TEST checks (needs ksefiz testenv-setup)
.venv/bin/ksefiz --help                   # detect, validate, match, decode, generate, lookup, query
.venv/bin/python scripts/build_workbook.py [--demo | --local]
```

Full setup is in `README.md`. TEST state lives outside the repo:
- `~/.ksefiz/test/`: contexts, seals, invoice list, rate-limit state
- `~/.ksefiz/logs/`: the audit log from Excel runs
- OS keyring service `ksefiz`: the tokens
