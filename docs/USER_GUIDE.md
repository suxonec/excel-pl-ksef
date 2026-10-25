# KSeF IZ Tool: user guide

For cash application (AR) and accounts payable (AP) teams. Version 0.1 (proof of concept, KSeF **TEST** environment).

## What the tool is for

From 1 January 2027, a Polish B2B bank transfer must quote either the invoice's **KSeF number** or a **collective identifier (IZ)** that groups two or more invoices from the same seller.

| You are… | You need to… | Tool sheet |
|---|---|---|
| Cash application (we are the seller) | find IZ in bank statement lines, turn each IZ into its KSeF numbers, match them to open items | `1_Detect` → `2_Decode` → `3_Match` |
| Accounts payable (we are the buyer) | create one IZ for several invoices of one supplier, put it in the payment reference | `4_Generate` |
| Either (partial payments, credit notes) | find every IZ an invoice appears in | `5_Lookup` |
| Either (reconciliation) | list the IZ created in or involving our company over a period | `6_Query` |

This is a bridge until the ERP does this natively. It replaces manual lookups in a vendor screen, which take about 5 minutes per IZ.

## What an IZ looks like

```
1111111111-IZ202607-65ED02180000-E7      (example from the MF API documentation)
└seller NIP┘  └yyyymm┘ └technical part┘ └check digits
```

- 35 characters. The first 10 digits are always the **seller's** NIP, even when the buyer created the IZ (confirmed on KSeF TEST).
- The last 2 characters are a checksum, so the tool can tell a typing or bank mangling error from a real IZ **without** contacting KSeF.
- A KSeF invoice number has the same shape and checksum, with a date instead of `IZYYYYMM`: `5265877635-20250826-0100001AF629-AF` (MF documentation example).

## Try it with sample data

On the `README` sheet, click **Load sample data**. It fills the sheets with synthetic examples. No real companies are involved, and the NIPs fail the checksum on purpose:

| Sheet | Sample | Try |
|---|---|---|
| `1_Detect` | 1,000 bank statement lines with IZ and KSeF numbers in mangled forms | **Detect IDs** finds the IDs. Lines without a valid ID turn red. |
| `2_Decode` | decoded invoice rows, as KSeF would return them | used as input for `3_Match` |
| `3_Match` | ERP open items | **Match open items** shows every match status |
| `4_Generate` | 10 invoices in 5 groups | **Check (no API)** shows one group of each outcome: ready, mixed sellers, duplicate, single invoice |

It replaces the current contents of those four sheets, after asking. The **KSeF buttons** (Decode, Generate IZ, Look up) need invoices that really exist in KSeF; the sample numbers don't.

## Sheets at a glance

| Sheet | You paste / enter | You click | The tool writes |
|---|---|---|---|
| `README` | – | **Run self-test** (after updates) | – |
| `Config` | environment, helper path, profile (set once) | – | `budget_left` after each KSeF run |
| `1_Detect` | statement lines in A:F | **Detect IDs** | found IDs in H:N |
| `2_Decode` | IZ in column A | **Take IZ from 1_Detect**, **Decode (KSeF)** | invoices in C:K |
| `3_Match` | ERP open items in A:F | **Match open items** | match result in H:S |
| `4_Generate` | invoices to group in A:F | **Check (no API)**, **Generate IZ (KSeF)** | IZ per group in H:L |
| `5_Lookup` | KSeF numbers in column A | **Look up (KSeF)** | IZ per invoice in C:H |
| `6_Query` | dates in B1/B2, own_only in B3 | **Run query (KSeF)** | IZ list from row 6 |
| `Log` | – | – | one row per run (append-only) |

The banner at the top of `README` shows the environment: **green TEST**, **amber DEMO**, **red PROD**. Results are always written as plain values. Re-running a step overwrites that step's result columns only.

Buttons marked *(KSeF)* contact KSeF. The others work fully **offline**, including when KSeF is down.

## Cash application: statement → invoices → open items

1. **1_Detect.** Paste the statement lines into A2:F: line id, date, amount, currency, payer, free text (the payment title). Click **Detect IDs**.
   - The tool finds IZ and KSeF numbers even when the bank has mangled them: lowercase, missing hyphens, spaces instead of hyphens, line breaks in the middle, or text glued to the front (`REF:`, `FV`).
   - Each hit shows `kind` (IZ / KSEF), the cleaned-up `normalised` form, `crc_ok` and `valid`.
   - `valid = N` with reason `crc` means the ID is mistyped. Ask the payer, or look it up by amount. Don't send it to KSeF.
   - Statement lines with **no valid IZ or KSeF number** (nothing found, or only mistyped IDs) are highlighted **red** in A:F. These are the ones to clear manually, for example by amount, payer or invoice number in the text. The status bar and the `Log` sheet show how many there are. The highlight is reset on every run.
2. **2_Decode.** Click **Take IZ from 1_Detect**. It copies every *valid* IZ, once each, into column A. You can also paste IZ there yourself. Then click **Decode (KSeF)**.
   - You get one row per invoice: IZ, KSeF number, amount, currency, description and `details_hidden`.
   - Amount and description are visible when your company created the IZ or is the seller or buyer on the invoice. Otherwise `details_hidden = true` and the amount is blank.
3. **3_Match.** Paste the ERP open-items extract into A2:F: document no., customer, KSeF number, amount, currency, due date. Click **Match open items**.

| Match status | Meaning | Typical action |
|---|---|---|
| `Matched` | KSeF number found, amounts and currency agree | clear the item |
| `Matched (no amount)` | found, but no payment amount in the IZ (not given, or hidden) | clear against the bank amount |
| `Amount diff` | found, total paid ≠ open amount (see `diff`) | partial or over-payment: follow the usual process |
| `Currency diff` | found, currency differs | check with the payer |
| `Unmatched` | KSeF number not in the ERP extract | wrong extract, or not our invoice |
| `Duplicate in ERP` | the KSeF number appears on 2+ ERP lines | fix the ERP data first |

When one invoice is paid through several IZ (`times_referenced` > 1), the status compares the **sum** of all payments (`total_paid`) with the open amount.

## Accounts payable: create an IZ for a supplier payment

1. **4_Generate.** Paste the invoices into A2:F: `seller_nip`, `ksef_number`, `amount`, `currency`, `description`, `group_key`.
   - Rows with the same `group_key` become one IZ. Leave it blank to get one IZ per seller.
   - `amount` and `currency` are optional, but if you fill one you must fill both. `description` can be up to 512 characters.
2. Click **Check (no API)**. Each group shows `ready` or `rejected before API: …`. Typical reasons:

| Rejection | Meaning |
|---|---|
| `mixed sellers … (KSeF error 71004)` | an IZ may contain invoices of **one** seller only |
| `duplicate KSeF numbers … (71005)` | the same invoice appears twice in a group |
| `1 invoice(s); an IZ needs at least 2` | one invoice: just quote its KSeF number |
| `… the context limit is 500 per IZ` | split the group (the limit is read from KSeF) |
| `invalid KSeF number (crc / format …)` | mistyped number |
| `seller_nip … differs from KSeF number prefix` | wrong seller NIP on that row |

3. Fix the rejected groups, then click **Generate IZ (KSeF)** and confirm. Each ready group gets its IZ in column I. Copy it into the transfer's payment reference.

A generated IZ is permanent in KSeF. On PROD it has legal effect, so check before you generate. One invoice can be in at most **132 IZ** per company.

## Reverse lookup and period query

- **5_Lookup.** Paste KSeF numbers into column A and click **Look up (KSeF)**. You get every IZ that contains each invoice, newest first, including IZ created by the other party (`created_in_current_context = false`). `not in any IZ` is a normal result.
- **6_Query.** Enter dates in B1 (from) and B2 (to). Set B3 `own_only` to TRUE to see only IZ your company created. Click **Run query (KSeF)**. Any range works: the tool splits it into 100-day windows automatically. As the seller you also see IZ that buyers created over your invoices.

## While a KSeF job runs

- The status bar shows progress, for example `KSeF decode: 3/5 (Esc = cancel)`. Excel is busy until the job finishes.
- **Esc** (Mac: also **Cmd + .**) cancels after the current request. Everything done so far is written to the sheet, and the message says `cancelled`.
- At the end, a message box appears only if something went wrong. It shows the KSeF error code and a trace id. Quote both when you report a problem.
- `Config > budget_left` shows how many KSeF requests are left: KSeF allows 240 per hour for these functions, shared by all four buttons. One decode request handles 10 IZ, so about 2,400 IZ per hour. When the budget is used up, the tool waits by itself.

## When KSeF is unavailable

- Detect, the checksum validation and Match keep working offline.
- KSeF buttons fail with a clear message. Fall back to paying or clearing with individual KSeF numbers.
- KSeF TEST has daily maintenance from **16:00 to 18:00** (Polish time).

## Troubleshooting

| Message | Fix |
|---|---|
| `ksefiz was not found…` | Windows: run the installer. Mac / manual setup: set `Config > helper_path` to the `ksefiz` executable. |
| `Could not start ksefiz through AppleScript` (Mac) | IT runs `scripts/install_mac_helper.sh` once. Allow Excel when macOS asks. |
| `no … context/token for profile …` | the token is not set up on this computer: run **KSeF IZ Tool setup** (Windows) or `ksefiz token-set` |
| `ksefiz did not report progress for 900 s` | network or KSeF problem. The job folder is kept for IT. |
| `HTTP 429` / long waits | the request budget is used up. Wait, or split the work. |
| `450 … Token nieaktywny` / `415 …` | the KSeF token was revoked or lacks the permission (IT) |
| Self-test not `42/42 passed` | the VBA modules are out of date. Re-import them (IT). |

Every button run adds a line to the `Log` sheet: time, user, environment, operation, number of KSeF requests, input fingerprint, result, error code and trace id. Don't delete rows. It's the audit trail.

## Installing on Windows

1. Run `KSeF-IZ-Tool-Setup-<version>.exe`. No administrator rights are needed.
   - Until the installer is code-signed, Windows shows **"Windows protected your PC"**. Click **More info → Run anyway**.
   - The installer puts the helper in `%LOCALAPPDATA%\Programs\KSeF IZ Tool\` and the workbook in `%USERPROFILE%\KSeF IZ Tool\KSeF_IZ_Tool.xlsm`. It registers **only that folder** as an Excel *Trusted Location*, so the workbook's macros run without a warning.
   - An existing workbook is kept as `KSeF_IZ_Tool.backup-<date>.xlsm` before it's replaced.
2. The **setup assistant** opens at the end of the installation:
   1. Choose the environment. **TEST** is for training. On TEST you can tick *"create a test company for me"* if you have no token.
   2. Enter the company NIP.
   3. Paste the KSeF token. It needs only the `CollectiveIdentifierManage` permission and is stored in Windows Credential Manager.
   4. The assistant tests the connection and shows a green tick or a plain-language error.
3. Open **KSeF IZ Tool** from the Start menu. The workbook picks up the environment and NIP by itself, adds its buttons, and checks itself on open: `README` row 5 shows `Self-test 42/42 passed`. There's nothing to change in `Config`.

To replace an expired token or switch environments (for example to PROD after sign-off), run **KSeF IZ Tool setup** from the Start menu again.

**Uninstalling** (Settings → Apps → KSeF IZ Tool) removes the helper and the Trusted Location. Your workbooks in `%USERPROFILE%\KSeF IZ Tool\` and the stored token are kept. Remove the token in Credential Manager (entry `ksefiz`) if needed.

**If macros are still blocked:** company policy may forbid user Trusted Locations. The assistant warns about this on its last page. Ask IT to trust the folder `%USERPROFILE%\KSeF IZ Tool\`. Keep working copies inside that folder: copies saved to OneDrive, SharePoint or e-mail attachments are not trusted.

## Manual setup (Mac, or without the installer)

1. Install the helper: Python 3.11+, then `pip install` from the repository (see `README.md`). On Mac, also run `scripts/install_mac_helper.sh`.
2. Store the KSeF token: `ksefiz token-set --env <test|demo|prod> --nip <NIP>`. The input is hidden, the token is tested, and it is saved in the credential store. On TEST, `ksefiz testenv-setup` creates a test company instead. The token needs **only** the `CollectiveIdentifierManage` permission.
3. Build the workbook: `scripts/build_workbook.py` (empty) or `--demo` (synthetic example data), import `excel/vba/*.bas`, run `SetupWorkbook` and save as `.xlsm`. On Windows, `scripts/build_xlsm.ps1` produces the release workbook. Keep working copies that contain real data out of shared folders and out of git.
4. In `Config`, set `helper_path`, `environment` (TEST until sign-off), and `profile` if more than one company context is used. On Windows these are filled in automatically from the setup assistant's settings.
5. Click **Run self-test**. It must show `42/42 passed`.
