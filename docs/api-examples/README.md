# KSeF collective identifier (IZ): API message examples

Request and response examples for every call `ksefiz` makes to KSeF for collective identifiers, including the session calls around them. Each scenario is one JSON file. [ksef-iz.http](ksef-iz.http) has the same calls in a runnable form (TEST only).

Shapes come from `open-api.json` in [CIRFMF/ksef-api](https://github.com/CIRFMF/ksef-api) (API 2.9.0, checked 2026-10-27). Each file's `source` says whether the behaviour was also verified live on TEST ([../api-notes.md](../api-notes.md)). **If this folder and `open-api.json` disagree, `open-api.json` wins.**

## File format

```json
{
  "scenario": "what happens",
  "source": "where the shape comes from",
  "notes": ["optional remarks"],
  "request":  {"method": "POST", "url": "...", "query": {}, "headers": {}, "body": {}},
  "response": {"status": 201, "headers": {}, "body": {}},
  "excel_rows": [{"...": "the rows the tool writes to the result sheet"}]
}
```

- `request` and `response` are `null` when the tool rejects the input offline and sends nothing.
- `context_nip` (optional, in `request`) is the NIP the session is logged in for, when it is not the seller.
- `excel_rows` keys are the column lists in `src/ksefiz/operations.py` (`GENERATE_FIELDS`, `DECODE_FIELDS`, `LOOKUP_FIELDS`, `QUERY_FIELDS`).
- `<accessToken>` and similar values in angle brackets are placeholders. `"... 9 more items ..."` marks a shortened list.

## All identifiers are synthetic

NIPs fail the mod-11 checksum, so they cannot belong to a real company. KSeF numbers and IZ carry a correct CRC-8, so the offline validator accepts them. They don't exist in KSeF. `tests/test_api_examples.py` enforces both.

| Role | NIP |
|---|---|
| Seller (issues the invoices) | `3023936130` |
| Second seller (for the mixed-sellers error) | `1766801637` |
| Buyer (AP side) | `6510134498` |
| Third party with no role on the invoices | `4414286595` |

## Scenarios

| Operation | Endpoint | Files |
|---|---|---|
| Login, limits, close | `/auth/*`, `/security/public-key-certificates`, `/limits/context`, `/rate-limits` | [00-session/](00-session/) |
| M5 Generate | `POST /collective-identifiers` | [01-generate/](01-generate/): minimal, with payment, buyer generates, 71004, 71005, 71001, 71002, 21405, rejected offline |
| M3 Decode | `POST /collective-identifiers/invoices` | [02-decode/](02-decode/): with details, details hidden, no details given, paging, 21418, 21405, invalid IZ offline |
| M6 Reverse lookup | `GET /collective-identifiers/ksef/{ksefNumber}` | [03-lookup/](03-lookup/): found in 2 IZ, not in any IZ, paging, 21405 |
| M7 Period query | `POST /collective-identifiers/query` | [04-query/](04-query/): date range, all filters, paging, window over 100 days |
| Any of the above | | [05-common-errors/](05-common-errors/): 401, 403, 429 + `Retry-After`, legacy error shape, `X-System-Warning` |

## Rules shared by all four IZ endpoints

- Headers: `Authorization: Bearer <accessToken>`, `X-Error-Format: problem-details`. Errors then arrive as `application/problem+json`, with the KSeF code in `errors[].code`.
- Permission: any one of `InvoiceRead`, `InvoiceWrite`, `CollectiveIdentifierManage`.
- One shared rate-limit group `collectiveIdentifier`: 20/s, 120/min, 240/h.
- Paging: send the same request again with the `x-continuation-token` header. A `null` or missing `continuationToken` in the response means the last page. `pageSize`: decode 10..500, lookup and query 10..200.

## Known gaps

- The `details` text of 21405 validation errors comes from the KSeF validator and is not documented, so it is shown as a placeholder.
- What decode returns for a well-formed IZ that does not exist is not documented. The tool writes `no invoices returned` for any IZ that is missing from the response.
- `X-System-Warning` codes in [05-common-errors/05-system-warning-header.json](05-common-errors/05-system-warning-header.json) are made up. Only the format comes from the spec.
