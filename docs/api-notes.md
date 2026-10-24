# KSeF API notes

Checked against https://github.com/CIRFMF/ksef-api `open-api.json` (API 2.9.0, build 2.9.0-te) and
`faktury/numer-ksef.md` on 2026-10-24. **`open-api.json` wins over this file.**

## Identifier formats

| ID | Format | Length |
|---|---|---|
| IZ | `NNNNNNNNNN-IZYYYYMM-HHHHHHHHHHHH-CC` | 35 |
| KSeF number | `NNNNNNNNNN-YYYYMMDD-HHHHHHHHHHHH-CC` | 35 (spec still accepts the 36-char KSeF 1.0 form `…-HHHHHH-HHHHHH-CC`) |

- `CC` for both: CRC-8, poly `0x07`, init `0x00`, over the first 32 chars incl. hyphens, uppercase hex.
  Verified: IZ example `1111111111-IZ202607-65ED02180000` -> `E7`; KSeF example
  `5265877635-20250826-0100001AF629` -> `AF`.
- NIP schema pattern: `^[1-9]((\d[1-9])|([1-9]\d))\d{7}$`.
- KsefNumber schema: year `>= 2020`, month `01-12`, day `01-31`.
- `GenerateCollectiveIdentifierResponse` describes the IZ NIP as "NIP sprzedawcy" (seller NIP). See open question 1.

## Collective identifier endpoints

| Operation | Endpoint | Notes |
|---|---|---|
| Generate | `POST /collective-identifiers` | `invoices[]` min 2, default max 500 (`limity.md`); `description` max 512 |
| Decode | `POST /collective-identifiers/invoices` | `collectiveIdentifierNumbers` max 10; paged via `continuationToken` |
| Reverse lookup | `GET /collective-identifiers/ksef/{ksefNumber}` | paged |
| Period query | `POST /collective-identifiers/query` | `dateCreatedFrom/To` max 100 days; paged |
| Context limits | `GET /limits/context` | `CollectiveIdentifierEffectiveContextLimits.maxInvoices`; override max 5,000 |
| Rate limits | `GET /rate-limits` | |

- One invoice: max 132 IZ per context.
- Error codes used: `71001`, `71004` (mixed sellers), `71005` (duplicate KSeF numbers), `21418` (bad continuation token).
- Rate limits, collective endpoints: 20/s, 120/min, 240/h. `POST /auth/ksef-token`: 60/s.
- Send `X-Error-Format: problem-details`; errors arrive as `application/problem+json`.

## Authentication (verified live on TEST 2026-10-24)

- `GET /auth/{ref}` status codes: `100` in progress, `200` success, `415` no permissions, `425` revoked,
  `450` bad token / challenge / encryption (see `details`). Poll until the code is not 100.
- `POST /auth/xades-signature?verifyCertificateChain=false`, `Content-Type: application/xml`. Self-signed seals work on TEST only.
- Token encryption key: the `GET /security/public-key-certificates` entry whose `usage` contains `KsefTokenEncryption` and that is valid now. Send its `publicKeyId`.
- `POST /tokens` (10/s, 30/min, 120/h) returns the token **once**. Poll `GET /tokens/{ref}` until `Active`.
- `GET /limits/context` -> `{"collectiveIdentifier": {"maxInvoices": 500}}` for a fresh TEST context.
- Errors: `application/problem+json` with `errors[].code`, or the legacy `exception.exceptionDetailList[].exceptionCode`. 429 carries `Retry-After` in seconds.

## Online session (invoice sending, verified live on TEST 2026-10-24)

- `POST /sessions/online` `{formCode: {systemCode: "FA (3)", schemaVersion: "1-0E", value: "FA"}, encryption: {encryptedSymmetricKey, initializationVector, publicKeyId}}`.
- `POST /sessions/online/{ref}/invoices` takes the SHA-256 (Base64) and size of both the plain and the encrypted invoice, plus the Base64 ciphertext. Then `POST .../close`.
- `GET /sessions/{ref}/invoices/{invoiceRef}` codes: `100`/`150` processing, `200` success (`ksefNumber`), `440` duplicate (seller NIP + type + `P_2`), `450` semantic error.
- The NIP checksum is validated on PROD only. `P_1` must not be after the KSeF acceptance date.
- KSeF numbers issued by TEST pass the offline CRC-8 check (`iz.is_valid_ksef`).

## Collective identifiers (verified live on TEST 2026-10-24)

- One shared rate-limit group `collectiveIdentifier` (20/s, 120/min, 240/h) covers generate, decode, lookup and query.
- Paging: request header `x-continuation-token`, response `continuationToken` (null or absent = last page). `pageSize`: decode 10..500, lookup/query 10..200 (above that: 400 / 21405).
- Generate returns 201 `{collectiveIdentifierNumber}`. The IZ prefix is the seller NIP, even when the buyer generates.
- The API itself returns 71004 for mixed sellers and 71005 for duplicate KSeF numbers.
- Decode as either party shows payment and description (`detailsHidden=false`).
- Lookup returns IZ from all contexts that include the invoice (`createdInCurrentContext` tells them apart), newest first.
