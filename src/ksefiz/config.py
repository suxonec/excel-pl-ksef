"""Environment settings. Nothing in this module touches the network."""

from __future__ import annotations

from enum import StrEnum


class Env(StrEnum):
    TEST = "test"
    DEMO = "demo"
    PROD = "prod"


BASE_URLS: dict[Env, str] = {
    Env.TEST: "https://api-test.ksef.mf.gov.pl/v2",
    Env.DEMO: "https://api-demo.ksef.mf.gov.pl/v2",
    Env.PROD: "https://api.ksef.mf.gov.pl/v2",
}

DEFAULT_ENV = Env.TEST

# Limits from open-api.json / limity.md; effective values are read at run time from
# GET /limits/context and GET /rate-limits once the API client exists.
DECODE_BATCH_SIZE = 10
DEFAULT_MAX_INVOICES_PER_IZ = 500
MAX_IZ_PER_INVOICE = 132
QUERY_MAX_DAYS = 100
