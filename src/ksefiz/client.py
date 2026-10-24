"""httpx wrapper for KSeF API 2.0: base URL per environment, problem-details errors,
429 / Retry-After handling with exponential backoff and jitter."""

from __future__ import annotations

import random
import time
from collections.abc import Callable, Mapping
from dataclasses import dataclass, field
from typing import Any

import httpx

from ksefiz import __version__
from ksefiz.config import BASE_URLS, Env

RETRY_STATUSES = (429, 502, 503, 504)
MAX_RETRIES = 5
MAX_BACKOFF_S = 60.0


@dataclass
class ApiError(Exception):
    """An error response from KSeF. `codes` are the numeric KSeF error codes (e.g. 71004)."""

    status: int
    title: str
    detail: str
    codes: list[int] = field(default_factory=list)
    trace_id: str = ""
    body: Any = None

    def __str__(self) -> str:
        codes = f" codes={self.codes}" if self.codes else ""
        trace = f" traceId={self.trace_id}" if self.trace_id else ""
        return f"HTTP {self.status} {self.title}: {self.detail}{codes}{trace}"


def parse_error(resp: httpx.Response) -> ApiError:
    """Read problem+json (preferred) or the legacy ExceptionResponse shape."""
    try:
        body: Any = resp.json()
    except ValueError:
        return ApiError(resp.status_code, resp.reason_phrase, resp.text[:500])
    if not isinstance(body, dict):
        return ApiError(resp.status_code, resp.reason_phrase, str(body)[:500], body=body)

    codes: list[int] = []
    details: list[str] = []
    for e in body.get("errors") or []:
        if isinstance(e, dict):
            if isinstance(e.get("code"), int):
                codes.append(e["code"])
            details.append(str(e.get("description", "")))
            details.extend(str(d) for d in e.get("details") or [])
    exc = body.get("exception")
    if isinstance(exc, dict):
        for e in exc.get("exceptionDetailList") or []:
            if isinstance(e.get("exceptionCode"), int):
                codes.append(e["exceptionCode"])
            details.append(str(e.get("exceptionDescription", "")))
            details.extend(str(d) for d in e.get("details") or [])
    status_info = body.get("status")
    if isinstance(status_info, dict):
        details.extend(str(d) for d in status_info.get("details") or [])

    detail = str(body.get("detail") or "")
    if details:
        detail = (detail + " | " if detail else "") + "; ".join(d for d in details if d)
    trace = str(body.get("traceId") or (exc or {}).get("serviceCode") or "")
    if not trace:
        trace = resp.headers.get("x-request-id", "")
    return ApiError(resp.status_code, str(body.get("title") or resp.reason_phrase),
                    detail, codes, trace, body)


def retry_delay(resp: httpx.Response | None, attempt: int) -> float:
    if resp is not None:
        ra = resp.headers.get("Retry-After")
        if ra and ra.strip().isdigit():
            return min(float(ra), MAX_BACKOFF_S * 5)
    base: float = min(MAX_BACKOFF_S, 2.0 ** attempt)
    return base / 2 + random.uniform(0, base / 2)


class KsefClient:
    def __init__(
        self,
        env: Env,
        *,
        timeout: float = 30.0,
        transport: httpx.BaseTransport | None = None,
        sleep: Callable[[float], None] = time.sleep,
        on_request: Callable[[str, str, int, float], None] | None = None,
    ) -> None:
        self.env = env
        self._sleep = sleep
        self._on_request = on_request
        self._http = httpx.Client(
            base_url=BASE_URLS[env],
            timeout=timeout,
            transport=transport,
            headers={
                "X-Error-Format": "problem-details",
                "Accept": "application/json",
                "User-Agent": f"ksefiz/{__version__}",
            },
        )

    def close(self) -> None:
        self._http.close()

    def __enter__(self) -> KsefClient:
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()

    def request(
        self,
        method: str,
        path: str,
        *,
        json: Any = None,
        content: bytes | None = None,
        content_type: str | None = None,
        params: Mapping[str, Any] | None = None,
        bearer: str | None = None,
        extra_headers: Mapping[str, str] | None = None,
    ) -> Any:
        """Send one request; returns parsed JSON ({} for an empty body). Raises ApiError."""
        headers: dict[str, str] = dict(extra_headers or {})
        if bearer:
            headers["Authorization"] = f"Bearer {bearer}"
        if content_type:
            headers["Content-Type"] = content_type
        resp: httpx.Response | None = None
        for attempt in range(MAX_RETRIES + 1):
            start = time.monotonic()
            try:
                resp = self._http.request(method, path, json=json, content=content,
                                          params=params, headers=headers)
            except httpx.TransportError:
                if attempt == MAX_RETRIES:
                    raise
                self._sleep(retry_delay(None, attempt))
                continue
            if self._on_request:
                self._on_request(method, path, resp.status_code, time.monotonic() - start)
            if resp.status_code in RETRY_STATUSES and attempt < MAX_RETRIES:
                self._sleep(retry_delay(resp, attempt))
                continue
            break
        assert resp is not None
        if resp.status_code >= 400:
            raise parse_error(resp)
        if not resp.content:
            return {}
        return resp.json()
