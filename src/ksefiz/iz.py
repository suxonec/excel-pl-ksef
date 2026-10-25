"""Offline detection, normalisation and CRC-8 validation of IZ and KSeF numbers.

Formats (open-api.json, faktury/numer-ksef.md):
    IZ    NNNNNNNNNN-IZYYYYMM-HHHHHHHHHHHH-CC
    KSeF  NNNNNNNNNN-YYYYMMDD-HHHHHHHHHHHH-CC
Both are 35 characters and end in a CRC-8 (poly 0x07, init 0x00, no reflection,
no final XOR) over the first 32 characters, hyphens included.

The tolerant detection rules here are mirrored one-to-one by the VBA scanner in
excel/vba/modIZ.bas; tests/vectors*.csv keep the two in step. Change both together.
"""

from __future__ import annotations

import re
import string
from collections.abc import Iterator
from dataclasses import dataclass

# Characters that bank statements put inside or between ID parts.
SEPARATORS = "- \t\r\n\xa0"
# ASCII-only uppercasing: VBA mirror cannot match Unicode case rules (e.g. "ß" -> "SS")
_COMPACT_TABLE = str.maketrans(string.ascii_lowercase, string.ascii_uppercase, SEPARATORS)

_H = r"[- \t\xa0]"  # in-line separator
_BR = r"(?:\r\n|\r|\n)[ \t\xa0]*"  # line break plus indentation
_GAP = rf"{_H}{{0,3}}(?:{_BR}{_H}{{0,3}})?"  # between two ID parts


def _segment(classes: list[str]) -> str:
    """Characters of one ID part; a line break may split the part anywhere."""
    return f"(?:{_BR})?".join(classes)


_D, _X = r"\d", r"[0-9A-F]"
_IZ_TOLERANT = _GAP.join(
    [
        _segment([_D] * 10),
        _segment(["I", "Z"] + [_D] * 6),
        _segment([_X] * 12),
        _segment([_X] * 2),
    ]
)
_KSEF_TOLERANT = _GAP.join(
    [
        _segment([_D] * 10),
        _segment([_D] * 8),
        _segment([_X] * 6),  # KSeF 1.0 36-char form has a hyphen here
        _segment([_X] * 6),
        _segment([_X] * 2),
    ]
)
DETECT_RE = re.compile(
    rf"(?<![0-9])(?:(?P<iz>{_IZ_TOLERANT})|(?P<ksef>{_KSEF_TOLERANT}))(?![0-9A-F])",
    re.IGNORECASE | re.ASCII,
)

IZ_RE = re.compile(r"^\d{10}-IZ\d{6}-[0-9A-F]{12}-[0-9A-F]{2}$", re.ASCII)
KSEF_RE = re.compile(r"^\d{10}-\d{8}-[0-9A-F]{12}-[0-9A-F]{2}$", re.ASCII)
NIP_RE = re.compile(r"^[1-9](?:\d[1-9]|[1-9]\d)\d{7}$", re.ASCII)

KIND_IZ = "IZ"
KIND_KSEF = "KSEF"

REASON_OK = "OK"
REASON_FORMAT = "format"
REASON_NIP = "nip"
REASON_DATE = "date"
REASON_CRC = "crc"


@dataclass(frozen=True)
class IdCheck:
    kind: str  # KIND_IZ, KIND_KSEF or "" when the shape is unrecognisable
    normalised: str  # canonical 35-char form, "" when not 32 significant chars
    crc_ok: bool
    valid: bool
    reason: str


@dataclass(frozen=True)
class Detection:
    start: int
    end: int
    raw: str
    check: IdCheck


_NIP_WEIGHTS = (6, 5, 7, 2, 3, 4, 5, 6, 7)


def nip_checksum_ok(nip: str) -> bool:
    """Official mod-11 NIP checksum (real NIPs pass; TEST and synthetic NIPs need not)."""
    if len(nip) != 10 or not nip.isascii() or not nip.isdigit():
        return False
    s = sum(int(d) * w for d, w in zip(nip[:9], _NIP_WEIGHTS, strict=True)) % 11
    return s != 10 and s == int(nip[9])


def crc8(data: bytes) -> str:
    c = 0
    for x in data:
        c ^= x
        for _ in range(8):
            c = ((c << 1) ^ 0x07) & 0xFF if c & 0x80 else (c << 1) & 0xFF
    return f"{c:02X}"


def compact(s: str) -> str:
    return s.translate(_COMPACT_TABLE)


def normalise(s: str) -> str:
    """Canonical hyphenated form, or "" if the input is not 32 significant characters."""
    c = compact(s)
    if len(c) != 32:
        return ""
    return f"{c[:10]}-{c[10:18]}-{c[18:30]}-{c[30:]}"


def _date_ok(kind: str, n: str) -> bool:
    year = int(n[13:17]) if kind == KIND_IZ else int(n[11:15])
    month = int(n[17:19]) if kind == KIND_IZ else int(n[15:17])
    if year < 2020 or not 1 <= month <= 12:
        return False
    return kind == KIND_IZ or 1 <= int(n[17:19]) <= 31


def check(s: str) -> IdCheck:
    """Classify and validate one IZ or KSeF number, tolerating separators and case."""
    n = normalise(s)
    if not n:
        return IdCheck("", "", False, False, REASON_FORMAT)
    kind = KIND_IZ if n[11:13] == "IZ" else KIND_KSEF
    if not (IZ_RE if kind == KIND_IZ else KSEF_RE).match(n):
        return IdCheck(kind, n, False, False, REASON_FORMAT)
    crc_ok = crc8(n[:32].encode("ascii")) == n[33:]
    if not NIP_RE.match(n[:10]):
        reason = REASON_NIP
    elif not _date_ok(kind, n):
        reason = REASON_DATE
    elif not crc_ok:
        reason = REASON_CRC
    else:
        reason = REASON_OK
    return IdCheck(kind, n, crc_ok, reason == REASON_OK, reason)


def is_valid_iz(s: str) -> bool:
    r = check(s)
    return r.valid and r.kind == KIND_IZ


def is_valid_ksef(s: str) -> bool:
    r = check(s)
    return r.valid and r.kind == KIND_KSEF


def iter_detect(text: str) -> Iterator[Detection]:
    """Every IZ / KSeF-shaped token in free text, valid or not (CRC decoys included)."""
    for m in DETECT_RE.finditer(text):
        yield Detection(m.start(), m.end(), m.group(0), check(m.group(0)))


def detect(text: str) -> list[Detection]:
    return list(iter_detect(text))
