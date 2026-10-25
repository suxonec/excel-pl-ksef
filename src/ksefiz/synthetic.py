"""Seeded synthetic test data: statements, ERP open items and a fake decode result.

NIPs are drawn so that they match the KSeF NIP pattern but deliberately FAIL the
official mod-11 NIP checksum, so a generated NIP can never belong to a real company.
"""

from __future__ import annotations

import random
from dataclasses import dataclass, field
from pathlib import Path

from ksefiz import match as m
from ksefiz.io_csv import Row, write_rows
from ksefiz.iz import KIND_IZ, KIND_KSEF, crc8, detect, nip_checksum_ok

_HEX = "0123456789ABCDEF"

STATEMENT_FIELDS = ["line_id", "date", "amount", "currency", "payer", "free_text"]
OPEN_ITEM_FIELDS = ["doc_no", "customer", "ksef_number", "amount", "currency", "due_date"]
DECODED_FIELDS = ["iz", "ksef_number", "amount", "currency", "description", "details_hidden"]
EXPECTED_DETECT_FIELDS = ["line_id", "kind", "normalised", "valid"]
EXPECTED_MATCH_FIELDS = ["iz", "ksef_number", "status"]

MANGLE_STYLES = ("plain", "lower", "nohyphen", "spaces", "linebreak", "glued")


def fake_nip(rng: random.Random) -> str:
    while True:
        d = [str(rng.randint(1, 9))] + [str(rng.randint(0, 9)) for _ in range(9)]
        nip = "".join(d)
        if nip[1:3] != "00" and not nip_checksum_ok(nip):
            return nip


def _hex(rng: random.Random, n: int) -> str:
    return "".join(rng.choice(_HEX) for _ in range(n))


def make_iz(nip: str, yyyymm: str, tech: str) -> str:
    body = f"{nip}-IZ{yyyymm}-{tech}"
    return f"{body}-{crc8(body.encode('ascii'))}"


def make_ksef(nip: str, yyyymmdd: str, tech: str) -> str:
    body = f"{nip}-{yyyymmdd}-{tech}"
    return f"{body}-{crc8(body.encode('ascii'))}"


def break_crc(id_: str) -> str:
    good = id_[-2:]
    bad = f"{(int(good, 16) + 1) % 256:02X}"
    return id_[:-2] + bad


def mangle(id_: str, style: str, rng: random.Random) -> str:
    if style == "lower":
        return id_.lower()
    if style == "nohyphen":
        return id_.replace("-", "")
    if style == "spaces":
        return id_.replace("-", " ")
    if style == "linebreak":
        i = rng.randint(1, len(id_) - 1)
        return id_[:i] + rng.choice(["\n", "\r\n", "\n  "]) + id_[i:]
    if style == "glued":
        return rng.choice(["REF:", "IZ:", "FV", "/"]) + id_ + rng.choice(["/PLN", ";", ""])
    return id_


@dataclass
class Dataset:
    statements: list[Row] = field(default_factory=list)
    open_items: list[Row] = field(default_factory=list)
    decoded: list[Row] = field(default_factory=list)
    expected_detect: list[Row] = field(default_factory=list)
    expected_match: list[Row] = field(default_factory=list)


def _noise(rng: random.Random) -> str:
    """Filler text that must not be detected as an ID."""
    while True:
        parts = [
            rng.choice(["Payment", "Zaplata za", "Przelew", "Invoice", "Settlement"]),
            f"FV/{rng.randint(2025, 2027)}/{rng.randint(1, 12):02d}/{rng.randint(1, 999)}",
            "PL" + "".join(str(rng.randint(0, 9)) for _ in range(26)),
            "".join(rng.choice(_HEX + "GHJKLMNPQRSTUVWXYZ-") for _ in range(35)),
        ]
        rng.shuffle(parts)
        text = " ".join(parts[: rng.randint(1, 4)])
        if not detect(text):
            return text


def generate(seed: int = 20261024, lines: int = 1000, invoices: int = 200) -> Dataset:
    rng = random.Random(seed)
    ds = Dataset()
    own_nip = fake_nip(rng)

    # ERP open items for our own (seller) invoices
    ksefs: list[str] = []
    amounts: dict[str, float] = {}
    for i in range(invoices):
        ksef = make_ksef(own_nip, f"2026{rng.randint(1, 12):02d}{rng.randint(1, 28):02d}",
                         _hex(rng, 12))
        amt = round(rng.uniform(50, 25000), 2)
        ksefs.append(ksef)
        amounts[ksef] = amt
        ds.open_items.append({
            "doc_no": f"FI{510000000 + i}", "customer": f"CUST-{rng.randint(1, 60):03d}",
            "ksef_number": ksef, "amount": f"{amt:.2f}",
            "currency": "EUR" if i % 17 == 0 else "PLN",
            "due_date": f"2026-{rng.randint(1, 12):02d}-{rng.randint(1, 28):02d}",
        })
    # One invoice booked twice in the ERP extract
    ds.open_items.append(dict(ds.open_items[5], doc_no="FI519999999"))

    # Decoded IZ rows: customers paying 2-6 of our invoices per IZ.
    # Every row carries the match status it must get, set by construction.
    currency = {o["ksef_number"]: o["currency"] for o in ds.open_items}
    izs: list[str] = []

    def add(iz: str, k: str, amount: str, cur: str, status: str) -> None:
        ds.decoded.append({"iz": iz, "ksef_number": k, "amount": amount, "currency": cur,
                           "description": "", "details_hidden": "false" if amount else "true"})
        ds.expected_match.append({"iz": iz, "ksef_number": k, "status": status})

    def new_iz(yyyymm: str | None = None) -> str:
        iz = make_iz(own_nip, yyyymm or f"2026{rng.randint(1, 12):02d}", _hex(rng, 12))
        izs.append(iz)
        return iz

    pool = ksefs[:150]
    rng.shuffle(pool)
    pos = 0
    while len(pool) - pos >= 2:
        n = min(rng.randint(2, 6), len(pool) - pos)
        if len(pool) - pos - n == 1:
            n += 1
        iz = new_iz()
        for k in pool[pos:pos + n]:
            amt, roll = amounts[k], rng.random()
            dup = k == ksefs[5]
            if roll < 0.06:
                add(iz, k, f"{amt - 10:.2f}", currency[k],
                    m.STATUS_DUPLICATE_ERP if dup else m.STATUS_AMOUNT_DIFF)
            elif roll < 0.10:
                add(iz, k, "", "", m.STATUS_DUPLICATE_ERP if dup else m.STATUS_NO_AMOUNT)
            else:
                add(iz, k, f"{amt:.2f}", currency[k],
                    m.STATUS_DUPLICATE_ERP if dup else m.STATUS_MATCHED)
        pos += n

    # Partial payment: ksefs[160] split over two IZ, totals agree
    split, half = ksefs[160], round(amounts[ksefs[160]] / 2, 2)
    for part, partner in ((half, ksefs[161]), (round(amounts[split] - half, 2), ksefs[162])):
        iz = new_iz("202611")
        add(iz, split, f"{part:.2f}", currency[split], m.STATUS_MATCHED)
        add(iz, partner, f"{amounts[partner]:.2f}", currency[partner], m.STATUS_MATCHED)
    # Paid twice in full over two IZ: total exceeds the open amount
    twice = ksefs[164]
    for partner in (ksefs[165], ksefs[166]):
        iz = new_iz("202612")
        add(iz, twice, f"{amounts[twice]:.2f}", currency[twice], m.STATUS_AMOUNT_DIFF)
        add(iz, partner, f"{amounts[partner]:.2f}", currency[partner], m.STATUS_MATCHED)
    # Wrong currency, and a KSeF number the ERP does not know
    iz = new_iz("202612")
    add(iz, ksefs[163], f"{amounts[ksefs[163]]:.2f}", "USD", m.STATUS_CURRENCY_DIFF)
    add(iz, make_ksef(fake_nip(rng), "20260915", _hex(rng, 12)), "100.00", "PLN",
        m.STATUS_UNMATCHED)

    # Bank statement lines quoting IZ and KSeF numbers in mangled forms
    planted = [(KIND_IZ, iz) for iz in izs] + [(KIND_KSEF, k) for k in ksefs[170:200]]
    for i in range(lines):
        line_id = f"L{i + 1:04d}"
        text_parts = [_noise(rng)]
        for _ in range(rng.choice([0, 1, 1, 1, 2])):
            kind, id_ = rng.choice(planted)
            valid = rng.random() > 0.08
            shown = id_ if valid else break_crc(id_)
            text_parts.append(mangle(shown, rng.choice(MANGLE_STYLES), rng))
            ds.expected_detect.append({"line_id": line_id, "kind": kind, "normalised": shown,
                                       "valid": "Y" if valid else "N"})
            text_parts.append(_noise(rng))
        ds.statements.append({
            "line_id": line_id, "date": f"2026-{rng.randint(1, 12):02d}-{rng.randint(1, 28):02d}",
            "amount": f"{rng.uniform(50, 50000):.2f}", "currency": "PLN",
            "payer": f"PAYER-{rng.randint(1, 60):03d}", "free_text": " ".join(text_parts),
        })
    return ds


def write(ds: Dataset, out_dir: Path) -> None:
    write_rows(out_dir / "statements_sample.csv", STATEMENT_FIELDS, ds.statements)
    write_rows(out_dir / "open_items_sample.csv", OPEN_ITEM_FIELDS, ds.open_items)
    write_rows(out_dir / "decoded_sample.csv", DECODED_FIELDS, ds.decoded)
    write_rows(out_dir / "statements_expected.csv", EXPECTED_DETECT_FIELDS, ds.expected_detect)
    write_rows(out_dir / "match_expected.csv", EXPECTED_MATCH_FIELDS, ds.expected_match)
