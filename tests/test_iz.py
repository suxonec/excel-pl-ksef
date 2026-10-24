import csv
import random
import string
import time
from pathlib import Path

import pytest

from ksefiz import synthetic
from ksefiz.iz import (
    KIND_IZ,
    KIND_KSEF,
    check,
    crc8,
    detect,
    is_valid_iz,
    is_valid_ksef,
    normalise,
)

HERE = Path(__file__).parent
SPEC_IZ = "1111111111-IZ202607-65ED02180000-E7"  # open-api.json example
SPEC_KSEF = "5265877635-20250826-0100001AF629-AF"  # faktury/numer-ksef.md example


def _vectors(name: str) -> list[dict[str, str]]:
    with (HERE / name).open(encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


def test_crc8_spec_examples() -> None:
    assert crc8(SPEC_IZ[:32].encode()) == "E7"
    assert crc8(SPEC_KSEF[:32].encode()) == "AF"


def test_spec_examples_valid() -> None:
    assert is_valid_iz(SPEC_IZ)
    assert is_valid_ksef(SPEC_KSEF)
    assert not is_valid_ksef(SPEC_IZ)
    assert not is_valid_iz(SPEC_KSEF)


@pytest.mark.parametrize("row", _vectors("vectors.csv"), ids=lambda r: repr(r["input"])[:40])
def test_check_vectors(row: dict[str, str]) -> None:
    c = check(row["input"])
    assert c.kind == row["kind"]
    assert c.normalised == row["normalised"]
    assert ("Y" if c.crc_ok else "N") == row["crc_ok"]
    assert ("Y" if c.valid else "N") == row["valid"]
    assert c.reason == row["reason"]


@pytest.mark.parametrize("row", _vectors("vectors_detect.csv"), ids=lambda r: repr(r["text"])[:40])
def test_detect_vectors(row: dict[str, str]) -> None:
    found = "|".join(
        f"{d.check.normalised}:{'Y' if d.check.valid else 'N'}" for d in detect(row["text"])
    )
    assert found == row["found"]


@pytest.mark.parametrize("style", synthetic.MANGLE_STYLES)
def test_every_mangling_style_normalises(style: str) -> None:
    rng = random.Random(style)
    for _ in range(200):
        nip = synthetic.fake_nip(rng)
        for id_ in (synthetic.make_iz(nip, "202611", "ABCDEF012345"),
                    synthetic.make_ksef(nip, "20261103", "0123456789AB")):
            hits = detect(f"x {synthetic.mangle(id_, style, rng)} y")
            assert [h.check.normalised for h in hits] == [id_]
            assert hits[0].check.valid


def test_normalise_rejects_wrong_length() -> None:
    assert normalise(SPEC_IZ + "0") == ""
    assert normalise(SPEC_IZ[:-1]) == ""


def test_no_false_positives_on_random_35_char_strings() -> None:
    rng = random.Random(1)
    alphabet = string.ascii_uppercase + string.digits + "-"
    valid = 0
    for _ in range(10_000):
        s = "".join(rng.choice(alphabet) for _ in range(35))
        valid += sum(d.check.valid for d in detect(s))
        valid += check(s).valid
    assert valid == 0


def test_no_false_positives_on_random_shaped_ids() -> None:
    """Right shape but random check digits: only the CRC stands between us and a hit."""
    rng = random.Random(2)
    accepted = 0
    for _ in range(5_000):
        id_ = synthetic.make_iz(synthetic.fake_nip(rng), "202610", "0" * 12)
        bad = synthetic.break_crc(id_)
        accepted += check(bad).valid
    assert accepted == 0


def test_statement_1000_lines_matches_expected_and_is_fast() -> None:
    ds = synthetic.generate()
    start = time.perf_counter()
    got = [
        (s["line_id"], d.check.normalised, "Y" if d.check.valid else "N")
        for s in ds.statements
        for d in detect(s["free_text"])
    ]
    elapsed = time.perf_counter() - start
    expected = [(e["line_id"], e["normalised"], e["valid"]) for e in ds.expected_detect]
    assert got == expected
    assert elapsed < 5.0


def test_generated_nips_can_never_be_real() -> None:
    rng = random.Random(3)
    for _ in range(1_000):
        nip = synthetic.fake_nip(rng)
        assert not synthetic.nip_checksum_ok(nip)
        assert check(synthetic.make_iz(nip, "202601", "F" * 12)).kind == KIND_IZ
        assert check(synthetic.make_ksef(nip, "20260101", "F" * 12)).kind == KIND_KSEF
