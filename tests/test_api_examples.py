"""docs/api-examples stays synthetic and in step with the code: valid CRCs, fake NIPs,
excel_rows columns equal to the operations.*_FIELDS lists."""

import json
import re
from pathlib import Path

import pytest

from ksefiz import operations as ops
from ksefiz.iz import check, nip_checksum_ok

ROOT = Path(__file__).parent.parent / "docs" / "api-examples"
FILES = sorted(ROOT.rglob("*.json"))
ID_RE = re.compile(r"\b\d{10}-(?:IZ\d{6}|\d{8})-[0-9A-F]{12}-[0-9A-F]{2}\b")
FIELDS = {"01-generate": ops.GENERATE_FIELDS, "02-decode": ops.DECODE_FIELDS,
          "03-lookup": ops.LOOKUP_FIELDS, "04-query": ops.QUERY_FIELDS}


def test_examples_exist() -> None:
    assert len(FILES) > 30
    assert (ROOT / "ksef-iz.http").exists()


@pytest.mark.parametrize("path", FILES, ids=lambda p: f"{p.parent.name}/{p.stem}")
def test_example(path: Path) -> None:
    text = path.read_text(encoding="utf-8")
    doc = json.loads(text)
    assert {"scenario", "source", "request", "response"} <= doc.keys()
    if doc["request"] is not None:
        assert doc["request"]["url"].startswith("https://api-test.ksef.mf.gov.pl/v2/")
        assert isinstance(doc["response"]["status"], int)

    rejected_offline = {r.get("status", "") for r in doc.get("excel_rows", [])}
    for id_ in ID_RE.findall(text):
        c = check(id_)
        assert c.valid or f"invalid IZ ({c.reason})" in rejected_offline, id_
        assert not nip_checksum_ok(id_[:10]), f"{id_}: NIP passes mod-11, use a synthetic one"
    for nip in re.findall(r'"(?:value|context_nip)": "(\d{10})"', text):
        assert not nip_checksum_ok(nip), nip

    for row in doc.get("excel_rows", []):
        assert list(row) == FIELDS[path.parent.name]


def test_http_file_ids_are_valid() -> None:
    for id_ in ID_RE.findall((ROOT / "ksef-iz.http").read_text(encoding="utf-8")):
        assert check(id_).valid, id_
