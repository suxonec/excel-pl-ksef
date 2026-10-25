"""Prove that a macro workbook contains exactly the VBA in excel/vba/ and nothing else.

    python scripts/verify_xlsm_vba.py excel/dist/KSeF_IZ_Tool.xlsm

- every excel/vba/*.bas module must be present with identical code,
- the workbook document module (ThisWorkbook; localised names such as "Ten_skoroszyt"
  are recognised by type) must equal excel/vba/ThisWorkbook.vba,
- sheet modules and any other module must be empty (at most `Option Explicit`),
- no user or machine details in the file metadata (run scripts/sanitize_xlsm.py first).

Comparison ignores `Attribute` lines, trailing spaces, line endings and leading/trailing
blank lines. Identifier case differences are reported but accepted, because the VBA editor
re-cases identifiers project-wide. Exit code 1 on any other difference.
"""

from __future__ import annotations

import difflib
import re
import sys
import zipfile
from pathlib import Path

from oletools.olevba import VBA_Parser

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sanitize_xlsm import personal_details  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
WORKBOOK_CLASS = "00020819"  # VB_Base of the workbook document module
SHEET_CLASS = "00020820"


def normalise(code: str) -> str:
    lines = [ln.rstrip() for ln in code.replace("\r\n", "\n").replace("\r", "\n").split("\n")
             if not ln.startswith("Attribute ")]
    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines)


def expected_modules() -> tuple[dict[str, str], str]:
    mods = {}
    for p in sorted((ROOT / "excel" / "vba").glob("*.bas")):
        text = p.read_text(encoding="ascii")
        m = re.search(r'^Attribute VB_Name = "([^"]+)"', text, re.M)
        mods[m.group(1) if m else p.stem] = normalise(text)
    workbook = normalise((ROOT / "excel" / "vba" / "ThisWorkbook.vba").read_text(encoding="ascii"))
    return mods, workbook


def main(path: Path) -> int:
    expected, expected_wb = expected_modules()
    seen: set[str] = set()
    problems: list[str] = []
    notes: list[str] = []

    def compare(label: str, want: str, got: str) -> None:
        if got == want:
            return
        if got.lower() == want.lower():
            notes.append(f"{label}: identical except identifier case (VBA editor re-casing)")
            return
        diff = list(difflib.unified_diff(want.split("\n"), got.split("\n"), "repo", "workbook",
                                         n=1, lineterm=""))
        problems.append(f"{label}: code differs from the repository\n  " +
                        "\n  ".join(diff[:20]))

    parser = VBA_Parser(str(path))
    try:
        for _file, _stream, vba_filename, code in parser.extract_macros():
            name = Path(vba_filename).stem
            base = re.search(r'Attribute VB_Base = "0\{([0-9A-F]{8})', code)
            kind = base.group(1) if base else ""
            body = normalise(code)
            if name in expected:
                seen.add(name)
                compare(f"module {name}", expected[name], body)
            elif kind == WORKBOOK_CLASS:
                seen.add("<workbook>")
                compare(f"workbook module {name}", expected_wb, body)
            elif body not in ("", "Option Explicit"):
                what = "sheet module" if kind == SHEET_CLASS else "unexpected module"
                problems.append(f"{what} {name} contains code (must be empty)")
    finally:
        parser.close()

    with zipfile.ZipFile(path) as z:
        for detail in personal_details({n: z.read(n) for n in z.namelist()}):
            problems.append(f"personal/machine detail: {detail} - run scripts/sanitize_xlsm.py")

    for name in sorted(set(expected) - seen):
        problems.append(f"module {name} is missing from the workbook")
    if "<workbook>" not in seen:
        problems.append("workbook document module not found")

    for n in notes:
        print("note:", n)
    if problems:
        print(f"FAIL {path}:")
        for p in problems:
            print(" -", p)
        return 1
    print(f"OK {path}: {len(expected)} modules + workbook module match excel/vba/, "
          "no other code")
    return 0


if __name__ == "__main__":
    sys.exit(main(Path(sys.argv[1])))
