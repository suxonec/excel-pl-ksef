"""The workbook builder (Python) and the macros (VBA) agree on the hidden _Sample sheet."""

import importlib.util
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _builder():  # scripts/ is not a package
    spec = importlib.util.spec_from_file_location("build_workbook",
                                                  ROOT / "scripts" / "build_workbook.py")
    assert spec and spec.loader
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_sample_sheet_matches_modsample_vba() -> None:
    bw = _builder()
    wb = bw.build(demo=False)
    ws = wb["_Sample"]
    assert ws.sheet_state == "hidden"

    vba = (ROOT / "excel" / "vba" / "modSample.bas").read_text(encoding="ascii")
    calls = re.findall(r'CopyBlock src, (\d+), "(\w+)", "(\w+)", "(\w+)"', vba)
    targets = {"1_Detect": ("A2", "line_id"), "2_Decode": ("C2", "iz"),
               "3_Match": ("A2", "doc_no"), "4_Generate": ("A2", "seller_nip")}
    assert len(calls) == len(bw.SAMPLE_BLOCKS) == 4
    for (col, fields, count_cell), (vcol, vcount, sheet, anchor) in zip(
            bw.SAMPLE_BLOCKS, calls, strict=True):
        assert (int(vcol), vcount) == (col, count_cell)
        assert len(fields) == 6  # CopyBlock copies 6 columns
        assert ws.cell(row=1, column=col).value == fields[0]
        assert ws[count_cell].value > 0
        # target sheet column under the anchor has the same header as the sample block
        exp_anchor, exp_header = targets[sheet]
        assert anchor == exp_anchor
        assert wb[sheet][anchor.replace("2", "1")].value == exp_header == fields[0]


def test_template_visible_sheets_are_empty() -> None:
    bw = _builder()
    wb = bw.build(demo=False)
    for name in ("1_Detect", "2_Decode", "3_Match", "4_Generate", "5_Lookup"):
        cells = [c.value for row in wb[name].iter_rows(min_row=2, max_col=12) for c in row
                 if c.value != bw.HELPER_NOTE]
        assert not any(cells), name
