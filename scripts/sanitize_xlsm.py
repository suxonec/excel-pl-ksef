"""Remove machine and user details that Excel writes into a workbook on save, so the
release workbook can be published:

- docProps/core.xml: dc:creator and cp:lastModifiedBy -> "ksefiz"
- docProps/app.xml:  Company / Manager -> empty
- xl/workbook.xml:   x15ac:absPath (the folder it was last saved in, incl. the user name)

    python scripts/sanitize_xlsm.py excel/dist/KSeF_IZ_Tool.xlsm

Rewrites only those XML parts; every other part (including vbaProject.bin and the button
drawings) is copied byte for byte. scripts/verify_xlsm_vba.py fails if any of this is present.
"""

from __future__ import annotations

import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

NEUTRAL = "ksefiz"
ABS_PATH = re.compile(
    r'<mc:AlternateContent[^>]*>\s*<mc:Choice Requires="x15">\s*<x15ac:absPath\b[^>]*/>\s*'
    r"</mc:Choice>\s*</mc:AlternateContent>|<x15ac:absPath\b[^>]*/>")


def personal_details(parts: dict[str, bytes]) -> list[str]:
    """What still identifies a person or machine (empty list = clean)."""
    found = []
    core = parts.get("docProps/core.xml", b"").decode("utf-8", "replace")
    for tag in ("dc:creator", "cp:lastModifiedBy"):
        m = re.search(rf"<{tag}>(.*?)</{tag}>", core, re.S)
        if m and m.group(1).strip() not in ("", NEUTRAL):
            found.append(f"{tag} = {m.group(1).strip()!r}")
    app = parts.get("docProps/app.xml", b"").decode("utf-8", "replace")
    for tag in ("Company", "Manager"):
        m = re.search(rf"<{tag}>(.*?)</{tag}>", app, re.S)
        if m and m.group(1).strip():
            found.append(f"{tag} = {m.group(1).strip()!r}")
    if "absPath" in parts.get("xl/workbook.xml", b"").decode("utf-8", "replace"):
        found.append("xl/workbook.xml contains x15ac:absPath (last save folder)")
    return found


def _clean(name: str, data: bytes) -> bytes:
    text = data.decode("utf-8")
    if name == "docProps/core.xml":
        for tag in ("dc:creator", "cp:lastModifiedBy"):
            text = re.sub(rf"<{tag}>.*?</{tag}>", f"<{tag}>{NEUTRAL}</{tag}>", text, flags=re.S)
    elif name == "docProps/app.xml":
        for tag in ("Company", "Manager"):
            text = re.sub(rf"<{tag}>.*?</{tag}>", f"<{tag}></{tag}>", text, flags=re.S)
    elif name == "xl/workbook.xml":
        text = ABS_PATH.sub("", text)
    return text.encode("utf-8")


def sanitize(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as zin:
        infos = zin.infolist()
        parts = {i.filename: zin.read(i.filename) for i in infos}
    before = personal_details(parts)
    targets = {"docProps/core.xml", "docProps/app.xml", "xl/workbook.xml"}
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm", dir=path.parent) as tmp:
        tmp_path = Path(tmp.name)
    with zipfile.ZipFile(tmp_path, "w") as zout:
        for info in infos:
            data = parts[info.filename]
            if info.filename in targets:
                data = _clean(info.filename, data)
            zout.writestr(info, data, compress_type=info.compress_type)
    with zipfile.ZipFile(tmp_path) as check:
        left = personal_details({n: check.read(n) for n in check.namelist()})
        if left or check.testzip() is not None:
            tmp_path.unlink()
            raise SystemExit(f"sanitising failed: {left}")
    shutil.move(str(tmp_path), path)
    return before


if __name__ == "__main__":
    target = Path(sys.argv[1])
    removed = sanitize(target)
    print(f"{target}: " + ("removed " + "; ".join(removed) if removed else "already clean"))
