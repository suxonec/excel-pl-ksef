"""M8 audit log: one JSON line per API call in logs/audit.jsonl (append-only).

Records method, path, HTTP status and duration - never request bodies or token values.
"""

from __future__ import annotations

import getpass
import json
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

DEFAULT_LOG = Path("logs") / "audit.jsonl"


class AuditLog:
    def __init__(self, env: str, operation: str, path: Path = DEFAULT_LOG) -> None:
        self.env = env
        self.operation = operation
        self.path = path
        self.requests = 0
        path.parent.mkdir(parents=True, exist_ok=True)

    def write(self, **fields: Any) -> None:
        rec = {"timestamp": datetime.now(UTC).isoformat(timespec="milliseconds"),
               "user": getpass.getuser(), "environment": self.env,
               "operation": self.operation, **fields}
        with self.path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")

    def on_request(self, method: str, path: str, status: int, seconds: float) -> None:
        self.requests += 1
        self.write(method=method, path=path, status=status, ms=round(seconds * 1000))
