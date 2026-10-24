"""Client-side throttle for one KSeF rate-limit group (e.g. `collectiveIdentifier`, shared by
all four collective endpoints): per-second, per-minute and per-hour sliding windows, all
enforced together. Timestamps can be persisted so separate CLI runs share the hourly budget.
"""

from __future__ import annotations

import json
import time
from collections import deque
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from pathlib import Path

WINDOWS_S = {"perSecond": 1.0, "perMinute": 60.0, "perHour": 3600.0}
COLLECTIVE_DEFAULTS = {"perSecond": 20, "perMinute": 120, "perHour": 240}


@dataclass(frozen=True)
class Budget:
    per_second: int
    per_minute: int
    per_hour: int


class RateLimiter:
    def __init__(
        self,
        limits: Mapping[str, int] = COLLECTIVE_DEFAULTS,
        *,
        state_path: Path | None = None,
        clock: Callable[[], float] = time.time,
        sleep: Callable[[float], None] = time.sleep,
    ) -> None:
        self.limits = {k: int(limits[k]) for k in WINDOWS_S}
        self._clock = clock
        self._sleep = sleep
        self._state_path = state_path
        self._stamps: deque[float] = deque()
        if state_path and state_path.exists():
            try:
                saved = json.loads(state_path.read_text(encoding="utf-8"))
                self._stamps.extend(sorted(float(t) for t in saved))
            except (ValueError, TypeError):
                pass
        self._prune(self._clock())

    def _prune(self, now: float) -> None:
        horizon = now - WINDOWS_S["perHour"]
        while self._stamps and self._stamps[0] <= horizon:
            self._stamps.popleft()

    def _count_since(self, since: float) -> int:
        return sum(1 for t in self._stamps if t > since)

    def wait_time(self, now: float | None = None) -> float:
        """Seconds until one more request fits every window (0 = now)."""
        now = self._clock() if now is None else now
        self._prune(now)
        wait = 0.0
        for name, span in WINDOWS_S.items():
            limit = self.limits[name]
            inside = [t for t in self._stamps if t > now - span]
            if len(inside) >= limit:
                # oldest stamp that must leave the window before we fit
                wait = max(wait, inside[len(inside) - limit] + span - now)
        return wait

    def acquire(self) -> None:
        while (w := self.wait_time()) > 0:
            self._sleep(w + 0.01)
        self._stamps.append(self._clock())
        self._save()

    def remaining(self) -> Budget:
        now = self._clock()
        self._prune(now)
        return Budget(
            self.limits["perSecond"] - self._count_since(now - 1),
            self.limits["perMinute"] - self._count_since(now - 60),
            self.limits["perHour"] - self._count_since(now - 3600),
        )

    def _save(self) -> None:
        if self._state_path:
            self._state_path.parent.mkdir(parents=True, exist_ok=True)
            self._state_path.write_text(json.dumps(list(self._stamps)), encoding="utf-8")
