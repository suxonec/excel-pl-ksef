from pathlib import Path

from ksefiz.ratelimit import RateLimiter


class FakeClock:
    def __init__(self) -> None:
        self.t = 1_000_000.0
        self.slept: list[float] = []

    def __call__(self) -> float:
        return self.t

    def sleep(self, s: float) -> None:
        self.slept.append(s)
        self.t += s


def _limiter(clock: FakeClock, path: Path | None = None, **limits: int) -> RateLimiter:
    lim = {"perSecond": 20, "perMinute": 120, "perHour": 240, **limits}
    return RateLimiter(lim, state_path=path, clock=clock, sleep=clock.sleep)


def test_per_second_window() -> None:
    c = FakeClock()
    rl = _limiter(c)
    for _ in range(20):
        rl.acquire()
    assert c.slept == []
    rl.acquire()  # 21st within the same second must wait ~1 s
    assert 0.9 < sum(c.slept) < 1.1


def test_hourly_budget_is_the_binding_limit() -> None:
    c = FakeClock()
    rl = _limiter(c)
    for _ in range(240):
        rl.acquire()
    assert rl.remaining().per_hour == 0
    assert 60 < c.t - 1_000_000.0 < 70  # 2 x 120 per minute -> one minute wait
    before = c.t
    rl.acquire()
    assert c.t - before > 0  # had to wait for the first stamp to leave the hour


def test_all_windows_enforced_together() -> None:
    c = FakeClock()
    rl = _limiter(c, perSecond=100, perMinute=5, perHour=1000)
    for _ in range(6):
        rl.acquire()
    assert 59 < c.t - 1_000_000.0 < 61


def test_budget_survives_restart(tmp_path: Path) -> None:
    c = FakeClock()
    path = tmp_path / "rl.json"
    rl = _limiter(c, path)
    for _ in range(10):
        rl.acquire()
    again = _limiter(c, path)
    assert again.remaining().per_hour == 230
    c.t += 3601
    assert _limiter(c, path).remaining().per_hour == 240


def test_corrupt_state_is_ignored(tmp_path: Path) -> None:
    path = tmp_path / "rl.json"
    path.write_text("not json", encoding="utf-8")
    assert _limiter(FakeClock(), path).remaining().per_hour == 240
