"""Layout regression for the setup assistant (buttons were cut off on a 150 % Windows screen
in 0.2.0). At every display scaling the Back / Next / Cancel buttons must stay fully
visible. Every page must also fit completely unless the window had to be capped to a
screen too small for that scaling (e.g. 200 % on a 1024x768 CI desktop). Skipped where no
display is available."""

from pathlib import Path

import pytest

tk = pytest.importorskip("tkinter")


@pytest.fixture(autouse=True)
def home(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("KSEFIZ_HOME", str(tmp_path))


@pytest.mark.parametrize("percent", [100, 150, 200])
def test_every_page_fits_and_buttons_visible(percent: int) -> None:
    from ksefiz import setup_gui

    try:
        w = setup_gui.Wizard(scale=percent / 100 * 96 / 72)
    except tk.TclError as e:  # no display (e.g. Linux CI)
        pytest.skip(f"no display: {e}")
    try:
        w.geometry("+0+0")
        w.update()
        for page in w.pages:
            if isinstance(page, setup_gui.Environment):
                w.env.set("prod")  # longest variant: PROD warning shown
                page._changed()
            for q in w.pages:
                q.pack_forget()
            page.pack(fill="both", expand=True)
            w.update()
            btn_bottom = w.next_btn.winfo_rooty() + w.next_btn.winfo_height()
            assert btn_bottom <= w.winfo_rooty() + w.winfo_height(), page.title
            assert w.next_btn.winfo_height() >= w.next_btn.winfo_reqheight(), page.title
            screen_limited = w.winfo_height() >= w.winfo_screenheight() - 80 - 2
            if not screen_limited:
                assert page.winfo_reqheight() <= w.body.winfo_height(), page.title
            if isinstance(page, setup_gui.Environment):
                w.env.set("test")
                page._changed()
    finally:
        w.destroy()


def test_window_stays_hidden_while_measuring(monkeypatch: pytest.MonkeyPatch) -> None:
    """The placeholder texts used for sizing must never be painted on screen."""
    from ksefiz import setup_gui

    states: list[str] = []
    original = setup_gui.Wizard._fit_window

    def spy(self: setup_gui.Wizard) -> None:
        states.append(self.state())
        original(self)
        states.append(self.state())

    monkeypatch.setattr(setup_gui.Wizard, "_fit_window", spy)
    try:
        w = setup_gui.Wizard()
    except tk.TclError as e:
        pytest.skip(f"no display: {e}")
    try:
        assert states == ["withdrawn", "withdrawn"]
        w.update()
        assert w.state() == "normal"
        assert w.pages[0].winfo_ismapped() and w.h1.cget("text") == "Welcome"
        texts = [c.cget("text") for p in w.pages for c in p.winfo_children()
                 if isinstance(c, tk.ttk.Label)]
        assert not any("XXXX" in t.replace(" ", "") or "Environment: X" in t for t in texts)
    finally:
        w.destroy()
