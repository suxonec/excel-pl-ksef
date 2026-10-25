"""KSeF IZ Tool setup assistant (Tkinter): environment -> company NIP -> KSeF token ->
connection test -> save (OS credential store + ~/.ksefiz/excel.ini) -> open the workbook.

Started by the Windows installer at the end of setup, and from the Start menu afterwards
("KSeF IZ Tool setup") to replace a token or switch environment. Logic lives in
userconfig.py; this module is only the UI.
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from collections.abc import Callable
from tkinter import messagebox, ttk
from typing import Any

from ksefiz import __version__, testenv, userconfig
from ksefiz.client import KsefClient
from ksefiz.config import Env
from ksefiz.credentials import get_token, load_context

ENV_TEXT = {
    Env.TEST: ("TEST", "Practice environment of the Ministry of Finance. No legal effect; "
                       "use for training and trying the tool."),
    Env.DEMO: ("DEMO", "Pre-production environment with production-like rules. Needs a real "
                       "company token."),
    Env.PROD: ("PROD", "Production. Generated collective identifiers are real records with "
                       "legal effect. Use only after sign-off."),
}


class Page(ttk.Frame):
    title = ""
    subtitle = ""

    def __init__(self, wiz: Wizard) -> None:
        super().__init__(wiz.body, padding=(wiz.px(24), wiz.px(16)))
        self.wiz = wiz

    def skip(self) -> bool:
        return False

    def on_show(self) -> None:
        pass

    def validate(self) -> str | None:
        return None

    def text(self, t: str, **kw: Any) -> ttk.Label:
        lbl = ttk.Label(self, text=t, wraplength=self.wiz.px(560), justify="left", **kw)
        lbl.pack(anchor="w", pady=(0, self.wiz.px(8)))
        return lbl


class Welcome(Page):
    title = "Welcome"
    subtitle = "Set up the KSeF IZ Tool on this computer"

    def __init__(self, wiz: Wizard) -> None:
        super().__init__(wiz)
        self.text("This assistant connects the Excel workbook to KSeF in four short steps:")
        for t in ("1.  choose the KSeF environment,", "2.  enter your company NIP,",
                  "3.  paste the KSeF token you received (it is stored in Windows Credential "
                  "Manager, never in the workbook),", "4.  test the connection."):
            self.text(t)
        self.text("You can run this assistant again at any time from the Start menu "
                  "(KSeF IZ Tool setup), e.g. to replace an expired token.")


class Environment(Page):
    title = "KSeF environment"
    subtitle = "Which KSeF system should the workbook use?"

    def __init__(self, wiz: Wizard) -> None:
        super().__init__(wiz)
        for env, (name, desc) in ENV_TEXT.items():
            ttk.Radiobutton(self, text=name, value=env.value, variable=wiz.env,
                            command=self._changed).pack(anchor="w", pady=(wiz.px(6), 0))
            ttk.Label(self, text=desc, wraplength=wiz.px(530), foreground="#555").pack(
                anchor="w", padx=(wiz.px(24), 0))
        self.create = ttk.Checkbutton(
            self, variable=wiz.create_test,
            text="I have no TEST token - create a test company for me (TEST only)")
        self.create.pack(anchor="w", pady=(wiz.px(14), 0))
        self.warning = ttk.Label(self, text="", foreground="#c00000", wraplength=wiz.px(560))
        self.warning.pack(anchor="w", pady=(wiz.px(10), 0))
        self._changed()

    def _changed(self) -> None:
        env = Env(self.wiz.env.get())
        self.create.state(["!disabled"] if env == Env.TEST else ["disabled"])
        if env != Env.TEST:
            self.wiz.create_test.set(False)
        self.warning.config(text="PROD: every collective identifier you generate is a real, "
                                 "permanent KSeF record. The workbook will ask for confirmation "
                                 "before each PROD run." if env == Env.PROD else "")


class Company(Page):
    title = "Company"
    subtitle = "The company (KSeF context) the token belongs to"

    def __init__(self, wiz: Wizard) -> None:
        super().__init__(wiz)
        self.text("Company NIP (10 digits, without the PL prefix):")
        self.entry = ttk.Entry(self, textvariable=wiz.nip, width=20, font=("TkDefaultFont", 12))
        self.entry.pack(anchor="w")
        self.hint = ttk.Label(self, text="", foreground="#c00000")
        self.hint.pack(anchor="w", pady=(wiz.px(6), 0))
        wiz.nip.trace_add("write", lambda *_: self._check())

    def skip(self) -> bool:
        return self.wiz.create_test.get()

    def on_show(self) -> None:
        self.entry.focus_set()
        self._check()

    def _check(self) -> None:
        nip = self.wiz.nip.get().strip()
        err = userconfig.validate_nip(nip, Env(self.wiz.env.get())) if nip else None
        self.hint.config(text=err or "")

    def validate(self) -> str | None:
        nip = self.wiz.nip.get().strip().replace("-", "").replace(" ", "")
        self.wiz.nip.set(nip)
        return userconfig.validate_nip(nip, Env(self.wiz.env.get()))


class Token(Page):
    title = "KSeF token"
    subtitle = "Paste the token generated in KSeF for this company"

    def __init__(self, wiz: Wizard) -> None:
        super().__init__(wiz)
        self.keep = ttk.Checkbutton(self, variable=wiz.keep_token,
                                    text="Keep the token already stored on this computer",
                                    command=self._toggle)
        self.text("The token needs only the CollectiveIdentifierManage permission. It is "
                  "saved in the operating system's credential store (Windows Credential "
                  "Manager), not in the workbook or any file.")
        self.entry = ttk.Entry(self, textvariable=wiz.token, show="•", width=64)
        self.entry.pack(anchor="w", pady=(wiz.px(4), 0))
        self.show = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="Show token", variable=self.show,
                        command=lambda: self.entry.config(show="" if self.show.get()
                                                          else "•")).pack(anchor="w")

    def skip(self) -> bool:
        return self.wiz.create_test.get()

    def on_show(self) -> None:
        stored = get_token(self.wiz.selected_env(), self.wiz.nip.get())
        if stored:
            self.keep.pack(anchor="w", before=self.entry, pady=(0, self.wiz.px(8)))
        else:
            self.keep.pack_forget()
            self.wiz.keep_token.set(False)
        self._toggle()
        self.entry.focus_set()

    def _toggle(self) -> None:
        self.entry.state(["disabled"] if self.wiz.keep_token.get() else ["!disabled"])

    def validate(self) -> str | None:
        if self.wiz.keep_token.get():
            return None
        token = self.wiz.token.get().strip()
        self.wiz.token.set(token)
        if not userconfig.looks_like_token(token):
            return "That does not look like a KSeF token (20-200 characters, no spaces)."
        return None


class Connect(Page):
    title = "Connection test"
    subtitle = "Logging in to KSeF with these settings"

    def __init__(self, wiz: Wizard) -> None:
        super().__init__(wiz)
        self.status = self.text("")
        self.bar = ttk.Progressbar(self, mode="indeterminate", length=wiz.px(420))
        self.bar.pack(anchor="w", pady=(wiz.px(4), wiz.px(12)))
        self.result = ttk.Label(self, text="", wraplength=wiz.px(560), justify="left")
        self.result.pack(anchor="w")
        self.retry = ttk.Button(self, text="Try again", command=self.start)
        self.ok = False
        self.q: queue.Queue[tuple[bool, str]] = queue.Queue()

    def on_show(self) -> None:
        self.start()

    def start(self) -> None:
        self.ok = False
        self.retry.pack_forget()
        self.wiz.set_next_enabled(False)
        self.result.config(text="", foreground="")
        self.status.config(text="Creating a test company on KSeF TEST..."
                           if self.wiz.create_test.get() else
                           f"Connecting to KSeF {self.wiz.env.get().upper()}...")
        self.bar.start(12)
        threading.Thread(target=self._work, daemon=True).start()
        self.after(150, self._poll)

    def _work(self) -> None:
        w = self.wiz
        try:
            if w.create_test.get():
                with KsefClient(Env.TEST) as client:
                    ctx = testenv.bootstrap(client, testenv.random_test_nip(),
                                            echo=lambda s: None)
                userconfig.write_excel_ini(Env.TEST, ctx.nip)
                w.nip_result = ctx.nip
                self.q.put((True, f"Test company created on KSeF TEST (NIP {ctx.nip}) with a "
                                  "CollectiveIdentifierManage token."))
                return
            env, nip = w.selected_env(), w.nip.get()
            token = get_token(env, nip) if w.keep_token.get() else w.token.get()
            res = userconfig.test_connection(env, nip, token or "")
            if res.ok:
                userconfig.save_setup(env, nip, None if w.keep_token.get() else token)
                w.nip_result = nip
            self.q.put((res.ok, res.message))
        except Exception as e:  # noqa: BLE001 - shown to the user
            self.q.put((False, userconfig.explain(e)))

    def _poll(self) -> None:
        try:
            ok, msg = self.q.get_nowait()
        except queue.Empty:
            self.after(150, self._poll)
            return
        self.bar.stop()
        self.ok = ok
        self.status.config(text="Done." if ok else "The connection test failed.")
        self.result.config(text=("✔  " if ok else "✖  ") + msg,
                           foreground="#2e7d32" if ok else "#c00000")
        if ok:
            self.wiz.token.set("")  # do not keep the secret in the UI
        else:
            self.retry.pack(anchor="w", pady=(self.wiz.px(12), 0))
        self.wiz.set_next_enabled(ok)

    def validate(self) -> str | None:
        return None if self.ok else "The connection test has not succeeded yet."


class Done(Page):
    title = "All set"
    subtitle = "The workbook is ready to use"

    def __init__(self, wiz: Wizard) -> None:
        super().__init__(wiz)
        self.summary = ttk.Label(self, text="", justify="left", wraplength=wiz.px(560))
        self.summary.pack(anchor="w")
        self.open_wb = tk.BooleanVar(value=True)
        self.open_box = ttk.Checkbutton(self, text="Open the workbook now",
                                        variable=self.open_wb)
        self.open_box.pack(anchor="w", pady=(wiz.px(14), 0))

    def on_show(self) -> None:
        wb = userconfig.workbook_path()
        trust = userconfig.trusted_location_status()
        lines = [f"Environment:   {self.wiz.env.get().upper()}",
                 f"Company NIP:   {self.wiz.nip_result}",
                 "KSeF token:    stored in the credential store",
                 f"Workbook:      {wb}" + ("" if wb.exists() else "  (not found)")]
        if trust == "blocked by policy":
            lines.append("\nMacros: company policy does not allow user Trusted Locations. "
                         "Ask IT to trust the folder above, otherwise Excel will block the "
                         "workbook's macros.")
        elif trust == "missing":
            lines.append("\nMacros: the workbook folder is not registered as an Excel "
                         "Trusted Location - re-run the installer, or enable macros when "
                         "Excel asks.")
        self.summary.config(text="\n".join(lines))
        if wb.exists():
            self.open_box.state(["!disabled"])
        else:
            self.open_wb.set(False)
            self.open_box.state(["disabled"])


class Wizard(tk.Tk):
    PAGES: tuple[type[Page], ...] = (Welcome, Environment, Company, Token, Connect, Done)

    def __init__(self, scale: float | None = None) -> None:
        super().__init__()
        # Hidden until laid out: _fit_window measures pages with placeholder texts, which
        # Windows would otherwise paint (seen in 0.2.0 as flashing "XXXX" screens)
        self.withdraw()
        if scale:  # tests: simulate a high-DPI screen (Tk points -> pixels)
            self.tk.call("tk", "scaling", scale)
        # On high-DPI Windows screens Tk renders text larger; size everything in scaled units
        self.factor = max(1.0, float(self.tk.call("tk", "scaling")) / (96 / 72))
        self.title(f"{userconfig.APP_NAME} setup")
        ini = userconfig.read_excel_ini()
        env = (ini.get("environment") or "TEST").lower()
        self.env = tk.StringVar(value=env if env in {e.value for e in Env} else "test")
        nip = ini.get("context_nip", "")
        if not nip:
            ctx = load_context(Env(self.env.get()))
            nip = ctx.nip if ctx else ""
        self.nip = tk.StringVar(value=nip)
        self.token = tk.StringVar(value="")
        self.keep_token = tk.BooleanVar(value=True)
        self.create_test = tk.BooleanVar(value=False)
        self.nip_result = nip

        head = tk.Frame(self, bg="#1f4e78", padx=self.px(24), pady=self.px(12))
        head.pack(fill="x")
        self.h1 = tk.Label(head, text="", bg="#1f4e78", fg="white",
                           font=("TkDefaultFont", 14, "bold"))
        self.h1.pack(anchor="w")
        self.h2 = tk.Label(head, text="", bg="#1f4e78", fg="#dbe7f3")
        self.h2.pack(anchor="w")
        # The button bar is packed before the page body so it always keeps its space
        foot = ttk.Frame(self, padding=(self.px(16), self.px(10)))
        foot.pack(fill="x", side="bottom")
        ttk.Separator(self).pack(fill="x", side="bottom")
        self.body = ttk.Frame(self)
        ttk.Label(foot, text=f"v{__version__}", foreground="#888").pack(side="left")
        self.cancel_btn = ttk.Button(foot, text="Cancel", command=self.destroy)
        self.cancel_btn.pack(side="right")
        self.next_btn = ttk.Button(foot, text="Next >", command=self.next)
        self.next_btn.pack(side="right", padx=(0, self.px(8)))
        self.back_btn = ttk.Button(foot, text="< Back", command=self.back)
        self.back_btn.pack(side="right", padx=(0, self.px(4)))
        self.body.pack(fill="both", expand=True)

        self.pages = [cls(self) for cls in self.PAGES]
        self._fit_window()
        self.i = 0
        self.show(0)
        self._center()
        self.deiconify()

    def _center(self) -> None:
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        if w <= 1 or h <= 1:  # not realised yet while withdrawn: use the planned size
            w, h = self.winfo_reqwidth(), self.winfo_reqheight()
            geo = self.geometry().split("+")[0]
            if "x" in geo:
                w, h = (int(v) for v in geo.split("x"))
        x = max(0, (self.winfo_screenwidth() - w) // 2)
        y = max(0, (self.winfo_screenheight() - h) // 3)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def px(self, n: float) -> int:
        """Design pixels (at 100 % display scaling) -> pixels on this screen."""
        return round(n * self.factor)

    def _fit_window(self) -> None:
        """Size the window for the tallest page, with its longest dynamic texts, so every
        page fits at any display scaling. Capped to the screen; stays resizable."""
        env_page = next(p for p in self.pages if isinstance(p, Environment))
        done_page = next(p for p in self.pages if isinstance(p, Done))
        conn_page = next(p for p in self.pages if isinstance(p, Connect))
        env_page.warning.config(text="X " * 90)
        done_page.summary.config(text="\n".join(["Environment: X"] * 4) + "\n\n" + "X " * 120)
        conn_page.result.config(text="X " * 100)
        conn_page.retry.pack(anchor="w")
        need_w = need_h = 0
        for p in self.pages:
            p.pack(fill="both", expand=True)
            self.update_idletasks()
            need_w = max(need_w, self.winfo_reqwidth())
            need_h = max(need_h, self.winfo_reqheight())
            p.pack_forget()
        env_page.warning.config(text="")
        done_page.summary.config(text="")
        conn_page.result.config(text="")
        conn_page.retry.pack_forget()
        w = min(max(need_w, self.px(640)), self.winfo_screenwidth() - 40)
        h = min(need_h + self.px(16), self.winfo_screenheight() - 80)
        self.geometry(f"{w}x{h}")
        self.minsize(min(w, self.px(560)), min(h, self.px(420)))

    def selected_env(self) -> Env:
        return Env(self.env.get())

    def set_next_enabled(self, on: bool) -> None:
        self.next_btn.state(["!disabled"] if on else ["disabled"])

    def show(self, i: int) -> None:
        for p in self.pages:
            p.pack_forget()
        self.i = i
        page = self.pages[i]
        self.h1.config(text=page.title)
        self.h2.config(text=page.subtitle)
        page.pack(fill="both", expand=True)
        last = i == len(self.pages) - 1
        self.next_btn.config(text="Finish" if last else "Next >")
        self.set_next_enabled(True)
        self.back_btn.state(["disabled"] if i == 0 or last else ["!disabled"])
        self.cancel_btn.state(["disabled"] if last else ["!disabled"])
        page.on_show()

    def _step(self, start: int, delta: int) -> int:
        j = start + delta
        while 0 <= j < len(self.pages) and self.pages[j].skip():
            j += delta
        return j

    def next(self) -> None:
        page = self.pages[self.i]
        err = page.validate()
        if err:
            messagebox.showwarning(self.title(), err, parent=self)
            return
        if self.i == len(self.pages) - 1:
            self.finish()
            return
        self.show(self._step(self.i, 1))

    def back(self) -> None:
        j = self._step(self.i, -1)
        if j >= 0:
            self.show(j)

    def finish(self) -> None:
        done = self.pages[-1]
        assert isinstance(done, Done)
        if done.open_wb.get():
            open_file(str(userconfig.workbook_path()))
        self.destroy()


def open_file(path: str, opener: Callable[[str], None] | None = None) -> None:
    if opener:
        opener(path)
    elif (startfile := getattr(os, "startfile", None)) is not None:  # Windows
        startfile(path)
    elif sys.platform == "darwin":
        subprocess.run(["open", path], check=False)  # noqa: S603, S607


def main() -> None:
    if sys.platform == "win32":
        import ctypes

        windll = getattr(ctypes, "windll", None)
        try:  # crisp text on high-DPI screens
            if windll is not None:
                windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            pass
    Wizard().mainloop()


if __name__ == "__main__":
    main()
