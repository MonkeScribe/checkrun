"""Basic Tk GUI: drop a raw file, pick an output folder, fill the template, make the IIF."""
from __future__ import annotations

import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from checkrun import core, settings

try:  # optional: real drag-and-drop. Without it the box is click-to-browse.
    from tkinterdnd2 import DND_FILES, TkinterDnD

    _Root = TkinterDnD.Tk
except ImportError:  # pragma: no cover
    DND_FILES, _Root = None, tk.Tk

FILETYPES = [("Spreadsheets", "*.xlsx *.xlsm *.xls *.csv"), ("All files", "*.*")]


class App(_Root):
    def __init__(self) -> None:
        super().__init__()
        self.title("Check Run → QuickBooks IIF")
        self.geometry("560x420")
        self.minsize(520, 380)

        cfg = settings.load()
        self.raw_path: Path | None = None
        self.filled_path: Path | None = None
        self.out_dir = tk.StringVar(value=cfg.get("output_dir", str(Path.home() / "Documents")))
        self.status = tk.StringVar(value="Drop a raw data file to begin.")

        pad = {"padx": 14, "pady": 6}

        self.drop = tk.Label(
            self, text="Drop raw data file here\n(or click to browse)", relief="groove", bd=2,
            height=5, cursor="hand2", wraplength=480,
        )
        self.drop.pack(fill="x", **pad)
        self.drop.bind("<Button-1>", lambda _e: self.browse_raw())
        if DND_FILES:
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", self.on_drop)

        folder = ttk.Frame(self)
        folder.pack(fill="x", **pad)
        ttk.Label(folder, text="Output folder:").pack(side="left")
        ttk.Entry(folder, textvariable=self.out_dir).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(folder, text="Browse…", command=self.browse_out).pack(side="left")

        self.fill_btn = ttk.Button(self, text="1. Fill template", command=self.do_fill, state="disabled")
        self.fill_btn.pack(fill="x", **pad)
        self.iif_btn = ttk.Button(self, text="2. Convert to IIF", command=self.do_iif, state="disabled")
        self.iif_btn.pack(fill="x", padx=14, pady=(0, 6))
        ttk.Button(self, text="Use a different filled template…", command=self.browse_filled).pack(anchor="w", padx=14)

        ttk.Label(self, textvariable=self.status, wraplength=520, justify="left").pack(fill="both", expand=True, **pad)

    # -- inputs -------------------------------------------------------------
    def on_drop(self, event) -> None:
        paths = self.tk.splitlist(event.data)
        if paths:
            self.set_raw(Path(paths[0]))

    def browse_raw(self) -> None:
        p = filedialog.askopenfilename(title="Select raw data file", filetypes=FILETYPES)
        if p:
            self.set_raw(Path(p))

    def set_raw(self, path: Path) -> None:
        self.raw_path, self.filled_path = path, None
        self.drop.config(text=f"✔ {path.name}\n(drop another file to replace)")
        self.fill_btn.config(state="normal")
        self.iif_btn.config(state="disabled")
        self.status.set("Ready. Click “1. Fill template”.")

    def browse_filled(self) -> None:
        p = filedialog.askopenfilename(title="Select filled template", filetypes=[("Excel", "*.xlsx")])
        if p:
            self.filled_path = Path(p)
            self.iif_btn.config(state="normal")
            self.status.set(f"Using filled template: {self.filled_path.name}\nClick “2. Convert to IIF”.")

    def browse_out(self) -> None:
        p = filedialog.askdirectory(title="Choose output folder", initialdir=self.out_dir.get() or None)
        if p:
            self.out_dir.set(p)
            settings.save({"output_dir": p})

    # -- actions ------------------------------------------------------------
    def _out_dir(self) -> Path | None:
        d = Path(self.out_dir.get()).expanduser()
        try:
            d.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            messagebox.showerror("Output folder", f"Can't use that folder:\n{e}")
            return None
        settings.save({"output_dir": str(d)})
        return d

    def do_fill(self) -> None:
        out = self._out_dir()
        if not out or not self.raw_path:
            return
        dest = out / f"Checks_{datetime.now():%Y-%m-%d_%H%M}_template.xlsx"
        try:
            n = core.fill_template(self.raw_path, dest)
        except core.CheckRunError as e:
            messagebox.showerror("Can't fill template", str(e))
            self.status.set("Fix the problems above and try again.")
            return
        except Exception as e:  # unreadable file, permissions, etc.
            messagebox.showerror("Error", f"{type(e).__name__}: {e}")
            return
        self.filled_path = dest
        self.iif_btn.config(state="normal")
        note = f"\nNote: the template's own IIF Output tab only covers {core.TEMPLATE_CAPACITY} checks." if n > core.TEMPLATE_CAPACITY else ""
        self.status.set(f"Filled {n} check(s) → {dest}\nReview it if you like, then click “2. Convert to IIF”.{note}")

    def do_iif(self) -> None:
        out = self._out_dir()
        if not out or not self.filled_path:
            return
        dest = out / (self.filled_path.stem.removesuffix("_template") + ".iif")
        try:
            n = core.convert_template_to_iif(self.filled_path, dest)
        except core.CheckRunError as e:
            messagebox.showerror("Can't convert", str(e))
            return
        except Exception as e:
            messagebox.showerror("Error", f"{type(e).__name__}: {e}")
            return
        self.status.set(f"Wrote {n} check(s) → {dest}\nBack up QuickBooks before importing.")


def main() -> None:
    App().mainloop()
