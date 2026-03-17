"""
ExtraDataCleaner — gui.py
==========================
Windows 7 / XP classic themed GUI.
Launched automatically when the app is started without CLI arguments.
"""

from __future__ import annotations

import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

_HERE = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
sys.path.insert(0, str(_HERE))

from core import DataCleaner      # noqa: E402
import pandas as pd               # noqa: E402


# ── Windows 7 / XP Classic colour palette ─────────────────────────────────────
BG       = "#F0F0F0"   # System button-face grey
PANEL    = "#FFFFFF"   # Entry / panel background
HDR_TOP  = "#1B5EA6"   # Title bar deep blue (XP Royale / Luna)
HDR_BOT  = "#2878CE"   # Title bar lighter band
BORDER   = "#ABABAB"   # Standard Windows border
BTN      = "#E1E1E1"   # Default button face
BTN_H    = "#E8F4FF"   # Button hover tint
BTN_P    = "#2B6CB0"   # Primary/accent button (blue)
BTN_PH   = "#1A4F8A"   # Primary button hover
FG       = "#000000"   # Standard text
FG_DIM   = "#6D6D6D"   # Greyed label text
FG_HEAD  = "#FFFFFF"   # Text on blue header
SUCCESS  = "#1A7524"   # Dark green
WARNING  = "#8B5E00"   # Dark amber
ERROR    = "#C0392B"   # Dark red
LOG_BG   = "#FFFFFF"
LOG_FG   = "#000000"
FONT     = ("Tahoma", 9)
FONT_B   = ("Tahoma", 9, "bold")
FONT_H   = ("Tahoma", 11, "bold")
MONO     = ("Courier New", 8)


# ── small widget helpers ──────────────────────────────────────────────────────

class _Btn(tk.Button):
    """Classic raised Windows button with hover highlight."""
    def __init__(self, parent, text, cmd, primary=False, small=False, **kw):
        self._bg_up   = BTN_P if primary else BTN
        self._bg_down = BTN_PH if primary else BTN_H
        self._fg      = FG_HEAD if primary else FG
        super().__init__(
            parent, text=text, command=cmd,
            bg=self._bg_up, fg=self._fg,
            activebackground=self._bg_down, activeforeground=self._fg,
            relief="raised", bd=1, cursor="hand2",
            font=(FONT[0], 8 if small else 9),
            padx=4 if small else 10, pady=2 if small else 4,
            **kw
        )
        self.bind("<Enter>", lambda e: self.config(bg=self._bg_down))
        self.bind("<Leave>", lambda e: self.config(bg=self._bg_up))


def _sep(parent):
    return tk.Frame(parent, bg=BORDER, height=1)


def _cb(parent, text, default=True):
    """Checkbox returning its BooleanVar."""
    var = tk.BooleanVar(value=default)
    tk.Checkbutton(
        parent, text=text, variable=var,
        bg=BG, fg=FG, activebackground=BG, activeforeground=FG,
        selectcolor=BG, font=FONT
    ).pack(anchor="w", pady=1)
    return var


def _entry(parent, var, width=12):
    return tk.Entry(
        parent, textvariable=var, bg=PANEL, fg=FG,
        font=FONT, relief="sunken", bd=1, width=width,
        insertbackground=FG
    )


# ── main window ───────────────────────────────────────────────────────────────

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("ExtraDataCleaner")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(680, 570)
        self._files: list[Path] = []
        self._build_ui()
        self._center()
        self._load_icon()

    # ── icon ─────────────────────────────────────────────────────────────────

    def _load_icon(self):
        ico = _HERE / "icon.ico"
        try:
            if ico.exists():
                self.iconbitmap(str(ico))
        except Exception:
            pass

    # ── layout ───────────────────────────────────────────────────────────────

    def _build_ui(self):

        # ── blue title bar (XP Luna style) ───────────────────────────────────
        hbar = tk.Frame(self, bg=HDR_TOP, height=44)
        hbar.pack(fill="x")
        hbar.pack_propagate(False)
        tk.Frame(hbar, bg=HDR_BOT, height=2).pack(fill="x", side="bottom")
        inner_h = tk.Frame(hbar, bg=HDR_TOP)
        inner_h.pack(fill="both", expand=True, padx=12)
        tk.Label(inner_h, text="ExtraDataCleaner",
                 bg=HDR_TOP, fg=FG_HEAD, font=FONT_H).pack(side="left", pady=10)
        tk.Label(inner_h, text="   CSV & Excel data formatter",
                 bg=HDR_TOP, fg="#9FC8EF", font=FONT).pack(side="left", pady=10)

        # ── file selector ─────────────────────────────────────────────────────
        _sep(self).pack(fill="x")
        ff = tk.LabelFrame(self, text=" Input Files ",
                           bg=BG, fg=FG, font=FONT_B, bd=1, relief="groove",
                           padx=8, pady=6)
        ff.pack(fill="x", padx=10, pady=(8, 4))

        self._file_lbl = tk.Label(ff, text="No files selected",
                                  bg=BG, fg=FG_DIM, font=FONT,
                                  anchor="w", justify="left", wraplength=520)
        self._file_lbl.pack(side="left", fill="x", expand=True)
        br = tk.Frame(ff, bg=BG)
        br.pack(side="right")
        _Btn(br, "Browse...", self._browse_files, primary=True).pack(side="left", padx=(4, 0))
        _Btn(br, "Clear",     self._clear_files).pack(side="left", padx=(4, 0))

        # ── two-column options ────────────────────────────────────────────────
        cols = tk.Frame(self, bg=BG)
        cols.pack(fill="x", padx=10, pady=(0, 4))
        cols.columnconfigure(0, weight=1)
        cols.columnconfigure(1, weight=1)

        # left — cleaning options
        lf = tk.LabelFrame(cols, text=" Cleaning Options ",
                           bg=BG, fg=FG, font=FONT_B, bd=1, relief="groove",
                           padx=8, pady=4)
        lf.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        self._v_snake   = _cb(lf, "Convert column names to snake_case")
        self._v_nulls   = _cb(lf, "Unify null values  (N/A, none, #N/A…)")
        self._v_numeric = _cb(lf, "Parse numeric columns  (currency, %, etc.)")
        self._v_bool    = _cb(lf, "Normalise booleans  (yes/no → True/False)")
        self._v_erows   = _cb(lf, "Drop fully-empty rows")
        self._v_ecols   = _cb(lf, "Drop fully-empty columns")
        self._v_hidden  = _cb(lf, "Skip hidden Excel rows / columns")

        # right — output
        rf = tk.LabelFrame(cols, text=" Output Settings ",
                           bg=BG, fg=FG, font=FONT_B, bd=1, relief="groove",
                           padx=8, pady=4)
        rf.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        rf.columnconfigure(1, weight=1)

        def row(r, lbl, wfn, extra=None):
            tk.Label(rf, text=lbl, bg=BG, fg=FG, font=FONT,
                     anchor="w", width=16).grid(row=r, column=0, sticky="w", pady=3)
            wfn(rf).grid(row=r, column=1, sticky="ew", padx=(4, 0), pady=3)
            if extra:
                extra(rf).grid(row=r, column=2, padx=(3, 0), pady=3)

        self._outdir_var = tk.StringVar(value="Same as input")
        row(0, "Output folder:", lambda p: _entry(p, self._outdir_var, 16),
            extra=lambda p: _Btn(p, "...", self._browse_outdir, small=True))

        self._suffix_var = tk.StringVar(value="_clean")
        row(1, "Filename suffix:", lambda p: _entry(p, self._suffix_var, 10))

        self._fmt_var = tk.StringVar(value="csv")
        def fmt_widget(p):
            f = tk.Frame(p, bg=BG)
            for v, l in [("csv", "CSV (.csv)"), ("xlsx", "Excel (.xlsx)")]:
                tk.Radiobutton(f, text=l, variable=self._fmt_var, value=v,
                               bg=BG, fg=FG, activebackground=BG,
                               activeforeground=FG, selectcolor=BG,
                               font=FONT).pack(side="left", padx=(0, 8))
            return f
        row(2, "Output format:", fmt_widget)

        self._sheet_var = tk.StringVar()
        row(3, "Excel sheet:", lambda p: _entry(p, self._sheet_var, 14))

        self._hdr_var = tk.StringVar()
        row(4, "Header row:", lambda p: _entry(p, self._hdr_var, 4))
        tk.Label(rf, text="(blank = auto)", bg=BG, fg=FG_DIM,
                 font=(FONT[0], 8)).grid(row=4, column=2, padx=(3, 0))

        # ── action bar ────────────────────────────────────────────────────────
        _sep(self).pack(fill="x", pady=(4, 0))
        act = tk.Frame(self, bg=BG, pady=6)
        act.pack(fill="x", padx=10)
        self._run_btn = _Btn(act, "  Run Cleaner  ", self._run, primary=True)
        self._run_btn.pack(side="left")
        self._dry_var = tk.BooleanVar(value=False)
        tk.Checkbutton(act, text="Dry run  (report only — no files written)",
                       variable=self._dry_var,
                       bg=BG, fg=FG_DIM, selectcolor=BG,
                       activebackground=BG, activeforeground=FG,
                       font=FONT).pack(side="left", padx=14)

        # ── log area ──────────────────────────────────────────────────────────
        _sep(self).pack(fill="x")
        log_outer = tk.Frame(self, bg=BG)
        log_outer.pack(fill="both", expand=True, padx=10, pady=(4, 0))

        lh = tk.Frame(log_outer, bg=BG)
        lh.pack(fill="x")
        tk.Label(lh, text="Cleaning Report",
                 bg=BG, fg=FG, font=FONT_B).pack(side="left")
        _Btn(lh, "Clear", self._clear_log, small=True).pack(side="right")

        log_wrap = tk.Frame(log_outer, bg=BORDER, bd=1, relief="sunken")
        log_wrap.pack(fill="both", expand=True, pady=(4, 0))
        self._log = tk.Text(log_wrap, bg=LOG_BG, fg=LOG_FG, font=MONO,
                            relief="flat", bd=3, wrap="word",
                            insertbackground=FG,
                            selectbackground=BTN_P, selectforeground=FG_HEAD)
        sb = ttk.Scrollbar(log_wrap, command=self._log.yview)
        self._log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._log.pack(fill="both", expand=True)

        self._log.tag_config("ok",   foreground=SUCCESS, font=(MONO[0], MONO[1], "bold"))
        self._log.tag_config("warn", foreground=WARNING)
        self._log.tag_config("err",  foreground=ERROR,   font=(MONO[0], MONO[1], "bold"))
        self._log.tag_config("head", foreground=HDR_TOP, font=(MONO[0], MONO[1], "bold"))
        self._log.tag_config("dim",  foreground=FG_DIM)

        # ── status bar (classic Windows chrome at bottom) ─────────────────────
        _sep(self).pack(fill="x")
        self._status = tk.Label(self, text="  Ready",
                                bg=BG, fg=FG_DIM, font=(FONT[0], 8),
                                anchor="w", relief="sunken", bd=1)
        self._status.pack(fill="x", side="bottom")

        self._log_line("Ready.  Select CSV or Excel files and click Run Cleaner.", "dim")

    # ── actions ───────────────────────────────────────────────────────────────

    def _browse_files(self):
        paths = filedialog.askopenfilenames(
            title="Select CSV or Excel files",
            filetypes=[
                ("Spreadsheets", "*.csv *.tsv *.txt *.xlsx *.xls *.xlsm *.xlsb *.ods"),
                ("CSV files",    "*.csv *.tsv *.txt"),
                ("Excel files",  "*.xlsx *.xls *.xlsm *.xlsb *.ods"),
                ("All files",    "*.*"),
            ]
        )
        if paths:
            self._files = [Path(p) for p in paths]
            self._refresh_label()

    def _browse_outdir(self):
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            self._outdir_var.set(d)

    def _clear_files(self):
        self._files = []
        self._refresh_label()

    def _refresh_label(self):
        if not self._files:
            self._file_lbl.config(text="No files selected", fg=FG_DIM)
        elif len(self._files) == 1:
            self._file_lbl.config(text=str(self._files[0]), fg=FG)
        else:
            names = ", ".join(f.name for f in self._files[:3])
            more  = f"  (+{len(self._files)-3} more)" if len(self._files) > 3 else ""
            self._file_lbl.config(text=f"{len(self._files)} files: {names}{more}", fg=FG)

    def _clear_log(self):
        self._log.delete("1.0", "end")

    def _run(self):
        if not self._files:
            messagebox.showwarning("No files selected",
                                   "Please select at least one CSV or Excel file.")
            return
        self._run_btn.config(state="disabled", text="  Running...  ")
        self._set_status("Processing files…")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            hdr = self._hdr_var.get().strip()
            cleaner = DataCleaner(
                snake_case      = self._v_snake.get(),
                unify_nulls     = self._v_nulls.get(),
                numeric_coerce  = self._v_numeric.get(),
                bool_coerce     = self._v_bool.get(),
                drop_empty_rows = self._v_erows.get(),
                drop_empty_cols = self._v_ecols.get(),
                skip_hidden     = self._v_hidden.get(),
                sheet           = self._sheet_var.get().strip() or None,
                header_row      = int(hdr) if hdr.isdigit() else None,
            )
            suffix = self._suffix_var.get() or "_clean"
            fmt    = self._fmt_var.get()
            dry    = self._dry_var.get()
            odir   = self._outdir_var.get().strip()

            for src in self._files:
                self._log_line(f"\n{'─'*48}", "dim")
                self._log_line(f"Input:  {src.name}", "head")
                try:
                    frames = cleaner.clean_file(src)
                except Exception as exc:
                    self._log_line(f"ERROR: {exc}", "err")
                    continue

                for ln in cleaner.get_report().splitlines():
                    tag = "ok" if "→" in ln else ("warn" if "warn" in ln.lower() else "")
                    self._log_line(ln, tag)

                if dry:
                    self._log_line("[dry-run] No files written.", "warn")
                    continue

                out = Path(odir) if odir and odir != "Same as input" else src.parent
                out.mkdir(parents=True, exist_ok=True)
                for name, df in frames.items():
                    safe = name.replace("/", "-").replace("\\", "-")
                    stem = src.stem + suffix if (len(frames) == 1 and name == "Sheet1") \
                           else f"{src.stem}_{safe}{suffix}"
                    if fmt == "xlsx":
                        p = out / (stem + ".xlsx")
                        with pd.ExcelWriter(p, engine="openpyxl") as w:
                            df.to_excel(w, sheet_name=name[:31], index=False)
                    else:
                        p = out / (stem + ".csv")
                        df.to_csv(p, index=False, encoding="utf-8-sig")
                    self._log_line(f"  \u2713  {p.name}   \u2192   {p.parent}", "ok")

            self._log_line(f"\n{'─'*48}", "dim")
            self._log_line(
                "Dry run complete — no files written." if dry else "Done.",
                "warn" if dry else "ok"
            )
            self._set_status("Ready")

        except Exception as exc:
            self._log_line(f"Unexpected error: {exc}", "err")
            self._set_status("Error — see report")
        finally:
            self.after(0, lambda: self._run_btn.config(
                state="normal", text="  Run Cleaner  "))

    # ── thread-safe log / status ──────────────────────────────────────────────

    def _log_line(self, msg: str, tag: str = ""):
        def _do():
            if tag:
                self._log.insert("end", msg + "\n", tag)
            else:
                self._log.insert("end", msg + "\n")
            self._log.see("end")
        self.after(0, _do)

    def _set_status(self, msg: str):
        self.after(0, lambda: self._status.config(text=f"  {msg}"))

    def _center(self):
        self.update_idletasks()
        w, h = 730, 610
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")


# ── entry point ───────────────────────────────────────────────────────────────

def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
