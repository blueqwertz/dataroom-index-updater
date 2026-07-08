#!/usr/bin/env python3
"""
Dataroom Index Updater -- GUI
=============================

Kleines Fenster um update_index.py: zwei Dateifelder (Current-Index und
frischer Datenraum-Export), Browse-Buttons und Drag&Drop aus dem Explorer.
Optionen (Anker-Spalte, Markierungs-Spalten, Formatierung neuer Zeilen)
sind eingeklappt und werden in %APPDATA%\\IndexUpdater\\settings.json
gespeichert.

Start:  python app.py
"""

import json
import os
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

from update_index import run_update, shared_headers

APP_NAME = "Dataroom Index Updater"
AUTO = "Auto-detect"
SETTINGS_DIR = Path(os.environ.get("APPDATA", Path.home())) / "IndexUpdater"
SETTINGS_FILE = SETTINGS_DIR / "settings.json"

DEFAULT_SETTINGS = {
    "highlight_cols": "",
    "fill_enabled": True,
    "fill_color": "#FFFF00",
    "bold": False,
    "underline": "none",       # none | single | double
    "font_color_enabled": False,
    "font_color": "#FF0000",
    "options_expanded": False,
}


def load_settings():
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            data = json.load(f)
        return {**DEFAULT_SETTINGS, **data}
    except (OSError, ValueError):
        return dict(DEFAULT_SETTINGS)


def save_settings(data):
    try:
        SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except OSError:
        pass


def parse_drop(data):
    """tkinterdnd2 liefert z.B. '{C:/pfad mit leerzeichen/a.xlsx} C:/b.xlsx'"""
    paths, token, in_brace = [], "", False
    for ch in data:
        if ch == "{":
            in_brace = True
        elif ch == "}":
            in_brace = False
            paths.append(token)
            token = ""
        elif ch == " " and not in_brace:
            if token:
                paths.append(token)
            token = ""
        else:
            token += ch
    if token:
        paths.append(token)
    return [p for p in paths if p]


class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_NAME)
        root.resizable(False, False)
        self.settings = load_settings()

        pad = {"padx": 12, "pady": 4}
        frm = ttk.Frame(root, padding=(12, 10, 12, 6))
        frm.grid(sticky="nsew")

        # --- Dateifelder -------------------------------------------------
        self.current_var = tk.StringVar()
        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()

        self._file_row(frm, 0, "Current index (your enriched version):",
                       self.current_var)
        self._file_row(frm, 2, "New dataroom export:", self.input_var)

        for var in (self.current_var, self.input_var):
            var.trace_add("write", self._on_files_changed)

        # --- Optionen (einklappbar) --------------------------------------
        self.opt_visible = tk.BooleanVar(value=self.settings["options_expanded"])
        self.opt_btn = ttk.Button(frm, style="Toolbutton",
                                  command=self.toggle_options)
        self.opt_btn.grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))

        self.opt_frame = ttk.Frame(frm, padding=(16, 4, 0, 4))
        self._build_options(self.opt_frame)
        self._apply_options_visibility()

        # --- Output + Aktion ---------------------------------------------
        out_row = ttk.Frame(frm)
        out_row.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(10, 2))
        out_row.columnconfigure(1, weight=1)
        ttk.Label(out_row, text="Output:").grid(row=0, column=0, sticky="w")
        self.out_entry = ttk.Entry(out_row, textvariable=self.output_var, width=44)
        self.out_entry.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ttk.Button(out_row, text="...", width=3,
                   command=self.pick_output).grid(row=0, column=2)

        self.run_btn = ttk.Button(out_row, text="Update index",
                                  command=self.run, state="disabled")
        self.run_btn.grid(row=0, column=3, padx=(10, 0))

        # --- Statusleiste --------------------------------------------------
        self.status_var = tk.StringVar(value=self._ready_text())
        status = ttk.Label(root, textvariable=self.status_var, relief="sunken",
                           anchor="w", padding=(8, 3))
        status.grid(sticky="ew")
        root.columnconfigure(0, weight=1)

        root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ------------------------------------------------------------------ UI
    def _file_row(self, parent, row, label, var):
        ttk.Label(parent, text=label).grid(row=row, column=0, columnspan=3,
                                           sticky="w", pady=(6, 1))
        entry = ttk.Entry(parent, textvariable=var, width=58)
        entry.grid(row=row + 1, column=0, columnspan=2, sticky="ew")
        ttk.Button(parent, text="Browse...",
                   command=lambda: self.browse(var)).grid(row=row + 1, column=2,
                                                          padx=(6, 0))
        if HAS_DND:
            entry.drop_target_register(DND_FILES)
            entry.dnd_bind("<<Drop>>",
                           lambda e, v=var: self._on_drop(e, v))
        return entry

    def _build_options(self, f):
        f.columnconfigure(1, weight=1)

        ttk.Label(f, text="Anchor column(s):").grid(row=0, column=0, sticky="w",
                                                    pady=2)
        self.anchor_var = tk.StringVar(value=AUTO)
        self.anchor_box = ttk.Combobox(f, textvariable=self.anchor_var,
                                       values=[AUTO], state="readonly", width=34)
        self.anchor_box.grid(row=0, column=1, sticky="w", padx=(8, 0), pady=2)

        ttk.Label(f, text="Highlight columns:").grid(row=1, column=0, sticky="w",
                                                     pady=2)
        hl = ttk.Frame(f)
        hl.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=2)
        self.hl_var = tk.StringVar(value=self.settings["highlight_cols"])
        ttk.Entry(hl, textvariable=self.hl_var, width=12).pack(side="left")
        ttk.Label(hl, text='empty = all   (e.g. "B-D" or "B,D-F")',
                  foreground="#666").pack(side="left", padx=(8, 0))

        ttk.Label(f, text="Mark new rows:").grid(row=2, column=0, sticky="nw",
                                                 pady=2)
        mk = ttk.Frame(f)
        mk.grid(row=2, column=1, sticky="w", padx=(8, 0), pady=2)

        self.fill_on = tk.BooleanVar(value=self.settings["fill_enabled"])
        ttk.Checkbutton(mk, text="Fill", variable=self.fill_on).pack(side="left")
        self.fill_btn = tk.Button(mk, width=3, relief="ridge",
                                  bg=self.settings["fill_color"],
                                  command=lambda: self.pick_color("fill_color",
                                                                  self.fill_btn))
        self.fill_btn.pack(side="left", padx=(4, 14))

        self.bold_on = tk.BooleanVar(value=self.settings["bold"])
        ttk.Checkbutton(mk, text="Bold", variable=self.bold_on).pack(side="left",
                                                                     padx=(0, 14))

        ttk.Label(mk, text="Underline:").pack(side="left")
        self.underline_var = tk.StringVar(value=self.settings["underline"])
        ttk.Combobox(mk, textvariable=self.underline_var, width=7,
                     values=["none", "single", "double"],
                     state="readonly").pack(side="left", padx=(4, 14))

        self.fc_on = tk.BooleanVar(value=self.settings["font_color_enabled"])
        ttk.Checkbutton(mk, text="Font color", variable=self.fc_on).pack(side="left")
        self.fc_btn = tk.Button(mk, width=3, relief="ridge",
                                bg=self.settings["font_color"],
                                command=lambda: self.pick_color("font_color",
                                                                self.fc_btn))
        self.fc_btn.pack(side="left", padx=(4, 0))

    def toggle_options(self):
        self.opt_visible.set(not self.opt_visible.get())
        self._apply_options_visibility()

    def _apply_options_visibility(self):
        if self.opt_visible.get():
            self.opt_btn.config(text="Options  ▾")
            self.opt_frame.grid(row=5, column=0, columnspan=3, sticky="ew")
        else:
            self.opt_btn.config(text="Options  ▸")
            self.opt_frame.grid_forget()

    # -------------------------------------------------------------- Events
    def browse(self, var):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if path:
            var.set(os.path.normpath(path))

    def _on_drop(self, event, var):
        paths = parse_drop(event.data)
        if paths:
            var.set(os.path.normpath(paths[0]))
        return "copy"

    def pick_output(self):
        initial = self.output_var.get() or "index-updated.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save updated index as", defaultextension=".xlsx",
            initialfile=os.path.basename(initial),
            initialdir=os.path.dirname(initial) or None,
            filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.output_var.set(os.path.normpath(path))

    def pick_color(self, key, btn):
        _, hexcolor = colorchooser.askcolor(color=self.settings[key],
                                            parent=self.root)
        if hexcolor:
            self.settings[key] = hexcolor
            btn.config(bg=hexcolor)

    def _on_files_changed(self, *_):
        cur, inp = self.current_var.get().strip(), self.input_var.get().strip()
        ready = bool(cur and inp and os.path.isfile(cur) and os.path.isfile(inp))
        self.run_btn.config(state="normal" if ready else "disabled")

        if cur and os.path.isfile(cur) and not self.output_var.get().strip():
            p = Path(cur)
            self.output_var.set(str(p.with_name(p.stem + " (updated).xlsx")))

        if ready:
            self.status_var.set("Reading columns...")
            threading.Thread(target=self._load_headers, args=(cur, inp),
                             daemon=True).start()
        else:
            self.status_var.set(self._ready_text())

    def _load_headers(self, cur, inp):
        try:
            headers = shared_headers(inp, cur)
        except Exception as e:
            self.root.after(0, lambda: self.status_var.set(
                f"Could not read files: {e}"))
            return

        def apply():
            self.anchor_box.config(values=[AUTO] + sorted(headers))
            if self.anchor_var.get() not in [AUTO] + headers:
                self.anchor_var.set(AUTO)
            self.status_var.set(
                f"Ready — {len(headers)} shared columns found")
        self.root.after(0, apply)

    def _ready_text(self):
        base = "Ready"
        if not HAS_DND:
            base += "  (drag & drop unavailable — install tkinterdnd2)"
        return base

    # ----------------------------------------------------------------- Run
    def run(self):
        cur = self.current_var.get().strip()
        inp = self.input_var.get().strip()
        out = self.output_var.get().strip()
        if not out:
            messagebox.showwarning(APP_NAME, "Choose an output file first.")
            return
        if os.path.normcase(out) in (os.path.normcase(cur), os.path.normcase(inp)):
            messagebox.showwarning(
                APP_NAME, "The output file must differ from the input files.")
            return
        if os.path.exists(out) and not messagebox.askyesno(
                APP_NAME, f"{os.path.basename(out)} already exists.\nOverwrite?"):
            return

        anchor = self.anchor_var.get()
        key_override = None if anchor == AUTO else [anchor]

        def confirm(titles):
            if not titles:
                return messagebox.askyesno(
                    APP_NAME, "No new rows found — the index is already "
                    "up to date.\n\nSave an unchanged copy anyway?")
            shown = [str(t) for t in titles[:15]]
            more = len(titles) - len(shown)
            msg = f"{len(titles)} new row(s) will be added:\n\n" + \
                  "\n".join(f"  + {t}" for t in shown)
            if more > 0:
                msg += f"\n  ... and {more} more"
            msg += "\n\nContinue and save?"
            return messagebox.askyesno(APP_NAME, msg)

        self.run_btn.config(state="disabled")
        self.status_var.set("Working...")
        self.root.config(cursor="watch")
        self.root.update_idletasks()
        try:
            result = run_update(
                inp, cur, out,
                key_override=key_override,
                highlight_cols=self.hl_var.get(),
                fill_color=self.settings["fill_color"] if self.fill_on.get() else None,
                bold=self.bold_on.get(),
                underline=None if self.underline_var.get() == "none"
                          else self.underline_var.get(),
                font_color=self.settings["font_color"] if self.fc_on.get() else None,
                confirm=confirm,
                log=lambda *a: None)
        except Exception:
            self.status_var.set("Failed")
            messagebox.showerror(APP_NAME, "Update failed:\n\n"
                                 + traceback.format_exc(limit=3))
            return
        finally:
            self.root.config(cursor="")
            self.run_btn.config(state="normal")

        if result is None:
            self.status_var.set("Cancelled — nothing saved")
            return

        key = ", ".join(result["key"])
        self.status_var.set(
            f"Done — {len(result['new'])} new row(s), anchor: {key}")
        if messagebox.askyesno(
                APP_NAME,
                f"Saved: {os.path.basename(out)}\n"
                f"New rows: {len(result['new'])}\n"
                f"Anchor used: {key} ({result['key_reason']})\n\n"
                "Open the file now?"):
            os.startfile(out)

    def on_close(self):
        self.settings.update({
            "highlight_cols": self.hl_var.get(),
            "fill_enabled": self.fill_on.get(),
            "bold": self.bold_on.get(),
            "underline": self.underline_var.get(),
            "font_color_enabled": self.fc_on.get(),
            "options_expanded": self.opt_visible.get(),
        })
        save_settings(self.settings)
        self.root.destroy()


def main():
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
