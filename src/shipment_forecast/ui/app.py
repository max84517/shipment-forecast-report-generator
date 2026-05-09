"""
Shipment Forecast Report Generator — Main UI
CustomTkinter dark-mode interface.
"""
from __future__ import annotations

import re
import sys
import threading
from pathlib import Path
from typing import Callable

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk

import json

from shipment_forecast.paths import HISTORY_DIR, OUTPUT_DIR, REPORT_DIR, CONFIG_FILE, ensure_dirs
import shipment_forecast.processing.consolidate as consolidate_mod

# ── theme ─────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

MONTH_ABBRS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

SUBFOLDER_NAMES = {
    "monthly": "Monthly forecast",
    "spending": "Spending and rebate",
}


# ── helpers ───────────────────────────────────────────────────────────────────

def _popup_on_top(dialog: ctk.CTkToplevel) -> None:
    """Ensure a CTkToplevel appears in front and focused."""
    dialog.lift()
    dialog.focus_force()
    dialog.grab_set()


def _latest_excel(folder: Path) -> Path | None:
    """Return the most recently modified Excel file in folder."""
    excels = list(folder.glob("*.xls*"))
    if not excels:
        return None
    return max(excels, key=lambda f: f.stat().st_mtime)


def _list_excels(folder: Path) -> list[Path]:
    return sorted(folder.glob("*.xls*"), key=lambda f: f.name)


def _detect_suppliers(root: Path) -> list[str]:
    return sorted([d.name for d in root.iterdir() if d.is_dir()])


# ── Supplier row widget ───────────────────────────────────────────────────────

class SupplierRow(ctk.CTkFrame):
    """One row per supplier: checkbox + segmented source button + file dropdown."""

    def __init__(self, master, supplier: str, supplier_root: Path, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.supplier = supplier
        self.supplier_root = supplier_root
        self.columnconfigure(0, minsize=170)
        self.columnconfigure(1, minsize=155)
        self.columnconfigure(2, weight=1)

        self._enabled = ctk.BooleanVar(value=True)
        self._source_key = "monthly"  # "monthly" | "spending"
        self._selected_file: Path | None = None
        self._available_files: list[Path] = []

        # ── checkbox ──────────────────────────────────────────────────────────
        self.chk = ctk.CTkCheckBox(
            self, text=supplier, variable=self._enabled,
            width=160, font=("Arial", 12), command=self._on_toggle,
        )
        self.chk.grid(row=0, column=0, padx=(8, 4), sticky="w")

        # ── source type segmented button ──────────────────────────────────────
        self.src_btn = ctk.CTkSegmentedButton(
            self, values=["Forecast", "Shipment"],
            command=self._on_source_change,
            width=148, font=("Arial", 11),
            selected_color="#1f6aa5", selected_hover_color="#1a5a8a",
        )
        self.src_btn.set("Forecast")
        self.src_btn.grid(row=0, column=1, padx=(4, 8), sticky="w")

        # ── file dropdown button ───────────────────────────────────────────────
        self._file_var = ctk.StringVar(value="—")
        self.file_btn = ctk.CTkButton(
            self, textvariable=self._file_var,
            fg_color=("#2b2b2b", "#2b2b2b"), hover_color=("#3a3a3a", "#3a3a3a"),
            anchor="w", font=("Arial", 11), command=self._show_file_menu,
        )
        self.file_btn.grid(row=0, column=2, padx=(0, 8), sticky="ew")

        self._refresh_file_list()

    # ── internals ──────────────────────────────────────────────────────────────

    def _source_folder(self) -> Path:
        target_name = SUBFOLDER_NAMES[self._source_key].lower()
        supplier_dir = self.supplier_root / self.supplier
        for d in supplier_dir.iterdir():
            if d.is_dir() and d.name.lower() == target_name:
                return d
        # fallback: fuzzy match
        for d in supplier_dir.iterdir():
            if d.is_dir() and target_name.split()[0] in d.name.lower():
                return d
        return supplier_dir  # fallback

    def _refresh_file_list(self):
        folder = self._source_folder()
        self._available_files = _list_excels(folder)
        latest = _latest_excel(folder)
        if latest:
            self._selected_file = latest
            self._file_var.set(latest.name)
        else:
            self._selected_file = None
            self._file_var.set("(no Excel found)")

    def _on_source_change(self, value: str):
        self._source_key = "monthly" if value == "Forecast" else "spending"
        self._refresh_file_list()

    def _on_toggle(self):
        state = "normal" if self._enabled.get() else "disabled"
        self.src_btn.configure(state=state)
        self.file_btn.configure(state=state)

    def _show_file_menu(self):
        if not self._available_files:
            return
        menu = tk.Menu(self, tearoff=0, bg="#2b2b2b", fg="white",
                       activebackground="#1f6aa5", activeforeground="white")
        for f in self._available_files:
            menu.add_command(
                label=f.name,
                command=lambda p=f: self._select_file(p),
            )
        try:
            x = self.file_btn.winfo_rootx()
            y = self.file_btn.winfo_rooty() + self.file_btn.winfo_height()
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def _select_file(self, path: Path):
        self._selected_file = path
        self._file_var.set(path.name)

    # ── public ────────────────────────────────────────────────────────────────

    def is_enabled(self) -> bool:
        return self._enabled.get()

    def selected_file(self) -> Path | None:
        return self._selected_file if self._enabled.get() else None


# ── Path selector row ─────────────────────────────────────────────────────────

class PathSelector(ctk.CTkFrame):
    def __init__(self, master, label: str, default: Path, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.columnconfigure(1, weight=1)
        self._var = ctk.StringVar(value=str(default))
        ctk.CTkLabel(self, text=label, width=130, anchor="w",
                     font=("Arial", 11)).grid(row=0, column=0, padx=(4, 0))
        ctk.CTkEntry(self, textvariable=self._var,
                     font=("Arial", 11)).grid(row=0, column=1, padx=4, sticky="ew")
        ctk.CTkButton(self, text="Browse", width=70, height=28,
                      command=self._browse).grid(row=0, column=2, padx=(0, 4))

    def _browse(self):
        chosen = filedialog.askdirectory(initialdir=self._var.get())
        if chosen:
            self._var.set(chosen)
    @property
    def path(self) -> Path:
        return Path(self._var.get())


# ── Merge history dialog ──────────────────────────────────────────────────────

class MergeDialog(ctk.CTkToplevel):
    """Show available history files; let user pick which to merge."""

    def __init__(self, master, history_files: list[Path], output_dir: Path, on_confirm: Callable[[Path], None]):
        super().__init__(master)
        self.title("Merge History Files")
        self.resizable(False, False)
        self._on_confirm = on_confirm
        self._output_dir = output_dir
        self._vars: dict[Path, ctk.BooleanVar] = {}
        self._history_files = history_files

        ctk.CTkLabel(self, text="Select months to merge:", font=("Arial", 13, "bold")).pack(padx=20, pady=(14, 4))

        scroll = ctk.CTkScrollableFrame(self, width=400, height=200)
        scroll.pack(padx=20, pady=4, fill="both", expand=True)

        for f in history_files:
            var = ctk.BooleanVar(value=True)
            self._vars[f] = var
            ctk.CTkCheckBox(scroll, text=f.name, variable=var).pack(anchor="w", pady=2)

        # progress bar
        self._prog = ctk.CTkProgressBar(self, width=360)
        self._prog.pack(padx=20, pady=(6, 0))
        self._prog.set(0)
        self._prog_label = ctk.CTkLabel(self, text="", font=("Arial", 11))
        self._prog_label.pack()

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=10)
        self._merge_btn = ctk.CTkButton(btn_frame, text="Merge", command=self._merge)
        self._merge_btn.pack(side="left", padx=8)
        ctk.CTkButton(btn_frame, text="Cancel", fg_color="gray40",
                      command=self.destroy).pack(side="left", padx=8)

        _popup_on_top(self)

    def _merge(self):
        selected = [p for p, v in self._vars.items() if v.get()]
        if not selected:
            messagebox.showwarning("Warning", "No files selected.", parent=self)
            return
        self._merge_btn.configure(state="disabled")
        total = len(selected)

        def _work():
            try:
                out = consolidate_mod.merge_and_save(
                    selected,
                    output_dir=self._output_dir,
                    progress_cb=lambda i: self.after(
                        0, lambda i=i: self._update_prog(
                            i / total, f"Reading {selected[i].stem}..."
                        )
                    ),
                )
                self.after(0, lambda: self._finish(out))
            except Exception as exc:
                msg = str(exc)
                self.after(0, lambda m=msg: (
                    messagebox.showerror("Error", m, parent=self),
                    self.destroy(),
                ))

        self._prog.configure(mode="indeterminate")
        self._prog.start()
        self._prog_label.configure(text="Merging...")
        threading.Thread(target=_work, daemon=True).start()

    def _update_prog(self, value: float, label: str):
        self._prog.configure(mode="determinate")
        self._prog.stop()
        self._prog.set(value)
        self._prog_label.configure(text=label)

    def _finish(self, out: Path):
        self._prog.configure(mode="determinate")
        self._prog.stop()
        self._prog.set(1.0)
        self._prog_label.configure(text="Done!")
        self.destroy()
        self._on_confirm(out)


# ── Report Dialog ─────────────────────────────────────────────────────────────


class ReportDialog(ctk.CTkToplevel):
    """Let user pick a history file and reorder suppliers, then generate pivot report."""

    def __init__(
        self,
        master,
        history_files: list[Path],
        supplier_order: list[str],
        on_confirm: Callable[[Path, list[str]], None],
    ):
        super().__init__(master)
        self.title("Generate Report")
        self.resizable(False, False)
        self._history_files = history_files
        self._on_confirm = on_confirm

        # ── file selection ────────────────────────────────────────────────────
        ctk.CTkLabel(self, text="History File:", anchor="w",
                     font=("Arial", 12, "bold")).pack(padx=18, pady=(16, 4), fill="x")
        self._file_var = ctk.StringVar(
            value=history_files[-1].name if history_files else ""
        )
        ctk.CTkOptionMenu(
            self, values=[f.name for f in reversed(history_files)],
            variable=self._file_var, width=360,
        ).pack(padx=18, pady=(0, 12))

        # ── supplier order ────────────────────────────────────────────────────
        ctk.CTkLabel(self, text="Supplier Order  (drag to reorder):", anchor="w",
                     font=("Arial", 12, "bold")).pack(padx=18, pady=(0, 4), fill="x")

        lb_frame = tk.Frame(self, bg="#2b2b2b", relief="flat", bd=0)
        lb_frame.pack(padx=18, pady=(0, 12), fill="x")

        self._listbox = tk.Listbox(
            lb_frame,
            selectmode=tk.SINGLE,
            bg="#2b2b2b", fg="white",
            selectbackground="#1f6aa5", selectforeground="white",
            relief="flat", borderwidth=0, highlightthickness=1,
            highlightcolor="#555555", highlightbackground="#444444",
            font=("Arial", 12), height=len(supplier_order),
            activestyle="none",
        )
        self._listbox.pack(padx=4, pady=4, fill="x")
        for s in supplier_order:
            self._listbox.insert(tk.END, s)

        self._drag_src: int | None = None
        self._listbox.bind("<ButtonPress-1>", self._drag_start)
        self._listbox.bind("<B1-Motion>", self._drag_motion)

        # ── confirm button ────────────────────────────────────────────────────
        ctk.CTkButton(self, text="Generate", width=160, height=34,
                      command=self._confirm).pack(pady=(0, 16))

        _popup_on_top(self)

    def _drag_start(self, event: tk.Event):
        self._drag_src = self._listbox.nearest(event.y)

    def _drag_motion(self, event: tk.Event):
        dst = self._listbox.nearest(event.y)
        if self._drag_src is None or dst == self._drag_src:
            return
        item = self._listbox.get(self._drag_src)
        self._listbox.delete(self._drag_src)
        self._listbox.insert(dst, item)
        self._listbox.selection_clear(0, tk.END)
        self._listbox.selection_set(dst)
        self._drag_src = dst

    def _confirm(self):
        file_name = self._file_var.get()
        history_path = next(
            (f for f in self._history_files if f.name == file_name), None
        )
        if not history_path:
            return
        supplier_order = list(self._listbox.get(0, tk.END))
        self.destroy()
        self._on_confirm(history_path, supplier_order)


# ── Main UI ───────────────────────────────────────────────────────────────────

# ── config helpers ───────────────────────────────────────────────────────────

def _load_config() -> dict:
    try:
        if CONFIG_FILE.exists():
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def _save_config(data: dict) -> None:
    try:
        CONFIG_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:
        pass


# ── Main UI ───────────────────────────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ensure_dirs()
        self.title("Shipment Forecast Report Generator")
        self.geometry("780x590")
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._supplier_root: Path | None = None
        self._supplier_rows: list[SupplierRow] = []
        self._cfg = _load_config()

        self._build_ui()
        self._apply_config()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        P = 14  # standard horizontal padding

        # ── supplier root row ─────────────────────────────────────────────────
        top = ctk.CTkFrame(self)
        top.pack(fill="x", padx=P, pady=(12, 4))
        top.columnconfigure(1, weight=1)

        ctk.CTkLabel(top, text="Supplier Root:", width=120, anchor="w",
                     font=("Arial", 12)).grid(row=0, column=0, padx=(6, 4), pady=6)
        self._root_var = ctk.StringVar(value="")
        ctk.CTkEntry(top, textvariable=self._root_var,
                     font=("Arial", 11)).grid(row=0, column=1, padx=4, pady=6, sticky="ew")
        ctk.CTkButton(top, text="Browse", width=75, height=30,
                      command=self._browse_root).grid(row=0, column=2, padx=4, pady=6)
        ctk.CTkButton(top, text="Load Suppliers", width=115, height=30,
                      command=self._load_suppliers).grid(row=0, column=3, padx=(4, 6), pady=6)

        # ── supplier list ─────────────────────────────────────────────────────
        self._supplier_frame_outer = ctk.CTkScrollableFrame(
            self, label_text="Suppliers", height=190,
            label_font=("Arial", 12, "bold"),
        )
        self._supplier_frame_outer.pack(fill="both", padx=P, pady=(2, 4), expand=False)

        # ── output path selectors ─────────────────────────────────────────────
        paths_frame = ctk.CTkFrame(self)
        paths_frame.pack(fill="x", padx=P, pady=(0, 4))
        paths_frame.columnconfigure(0, weight=1)
        self._output_sel = PathSelector(paths_frame, "Data Output:", OUTPUT_DIR)
        self._output_sel.grid(row=0, column=0, sticky="ew", padx=6, pady=2)
        self._report_sel = PathSelector(paths_frame, "Report Output:", REPORT_DIR)
        self._report_sel.grid(row=1, column=0, sticky="ew", padx=6, pady=2)

        # ── progress bar ──────────────────────────────────────────────────────
        prog_frame = ctk.CTkFrame(self, fg_color="transparent")
        prog_frame.pack(fill="x", padx=P, pady=(2, 0))
        self._prog_bar = ctk.CTkProgressBar(prog_frame)
        self._prog_bar.pack(fill="x", side="left", expand=True, padx=(0, 8))
        self._prog_bar.set(0)
        self._prog_label = ctk.CTkLabel(prog_frame, text="", width=200, anchor="w",
                                        font=("Arial", 11))
        self._prog_label.pack(side="left")

        # ── log area ──────────────────────────────────────────────────────────
        self._log = ctk.CTkTextbox(self, height=85, state="disabled", wrap="word",
                                   font=("Consolas", 11))
        self._log.pack(fill="x", padx=P, pady=(4, 2))

        # ── action buttons ────────────────────────────────────────────────────
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(pady=8)
        ctk.CTkButton(btn_row, text="Consolidate Monthly Data", width=210, height=34,
                      font=("Arial", 12),
                      command=self._start_consolidate).pack(side="left", padx=12)
        ctk.CTkButton(btn_row, text="Merge Data", width=140, height=34,
                      font=("Arial", 12), fg_color="#2a6496",
                      command=self._start_merge).pack(side="left", padx=12)
        ctk.CTkButton(btn_row, text="Generate Report", width=160, height=34,
                      font=("Arial", 12), fg_color="#2e7d32",
                      command=self._start_report).pack(side="left", padx=12)

    # ── config load / save ────────────────────────────────────────────────────

    def _apply_config(self):
        if "supplier_root" in self._cfg:
            self._root_var.set(self._cfg["supplier_root"])
        if "output_path" in self._cfg:
            self._output_sel._var.set(self._cfg["output_path"])
        if "report_path" in self._cfg:
            self._report_sel._var.set(self._cfg["report_path"])

    def _persist_config(self):
        self._cfg["supplier_root"] = self._root_var.get()
        self._cfg["output_path"] = str(self._output_sel.path)
        self._cfg["report_path"] = str(self._report_sel.path)
        _save_config(self._cfg)

    # ── browse / load suppliers ───────────────────────────────────────────────

    def _browse_root(self):
        chosen = filedialog.askdirectory(initialdir=self._root_var.get() or None)
        if chosen:
            self._root_var.set(chosen)
            self._persist_config()

    def _load_suppliers(self):
        root = Path(self._root_var.get())
        if not root.is_dir():
            messagebox.showerror("Error", "Please select a valid supplier root directory.")
            return
        self._supplier_root = root

        # Clear existing rows
        for w in self._supplier_frame_outer.winfo_children():
            w.destroy()
        self._supplier_rows.clear()

        suppliers = _detect_suppliers(root)
        if not suppliers:
            messagebox.showinfo("Info", "No sub-folders found in the selected directory.")
            return

        # Header row
        hdr = ctk.CTkFrame(self._supplier_frame_outer, fg_color="transparent")
        hdr.pack(fill="x", pady=(0, 2))
        hdr.columnconfigure(0, minsize=170)
        hdr.columnconfigure(1, minsize=155)
        hdr.columnconfigure(2, weight=1)
        ctk.CTkLabel(hdr, text="Supplier", anchor="w", width=160,
                     font=("Arial", 11, "bold")).grid(row=0, column=0, padx=(8, 4), sticky="w")
        ctk.CTkLabel(hdr, text="Source Type", anchor="w", width=148,
                     font=("Arial", 11, "bold")).grid(row=0, column=1, padx=(4, 8), sticky="w")
        ctk.CTkLabel(hdr, text="Selected File", anchor="w",
                     font=("Arial", 11, "bold")).grid(row=0, column=2, padx=(0, 8), sticky="w")

        for sup in suppliers:
            row = SupplierRow(self._supplier_frame_outer, sup, root)
            row.pack(fill="x", pady=1)
            self._supplier_rows.append(row)

        self._persist_config()
        self._log_msg(f"Detected {len(suppliers)} supplier(s): {', '.join(suppliers)}")

    # ── logging / progress ────────────────────────────────────────────────────

    def _log_msg(self, msg: str):
        self._log.configure(state="normal")
        self._log.insert("end", msg + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")

    def _set_progress(self, value: float, label: str):
        """value < 0 = indeterminate; 0.0–1.0 = determinate. Call from main thread."""
        if value < 0:
            self._prog_bar.configure(mode="indeterminate")
            self._prog_bar.start()
        else:
            self._prog_bar.configure(mode="determinate")
            self._prog_bar.stop()
            self._prog_bar.set(value)
        self._prog_label.configure(text=label)

    # ── consolidate ───────────────────────────────────────────────────────────

    def _start_consolidate(self):
        selected_files = [r.selected_file() for r in self._supplier_rows if r.selected_file()]
        if not selected_files:
            messagebox.showwarning("Warning", "No supplier files selected.")
            return

        self._log_msg("Scanning FY sheets...")
        self._set_progress(-1, "Scanning sheets...")

        def _run():
            try:
                common = consolidate_mod.common_fy_sheets(selected_files)
            except Exception as exc:
                msg = str(exc)
                self.after(0, lambda m=msg: (
                    self._set_progress(0, "Error"),
                    messagebox.showerror("Error", m),
                ))
                return

            if not common:
                self.after(0, lambda: (
                    self._set_progress(0, "Idle"),
                    messagebox.showerror(
                        "No Common Sheets",
                        "No common FY sheets found across the selected supplier files.\n"
                        "Please check your source files."
                    ),
                ))
                return

            self.after(0, lambda: (
                self._set_progress(0, "Ready"),
                self._show_fy_dialog(common, selected_files),
            ))

        threading.Thread(target=_run, daemon=True).start()

    def _show_fy_dialog(self, common_sheets: list[str], source_files: list[Path]):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Select FY Sheet & Start Month")
        dialog.resizable(False, False)

        ctk.CTkLabel(dialog, text="FY Sheet:", width=120, anchor="w").grid(row=0, column=0, padx=14, pady=10)
        fy_var = ctk.StringVar(value=common_sheets[0])
        ctk.CTkOptionMenu(dialog, values=common_sheets, variable=fy_var, width=130).grid(row=0, column=1, padx=8, pady=10)

        ctk.CTkLabel(dialog, text="Recent Month:", width=120, anchor="w").grid(row=1, column=0, padx=14, pady=10)
        import datetime
        cur_month = datetime.date.today().strftime("%b")
        month_var = ctk.StringVar(value=cur_month)
        ctk.CTkOptionMenu(dialog, values=MONTH_ABBRS, variable=month_var, width=130).grid(row=1, column=1, padx=8, pady=10)

        def _confirm():
            fy = fy_var.get()
            month = month_var.get()
            dialog.destroy()
            self._run_consolidate(source_files, fy, month)

        ctk.CTkButton(dialog, text="Confirm", command=_confirm).grid(
            row=2, column=0, columnspan=2, pady=12)

        dialog.lift()
        dialog.focus_force()
        dialog.grab_set()

    def _run_consolidate(self, source_files: list[Path], fy_sheet: str, start_month: str):
        self._log_msg(f"Consolidating {fy_sheet} from {start_month}...")
        total = len(source_files)

        def _work():
            try:
                import pandas as pd

                # Step 1: copy
                self.after(0, lambda: self._set_progress(0.05, "Copying source files..."))
                copied = consolidate_mod.clear_and_copy_sources(source_files)
                self.after(0, lambda: self._log_msg(f"Copied {len(copied)} file(s) to source_data."))

                # Step 2: read each file
                dfs = []
                for i, f in enumerate(copied):
                    pct = 0.05 + 0.75 * i / max(total, 1)
                    lbl = f"Reading {f.name} ({i+1}/{total})..."
                    self.after(0, lambda p=pct, l=lbl: self._set_progress(p, l))
                    try:
                        df = consolidate_mod.read_supplier_sheet(f, fy_sheet, start_month)
                        if not df.empty:
                            df["Source File"] = f.name
                            dfs.append(df)
                    except Exception as exc:
                        _e, _n = str(exc), f.name
                        self.after(0, lambda e=_e, n=_n:
                                   self._log_msg(f"[WARN] {n}: {e}"))

                self.after(0, lambda: self._set_progress(0.82, "Merging data..."))

                if not dfs:
                    self.after(0, lambda: (
                        self._set_progress(0, "Idle"),
                        messagebox.showwarning(
                            "Warning", "Consolidation produced no data. Check source files."),
                    ))
                    return

                merged = pd.concat(dfs, ignore_index=True)

                # Add "Forecast Date" column, e.g. "FY26 05"
                import calendar
                month_num = str(
                    [m.lower() for m in calendar.month_abbr].index(start_month.lower())
                ).zfill(2)
                merged["Forecast Date"] = f"{fy_sheet.upper()} {month_num}"

                # Step 3: save
                self.after(0, lambda: self._set_progress(0.92, "Saving to history..."))
                out_path = consolidate_mod.save_to_history(merged, fy_sheet, start_month)
                self.after(0, lambda: (
                    self._set_progress(1.0, "Done!"),
                    self._log_msg(f"Saved: {out_path.name}"),
                ))
                self.after(300, lambda: self._prompt_merge_after_consolidate())

            except Exception as exc:
                import traceback
                _msg, _tb = str(exc), traceback.format_exc()
                self.after(0, lambda m=_msg, t=_tb: (
                    self._set_progress(0, "Error"),
                    self._log_msg(f"[ERROR] {m}"),
                    messagebox.showerror("Error", f"{m}\n\n{t}"),
                ))

        threading.Thread(target=_work, daemon=True).start()

    def _prompt_merge_after_consolidate(self):
        history_files = consolidate_mod.list_history_files()
        if not history_files:
            messagebox.showinfo("Done", "Consolidation complete. No history files to merge.")
            return
        answer = messagebox.askyesno(
            "Merge?",
            "Consolidation complete!\n\nDo you want to merge history files now?",
        )
        if answer:
            self._open_merge_dialog()

    # ── merge ─────────────────────────────────────────────────────────────────

    def _start_merge(self):
        self._open_merge_dialog()

    def _open_merge_dialog(self):
        history_files = consolidate_mod.list_history_files()
        if not history_files:
            messagebox.showinfo("Info", "No Rolling Forecast files found in history folder.")
            return

        def _on_done(out: Path):
            self._log_msg(f"Merged output saved: {out}")
            messagebox.showinfo("Done", f"Merged file saved:\n{out}")

        MergeDialog(self, history_files, self._output_sel.path, _on_done)

    # ── report generation ──────────────────────────────────────────────────────

    def _start_report(self):
        history_files = consolidate_mod.list_history_files()
        if not history_files:
            messagebox.showwarning("Warning", "No history files found. Please consolidate first.")
            return

        # Build supplier list sorted alphabetically from the selected history file.
        default_order: list[str] = []
        try:
            import pandas as pd
            df_check = pd.read_excel(history_files[-1], sheet_name=0, usecols=["GTK Suppliers"])
            default_order = sorted(
                df_check["GTK Suppliers"].dropna().unique().tolist(),
                key=lambda s: s.lower(),
            )
        except Exception:
            pass

        ReportDialog(self, history_files, default_order, self._run_report)

    def _run_report(self, history_path: Path, supplier_order: list[str]):
        self._set_progress(-1, "Generating report...")
        self._log_msg(f"Generating report from {history_path.name}...")

        def _work():
            try:
                out = consolidate_mod.generate_report(
                    history_path, supplier_order,
                    report_dir=self._report_sel.path,
                )
                self.after(0, lambda p=str(out): (
                    self._set_progress(1.0, "Done!"),
                    self._log_msg(f"Report saved: {p}"),
                    messagebox.showinfo("Done", f"Report saved:\n{p}"),
                ))
            except Exception as exc:
                import traceback
                _msg, _tb = str(exc), traceback.format_exc()
                self.after(0, lambda m=_msg, t=_tb: (
                    self._set_progress(0, "Error"),
                    self._log_msg(f"[ERROR] {m}"),
                    messagebox.showerror("Error", f"{m}\n\n{t}"),
                ))

        threading.Thread(target=_work, daemon=True).start()

    # ── close ─────────────────────────────────────────────────────────────────

    def _on_close(self):
        self._persist_config()
        self.destroy()
        sys.exit(0)


# ── entry point ───────────────────────────────────────────────────────────────

def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
