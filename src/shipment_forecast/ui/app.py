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

from shipment_forecast.paths import HISTORY_DIR, OUTPUT_DIR, REPORT_DIR, ensure_dirs
from shipment_forecast import processing

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
    """One row per supplier: checkbox + source-type slider + detected file dropdown."""

    SOURCE_TYPES = [SUBFOLDER_NAMES["monthly"], SUBFOLDER_NAMES["spending"]]

    def __init__(self, master, supplier: str, supplier_root: Path, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.supplier = supplier
        self.supplier_root = supplier_root

        self._enabled = ctk.BooleanVar(value=True)
        self._source_idx = ctk.IntVar(value=0)  # 0=monthly, 1=spending
        self._selected_file: Path | None = None
        self._available_files: list[Path] = []

        # ── checkbox ──────────────────────────────────────────────────────────
        self.chk = ctk.CTkCheckBox(
            self, text=supplier, variable=self._enabled,
            width=160, command=self._on_toggle,
        )
        self.chk.grid(row=0, column=0, padx=(0, 10), sticky="w")

        # ── source type slider ────────────────────────────────────────────────
        slider_frame = ctk.CTkFrame(self, fg_color="transparent")
        slider_frame.grid(row=0, column=1, padx=(0, 10))

        ctk.CTkLabel(slider_frame, text="Monthly", width=60, anchor="e").grid(row=0, column=0)
        self.slider = ctk.CTkSlider(
            slider_frame, from_=0, to=1, number_of_steps=1, width=60,
            variable=self._source_idx, command=self._on_slider,
        )
        self.slider.grid(row=0, column=1, padx=4)
        ctk.CTkLabel(slider_frame, text="Spending", width=60, anchor="w").grid(row=0, column=2)

        # ── file dropdown button ───────────────────────────────────────────────
        self._file_var = ctk.StringVar(value="—")
        self.file_btn = ctk.CTkButton(
            self, textvariable=self._file_var, width=260,
            fg_color=("#2b2b2b", "#2b2b2b"), hover_color=("#3a3a3a", "#3a3a3a"),
            anchor="w", command=self._show_file_menu,
        )
        self.file_btn.grid(row=0, column=2, padx=(0, 4))

        self._refresh_file_list()

    # ── internals ──────────────────────────────────────────────────────────────

    def _source_folder(self) -> Path:
        key = "monthly" if self._source_idx.get() == 0 else "spending"
        target_name = SUBFOLDER_NAMES[key].lower()
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

    def _on_slider(self, _=None):
        self._refresh_file_list()

    def _on_toggle(self):
        state = "normal" if self._enabled.get() else "disabled"
        self.slider.configure(state=state)
        self.file_btn.configure(state=state)

    def _show_file_menu(self):
        if not self._available_files:
            return
        menu = tk.Menu(self, tearoff=0, bg="#2b2b2b", fg="white",
                       activebackground="#3a3a3a", activeforeground="white")
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
        self._var = ctk.StringVar(value=str(default))
        ctk.CTkLabel(self, text=label, width=200, anchor="w").grid(row=0, column=0)
        ctk.CTkEntry(self, textvariable=self._var, width=320).grid(row=0, column=1, padx=4)
        ctk.CTkButton(self, text="Browse", width=70, command=self._browse).grid(row=0, column=2)

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

    def __init__(self, master, history_files: list[Path], on_confirm: Callable[[list[Path]], None]):
        super().__init__(master)
        self.title("Merge History Files")
        self.resizable(False, False)
        self._on_confirm = on_confirm
        self._vars: dict[Path, ctk.BooleanVar] = {}

        ctk.CTkLabel(self, text="Select months to merge:", font=("Arial", 13, "bold")).pack(padx=20, pady=(16, 6))

        scroll = ctk.CTkScrollableFrame(self, width=380, height=240)
        scroll.pack(padx=20, pady=4, fill="both", expand=True)

        for f in history_files:
            var = ctk.BooleanVar(value=True)
            self._vars[f] = var
            ctk.CTkCheckBox(scroll, text=f.name, variable=var).pack(anchor="w", pady=2)

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=12)
        ctk.CTkButton(btn_frame, text="Merge", command=self._merge).pack(side="left", padx=8)
        ctk.CTkButton(btn_frame, text="Ignore / Cancel", fg_color="gray40",
                      command=self.destroy).pack(side="left", padx=8)

    def _merge(self):
        selected = [p for p, v in self._vars.items() if v.get()]
        self.destroy()
        if selected:
            self._on_confirm(selected)


# ── Main UI ───────────────────────────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ensure_dirs()
        self.title("Shipment Forecast Report Generator")
        self.geometry("950x720")
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._supplier_root: Path | None = None
        self._supplier_rows: list[SupplierRow] = []

        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        # ── top: source root path ─────────────────────────────────────────────
        top = ctk.CTkFrame(self)
        top.pack(fill="x", padx=16, pady=(14, 4))

        ctk.CTkLabel(top, text="Supplier Source Root:", width=180, anchor="w").grid(row=0, column=0, padx=(4, 6))
        self._root_var = ctk.StringVar(value="")
        ctk.CTkEntry(top, textvariable=self._root_var, width=340).grid(row=0, column=1, padx=4)
        ctk.CTkButton(top, text="Browse", width=70, command=self._browse_root).grid(row=0, column=2, padx=4)
        ctk.CTkButton(top, text="Load Suppliers", width=110, command=self._load_suppliers).grid(row=0, column=3, padx=8)

        # ── supplier list ──────────────────────────────────────────────────────
        self._supplier_frame_outer = ctk.CTkScrollableFrame(self, label_text="Suppliers", height=240)
        self._supplier_frame_outer.pack(fill="both", padx=16, pady=6, expand=False)

        # ── path selectors ─────────────────────────────────────────────────────
        paths_frame = ctk.CTkFrame(self)
        paths_frame.pack(fill="x", padx=16, pady=4)

        self._output_sel = PathSelector(paths_frame, "Data Output Path:", OUTPUT_DIR)
        self._output_sel.pack(fill="x", padx=8, pady=2)
        self._report_sel = PathSelector(paths_frame, "Report Output Path:", REPORT_DIR)
        self._report_sel.pack(fill="x", padx=8, pady=2)

        # ── status / log area ──────────────────────────────────────────────────
        self._log = ctk.CTkTextbox(self, height=130, state="disabled", wrap="word")
        self._log.pack(fill="x", padx=16, pady=4)

        # ── action buttons ─────────────────────────────────────────────────────
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(pady=8)
        ctk.CTkButton(btn_row, text="Consolidate Monthly Data", width=200,
                      command=self._start_consolidate).pack(side="left", padx=10)
        ctk.CTkButton(btn_row, text="Merge Data", width=140, fg_color="#2a6496",
                      command=self._start_merge).pack(side="left", padx=10)

    # ── browse / load suppliers ───────────────────────────────────────────────

    def _browse_root(self):
        chosen = filedialog.askdirectory()
        if chosen:
            self._root_var.set(chosen)

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
        hdr.pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(hdr, text="Supplier", width=160, anchor="w", font=("Arial", 11, "bold")).grid(row=0, column=0)
        ctk.CTkLabel(hdr, text="Source Type", width=160, anchor="w", font=("Arial", 11, "bold")).grid(row=0, column=1)
        ctk.CTkLabel(hdr, text="Selected File", width=260, anchor="w", font=("Arial", 11, "bold")).grid(row=0, column=2)

        for sup in suppliers:
            row = SupplierRow(self._supplier_frame_outer, sup, root)
            row.pack(fill="x", pady=2)
            self._supplier_rows.append(row)

        self._log_msg(f"Detected {len(suppliers)} supplier(s): {', '.join(suppliers)}")

    # ── logging ───────────────────────────────────────────────────────────────

    def _log_msg(self, msg: str):
        self._log.configure(state="normal")
        self._log.insert("end", msg + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")

    # ── consolidate ───────────────────────────────────────────────────────────

    def _start_consolidate(self):
        selected_files = [r.selected_file() for r in self._supplier_rows if r.selected_file()]
        if not selected_files:
            messagebox.showwarning("Warning", "No supplier files selected.")
            return

        self._log_msg("Finding common FY sheets...")

        def _run():
            from shipment_forecast.processing.consolidate import (
                common_fy_sheets, clear_and_copy_sources,
                consolidate_monthly, save_to_history,
            )
            common = common_fy_sheets(selected_files)
            if not common:
                self.after(0, lambda: messagebox.showerror(
                    "No Common Sheets",
                    "No common FY sheets found across the selected supplier files.\n"
                    "Please check your source files."
                ))
                return

            self.after(0, lambda: self._show_fy_dialog(common, selected_files))

        threading.Thread(target=_run, daemon=True).start()

    def _show_fy_dialog(self, common_sheets: list[str], source_files: list[Path]):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Select FY Sheet & Month")
        dialog.resizable(False, False)
        dialog.grab_set()

        ctk.CTkLabel(dialog, text="FY Sheet:", width=100, anchor="w").grid(row=0, column=0, padx=12, pady=8)
        fy_var = ctk.StringVar(value=common_sheets[0])
        fy_menu = ctk.CTkOptionMenu(dialog, values=common_sheets, variable=fy_var, width=120)
        fy_menu.grid(row=0, column=1, padx=8, pady=8)

        ctk.CTkLabel(dialog, text="Recent Month:", width=100, anchor="w").grid(row=1, column=0, padx=12, pady=8)
        import datetime
        cur_month = datetime.date.today().strftime("%b")
        month_var = ctk.StringVar(value=cur_month)
        month_menu = ctk.CTkOptionMenu(dialog, values=MONTH_ABBRS, variable=month_var, width=120)
        month_menu.grid(row=1, column=1, padx=8, pady=8)

        def _confirm():
            dialog.destroy()
            self._run_consolidate(source_files, fy_var.get(), month_var.get())

        ctk.CTkButton(dialog, text="Confirm", command=_confirm).grid(
            row=2, column=0, columnspan=2, pady=12)

    def _run_consolidate(self, source_files: list[Path], fy_sheet: str, start_month: str):
        self._log_msg(f"Starting consolidation: {fy_sheet}, from {start_month}...")

        def _work():
            from shipment_forecast.processing.consolidate import (
                clear_and_copy_sources, consolidate_monthly, save_to_history,
            )
            try:
                copied = clear_and_copy_sources(source_files)
                self.after(0, lambda: self._log_msg(f"Copied {len(copied)} file(s) to source_data."))

                df = consolidate_monthly(copied, fy_sheet, start_month)
                if df.empty:
                    self.after(0, lambda: messagebox.showwarning(
                        "Warning", "Consolidation produced no data. Check source files."))
                    return

                out_path = save_to_history(df, fy_sheet, start_month)
                self.after(0, lambda: self._log_msg(f"Saved: {out_path.name}"))
                self.after(0, lambda: self._prompt_merge_after_consolidate())
            except Exception as exc:
                self.after(0, lambda: messagebox.showerror("Error", str(exc)))

        threading.Thread(target=_work, daemon=True).start()

    def _prompt_merge_after_consolidate(self):
        history_files = processing.consolidate.list_history_files()
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
        history_files = processing.consolidate.list_history_files()
        if not history_files:
            messagebox.showinfo("Info", "No Rolling Forecast files found in history folder.")
            return

        def _do_merge(selected: list[Path]):
            def _work():
                from shipment_forecast.processing.consolidate import merge_and_save
                try:
                    out = merge_and_save(selected)
                    self.after(0, lambda: self._log_msg(f"Merged output saved: {out}"))
                    self.after(0, lambda: messagebox.showinfo("Done", f"Merged file saved:\n{out}"))
                except Exception as exc:
                    self.after(0, lambda: messagebox.showerror("Error", str(exc)))

            threading.Thread(target=_work, daemon=True).start()

        MergeDialog(self, history_files, _do_merge)

    # ── close ─────────────────────────────────────────────────────────────────

    def _on_close(self):
        self.destroy()
        sys.exit(0)


# ── entry point ───────────────────────────────────────────────────────────────

def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
