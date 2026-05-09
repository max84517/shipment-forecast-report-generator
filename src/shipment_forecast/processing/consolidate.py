"""
Data processing: read supplier Excel files, melt value columns,
consolidate, and write history + output files.
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path
from typing import Iterable

import openpyxl
import pandas as pd

from shipment_forecast.paths import SOURCE_DATA_DIR, HISTORY_DIR, OUTPUT_DIR

# ── constants ────────────────────────────────────────────────────────────────

MONTH_ORDER = ["Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct"]
MONTH_TO_INT = {m: i for i, m in enumerate(MONTH_ORDER)}  # Nov=0, Oct=11
MONTH_ABBR_RE = re.compile(
    r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b", re.IGNORECASE
)

FEATURE_COLS = [
    "SPM", "GTK Suppliers", "Category", "Segment", "Series",
    "Production Year", "Platforms", "Product", "Size", "Color",
    "Location", "ODM", "Supplier Part#", "HP/ODM Part#",
]

VALUE_KEYWORDS = ["Table Price", "Unit Rebate", "Q'ty", "Rebate Amount"]

FY_SHEET_RE = re.compile(r"^FY\d{2}$", re.IGNORECASE)

# ── helpers ──────────────────────────────────────────────────────────────────


def _normalize_col(col: str) -> str:
    return " ".join(str(col).split())


def _detect_fy_sheets(path: Path) -> list[str]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheets = [s for s in wb.sheetnames if FY_SHEET_RE.match(s)]
    wb.close()
    return sheets


def common_fy_sheets(excel_paths: list[Path]) -> list[str]:
    """Return FY sheets present in ALL provided Excel files, sorted desc."""
    if not excel_paths:
        return []
    sets = [set(_detect_fy_sheets(p)) for p in excel_paths]
    common = set.intersection(*sets)
    return sorted(common, key=lambda s: int(s[2:]), reverse=True)


def _month_abbr(col_name: str) -> str | None:
    """Extract 3-letter month abbreviation from a column name."""
    m = MONTH_ABBR_RE.search(col_name)
    if m:
        return m.group(0).capitalize()
    return None


def _value_type(col_name: str) -> str | None:
    for kw in VALUE_KEYWORDS:
        if kw.lower() in col_name.lower():
            # normalise key name
            if "table price" in col_name.lower():
                return "Table Price"
            if "unit rebate" in col_name.lower():
                return "Unit Rebate"
            if "q'ty" in col_name.lower() or "qty" in col_name.lower():
                return "Q'ty"
            if "rebate amount" in col_name.lower():
                return "Rebate Amount"
    return None


def _month_index_from_start(month_abbr: str, start_month: str) -> int:
    """Return months since start_month in the fiscal cycle (0 = start)."""
    start_idx = MONTH_TO_INT.get(start_month, 0)
    month_idx = MONTH_TO_INT.get(month_abbr, 0)
    return (month_idx - start_idx) % 12


def _fy_quarter(month_abbr: str) -> str:
    """Map month to Q1-Q4 in fiscal year (Nov=Q1 start)."""
    idx = MONTH_TO_INT[month_abbr]  # Nov=0, Oct=11
    quarter = idx // 3 + 1
    return f"Q{quarter}"


def _year_for_month(fy_str: str, month_abbr: str) -> int:
    fy_year = int(fy_str[2:]) + 2000  # FY26 -> 2026
    if month_abbr in ("Nov", "Dec"):
        return fy_year - 1
    return fy_year


# ── core reader ──────────────────────────────────────────────────────────────


def read_supplier_sheet(path: Path, sheet_name: str, start_month: str) -> pd.DataFrame:
    """
    Read one FY sheet from a supplier Excel file.
    Returns a melted DataFrame with feature + month columns.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name]

    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        wb.close()
        return pd.DataFrame()

    # Header is row index 1 (0-based), i.e. 2nd row
    raw_header = rows[1]
    # Build header list; forward-fill None (merged cells)
    headers: list[str] = []
    last = ""
    for cell in raw_header:
        val = str(cell).strip() if cell is not None else ""
        if val == "" or val == "None":
            val = last
        else:
            last = val
        headers.append(val)

    data_rows = list(rows[2:])

    # Trim trailing rows that are entirely None — openpyxl often returns thousands
    # of phantom rows from Excel's stored used-range; they explode after ffill × months.
    while data_rows and all(v is None for v in data_rows[-1]):
        data_rows.pop()

    if not data_rows:
        wb.close()
        return pd.DataFrame()

    df_raw = pd.DataFrame(data_rows, columns=headers)

    # ── identify feature and value columns ───────────────────────────────────
    # Match feature columns (flexible: check if header starts-with or equals)
    feature_map: dict[str, str] = {}  # actual_col -> canonical name
    for h in headers:
        h_clean = _normalize_col(h)
        for fc in FEATURE_COLS:
            if h_clean.lower() == fc.lower() or h_clean.lower().startswith(fc.lower()):
                feature_map[h] = fc
                break

    # Value columns: (actual_col) -> (value_type, month_abbr)
    value_map: dict[str, tuple[str, str]] = {}
    for h in headers:
        vt = _value_type(h)
        if vt:
            ma = _month_abbr(h)
            if ma:
                value_map[h] = (vt, ma)

    if not value_map:
        wb.close()
        return pd.DataFrame()

    # Filter to months >= start_month in fiscal cycle
    def keep_month(ma: str) -> bool:
        return _month_index_from_start(ma, start_month) >= 0  # keep all from start

    # months to keep (those at index >=0 from start, up to Oct)
    start_idx = MONTH_TO_INT[start_month]
    months_to_keep = {MONTH_ORDER[(start_idx + i) % 12] for i in range(12 - start_idx)}

    filtered_value_map = {k: v for k, v in value_map.items() if v[1] in months_to_keep}

    if not filtered_value_map:
        wb.close()
        return pd.DataFrame()

    # Build a tidy feature df (forward-fill merged cells in Category/Segment/Series)
    feat_cols_actual = list(feature_map.keys())
    # deduplicate: if same canonical name appears multiple times take first
    seen_canonical: set[str] = set()
    deduped_feat_cols: list[str] = []
    for col in feat_cols_actual:
        canon = feature_map[col]
        if canon not in seen_canonical:
            seen_canonical.add(canon)
            deduped_feat_cols.append(col)

    feat_df = df_raw[deduped_feat_cols].copy()
    feat_df.columns = [feature_map[c] for c in deduped_feat_cols]

    # Forward-fill merged cell columns
    for col in ("Category", "Segment", "Series"):
        if col in feat_df.columns:
            feat_df[col] = feat_df[col].replace("", pd.NA).ffill()

    # Drop rows where all feature cols are null/empty (phantom trailing rows)
    feat_df = feat_df.dropna(how="all")

    # Secondary guard: drop rows where neither GTK Suppliers nor Supplier Part#
    # has a real value — these are phantom / artifact rows (e.g. "Refresh" buttons).
    # Use GTK Suppliers keyword blocklist FIRST, then require at least one anchor col.
    _ARTIFACTS = {"refresh", "total", "note", "notes", "subtotal",
                  "click", "update", "button", "reset"}

    def _is_artifact(val) -> bool:
        if pd.isna(val):
            return False
        s = str(val).strip()
        return s.lower() in _ARTIFACTS

    if "GTK Suppliers" in feat_df.columns:
        feat_df = feat_df[~feat_df["GTK Suppliers"].apply(_is_artifact)]

    # Require at least one "anchor" identifier column to be non-null/non-empty.
    # This avoids dropping real rows where only one column is filled.
    anchor_cols = [c for c in ("GTK Suppliers", "Supplier Part#", "HP/ODM Part#", "ODM")
                   if c in feat_df.columns]
    if anchor_cols:
        has_anchor = feat_df[anchor_cols].replace("", pd.NA).notna().any(axis=1)
        feat_df = feat_df[has_anchor]

    if feat_df.empty:
        wb.close()
        return pd.DataFrame()

    # ── build value columns aligned to surviving feat_df rows ─────────────────
    # IMPORTANT: save original df_raw indices BEFORE any reset_index so that
    # value_df rows are always correctly matched to feat_df rows.
    surviving_idx = feat_df.index  # original positional indices into df_raw
    feat_df = feat_df.reset_index(drop=True)

    value_df = df_raw.iloc[surviving_idx].reset_index(drop=True)

    # ── pivot value cols: one row per (feature combo, month) ─────────────────
    # Group value columns by month, collect {month: {vtype: col}}
    month_vtype_col: dict[str, dict[str, str]] = {}
    for col, (vt, ma) in filtered_value_map.items():
        month_vtype_col.setdefault(ma, {})[vt] = col

    month_dfs: list[pd.DataFrame] = []
    for ma, vtype_cols in month_vtype_col.items():
        m_df = feat_df.copy()
        m_df["Month"] = ma
        for vt in VALUE_KEYWORDS:
            actual_col = vtype_cols.get(vt)
            if actual_col:
                col_data = value_df[actual_col] if actual_col in value_df.columns else pd.Series(dtype=float)
                m_df[vt] = pd.to_numeric(col_data, errors="coerce").fillna(0).values
            else:
                m_df[vt] = 0
        month_dfs.append(m_df)

    if not month_dfs:
        wb.close()
        return pd.DataFrame()

    result = pd.concat(month_dfs, ignore_index=True)

    # ── add FY / Quarter / Year cols ─────────────────────────────────────────
    fy = sheet_name.upper()  # e.g. FY26
    result["FY"] = result["Month"].apply(lambda m: f"{fy} {_fy_quarter(m)}")
    result["Year"] = result["Month"].apply(lambda m: _year_for_month(fy, m))

    wb.close()
    return result


# ── source data management ────────────────────────────────────────────────────


def clear_and_copy_sources(file_paths: list[Path]) -> list[Path]:
    """Clear source_data dir contents and copy given files into it.

    Avoids shutil.rmtree on the directory itself to prevent WinError 5
    on OneDrive-synced paths where the folder is locked by the sync client.
    """
    SOURCE_DATA_DIR.mkdir(parents=True, exist_ok=True)
    # Delete existing files one by one (don't remove the directory)
    for existing in SOURCE_DATA_DIR.iterdir():
        try:
            if existing.is_file():
                existing.unlink()
            elif existing.is_dir():
                shutil.rmtree(existing, ignore_errors=True)
        except OSError:
            pass  # skip locked files silently

    copied: list[Path] = []
    for src in file_paths:
        dst = SOURCE_DATA_DIR / src.name
        # If duplicate name, prefix with supplier folder name
        if dst.exists():
            dst = SOURCE_DATA_DIR / f"{src.parent.parent.name}_{src.name}"
        shutil.copy2(src, dst)
        copied.append(dst)
    return copied


# ── consolidation ─────────────────────────────────────────────────────────────


def consolidate_monthly(
    source_files: list[Path],
    fy_sheet: str,
    start_month: str,
) -> pd.DataFrame:
    """Read all source files for the given FY sheet and concatenate."""
    dfs: list[pd.DataFrame] = []
    for f in source_files:
        try:
            df = read_supplier_sheet(f, fy_sheet, start_month)
            if not df.empty:
                df["Source File"] = f.name
                dfs.append(df)
        except Exception as exc:
            print(f"[WARN] Failed to read {f.name}: {exc}")
    if not dfs:
        return pd.DataFrame()
    return pd.concat(dfs, ignore_index=True)


def save_to_history(df: pd.DataFrame, fy_sheet: str, start_month: str) -> Path:
    """Save consolidated df to history as 'Rolling Forecast FYXX MM.xlsx'."""
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    fy = fy_sheet.upper()  # FY26
    # Convert month abbr to zero-padded number
    month_num = str(MONTH_ORDER.index(start_month) + 1).zfill(2)
    # Fiscal month index: Nov=01 ... Oct=12; convert to calendar-like label
    # Actually the user wants MM = calendar month number of start_month
    import calendar
    cal_abbrs = [calendar.month_abbr[i] for i in range(1, 13)]
    if start_month in cal_abbrs:
        month_num = str(cal_abbrs.index(start_month) + 1).zfill(2)

    fname = f"Rolling Forecast {fy} {month_num}.xlsx"
    out_path = HISTORY_DIR / fname
    sheet_label = f"forecast based on {start_month}"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_label, index=False)
    return out_path


def list_history_files() -> list[Path]:
    """Return Rolling Forecast files in history dir."""
    pattern = re.compile(r"^Rolling Forecast FY\d{2} \d{2}\.xlsx$", re.IGNORECASE)
    return sorted(
        [f for f in HISTORY_DIR.glob("*.xlsx") if pattern.match(f.name)],
        key=lambda f: f.name,
    )


def merge_and_save(
    selected_paths: list[Path],
    progress_cb: "Callable[[int], None] | None" = None,
) -> Path:
    """Merge selected history files and write to output/forecast data.xlsx."""
    from typing import Callable  # noqa: F401 – used in type hint string above
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    dfs: list[pd.DataFrame] = []
    for i, p in enumerate(selected_paths):
        if progress_cb:
            progress_cb(i)
        xl = pd.ExcelFile(p, engine="openpyxl")
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            df["_source"] = p.stem
            dfs.append(df)
    if not dfs:
        return OUTPUT_DIR / "forecast data.xlsx"
    merged = pd.concat(dfs, ignore_index=True)
    merged.drop(columns=["_source"], inplace=True, errors="ignore")
    out = OUTPUT_DIR / "forecast data.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="Forecast Data", index=False)
    return out
