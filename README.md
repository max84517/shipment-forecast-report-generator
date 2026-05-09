# Shipment Forecast Report Generator

A desktop tool (CustomTkinter dark-mode UI) for consolidating supplier monthly forecast Excel files into rolling forecast reports.

---

## Features

- **Supplier auto-detection**: Point the tool at a root folder; it detects supplier sub-folders and lists them with checkboxes.
- **Per-supplier source type**: A slider per supplier toggles between **Monthly Forecast** and **Spending and Rebate** as the data source.
- **Smart file selection**: Automatically picks the most recently modified Excel in the target sub-folder; a dropdown lets you override.
- **Common FY sheet detection**: Finds FY sheets (e.g. `FY26`) present in **all** selected Excel files.
- **Flexible month filter**: Choose a "Recent Month" — only that month and later months (in fiscal order) are retained.
- **Structured data pipeline**:
  1. Copies source files to `data/source_data/` (cleared before each run).
  2. Reads the chosen FY sheet; fills down merged cells (Category / Segment / Series).
  3. Extracts feature columns + value columns (`Table Price`, `Unit Rebate`, `Q'ty`, `Rebate Amount`). Handles Excel files with duplicate header names caused by merged-cell forward-fill using positional indexing.
  4. Melts into long format (one row per feature combo × month).
  5. Adds `FY` (e.g. `FY26 Q2`), `Year`, `Month`, and `Forecast Date` (e.g. `FY26 05`) columns.
  6. Saves to `data/history/Rolling Forecast FYXX MM.xlsx`.
- **History merge**: Select any combination of saved Rolling Forecast files to merge into `data/output/forecast data.xlsx`.
- **Standalone Merge button**: Access the merge dialog at any time without re-running consolidation.
- **Generate Report**: Produces a pivot-table Excel with two sheets — **Keyboard** and **Peripheral** — each containing a Rebate Amount pivot and a Q'ty pivot stacked on the same sheet (3 blank rows apart). Columns are months in fiscal order (e.g. `May'26`, `Jun'26`) with Quarter subtotal columns (Excel SUM formulas) inserted after each quarter. Supplier order is alphabetical by default and can be reordered by drag-and-drop before confirming.
  - Rebate Amount cells use **Currency format** (`$#,##0.00`)
  - Quarter subtotal columns have **green background** (`#C1F0C8`), **thick black left/right borders**, and **bold** text
  - All cells use **Calibri 11** font

---

## Project Structure

```
shipment-forecast-report-generator/
├── src/
│   └── shipment_forecast/
│       ├── main.py               # Entry point
│       ├── paths.py              # Path resolution (dev + frozen)
│       ├── ui/
│       │   └── app.py            # CustomTkinter UI
│       └── processing/
│           └── consolidate.py    # Data processing logic
├── data/
│   ├── source_data/              # Temp copies of source Excel files
│   ├── report/                   # Report output
│   ├── history/                  # Rolling Forecast FYXX MM.xlsx files
│   └── output/                   # forecast data.xlsx (merged output)
├── pyproject.toml
├── poetry.toml
└── README.md
```

---

## Quick Start (No Python required)

1. Go to the [Releases page](https://github.com/max84517/shipment-forecast-report-generator/releases/latest)
2. Download **`ShipmentForecast.zip`**
3. Extract the zip — you will get a `ShipmentForecast\` folder
4. Run **`ShipmentForecast\ShipmentForecast.exe`**

> Keep the entire `ShipmentForecast\` folder together — the exe depends on the files next to it.

---

## Setup (Development)

### Prerequisites

- Python 3.11+
- [Poetry](https://python-poetry.org/)

### Install

```bash
poetry install
```

### Run

```bash
poetry run shipment-forecast
```

---

## Data Requirements

### Supplier Root Folder Structure

```
<Supplier Root>/
├── Liteon/
│   ├── Monthly forecast/
│   │   └── Liteon_Forecast_2026.xlsx
│   └── Spending and rebate/
│       └── Liteon_Spending_2026.xlsx
├── Acrox/
│   ├── Monthly forecast/
│   └── Spending and rebate/
...
```

### Excel Sheet Format

- Sheet names must match `FYxx` (e.g. `FY24`, `FY26`).
- **Row 2** is the header row.
- **Feature columns**: `SPM`, `GTK Suppliers`, `Category`, `Segment`, `Series`, `Production Year`, `Platforms`, `Product`, `Size`, `Color`, `Location`, `ODM`, `Supplier Part#`, `HP/ODM Part#`
- `Category`, `Segment`, `Series` may use merged cells — the tool forward-fills missing values.
- **Value columns**: any column whose name contains `Table Price`, `Unit Rebate`, `Q'ty`, or `Rebate Amount` followed by a 3-letter month abbreviation (e.g. `Table Price Apr`, `Mar Q'ty`).

### Fiscal Year Convention

| Months | Quarter |
|--------|---------|
| Nov, Dec, Jan | Q1 |
| Feb, Mar, Apr | Q2 |
| May, Jun, Jul | Q3 |
| Aug, Sep, Oct | Q4 |

Nov and Dec of `FY26` belong to calendar year **2025**; all other months belong to **2026**.

---

## Output Files

| File | Location | Description |
|------|----------|-------------|
| `Rolling Forecast FY26 05.xlsx` | `data/history/` | Consolidated data for one FY + start month |
| `forecast data.xlsx` | `data/output/` | Merged data from selected history files |
| `forecast pivot FY26 05.xlsx` | `data/report/` | Pivot report with Keyboard & Peripheral sheets (Rebate Amount + Q'ty) |

---

## License

Internal HP tool — not for external distribution.
