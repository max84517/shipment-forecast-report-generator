"""Project path resolution — works both in dev (poetry run) and frozen (PyInstaller)."""
import sys
from pathlib import Path


def _project_root() -> Path:
    if getattr(sys, "frozen", False):
        # PyInstaller onedir: executable is in the dist folder
        return Path(sys.executable).parent
    # Dev mode: src/shipment_forecast/paths.py → project root is 2 levels up from src
    return Path(__file__).parent.parent.parent


ROOT = _project_root()
DATA_DIR = ROOT / "data"
SOURCE_DATA_DIR = DATA_DIR / "source_data"
REPORT_DIR = DATA_DIR / "report"
HISTORY_DIR = DATA_DIR / "history"
OUTPUT_DIR = DATA_DIR / "output"
CONFIG_FILE = ROOT / "config.json"


def ensure_dirs() -> None:
    for d in (SOURCE_DATA_DIR, REPORT_DIR, HISTORY_DIR, OUTPUT_DIR):
        d.mkdir(parents=True, exist_ok=True)
