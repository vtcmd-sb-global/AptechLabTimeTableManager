#!/usr/bin/env python3
"""Aptech Monthly BSR – entry point."""

from pathlib import Path
import sys

ROOT = Path(__file__).parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from database import init_db
from models import list_available_report_months


def ensure_historical_data():
    """Import multi-month Excel once if fewer than 2 months of data exist."""
    try:
        months = list_available_report_months()
        if len(months) >= 3:
            return
        from import_historical import import_historical, HIST_CANDIDATES
        src = None
        for c in HIST_CANDIDATES + [ROOT / "data" / "21-August-2026 GLS Labstatus.xlsx"]:
            if c.exists():
                src = c
                break
        if src:
            print(f"Importing historical BSR data from {src.name}…")
            summary = import_historical(src)
            print(f"Imported {summary['total_allocations']} allocations across {len(summary['months'])} months.")
    except Exception as e:
        print("Historical import skipped:", e)


def main():
    init_db()
    ensure_historical_data()
    from ui.main_window import main as ui_main
    ui_main()


if __name__ == "__main__":
    main()
