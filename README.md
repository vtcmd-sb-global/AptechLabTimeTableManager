# Aptech Lab Timetable & Faculty Workload Manager

Desktop application that replaces the manual multi-sheet Excel “GLS Labstatus” workbook.

**One source of truth → every report updates automatically.**

## Features

- Central allocation form (enter data once)
- Automatic faculty workload calculation
- Automatic career / semester summaries + percentages
- Automatic MWF vs TTS batch & student totals
- Automatic STC / SP counts
- Faculty × Career matrix
- Visiting faculty list
- Lab capacity warnings
- Schedule conflict detection (same lab + time + day group)
- One-click Excel report export that follows the **GLS Labstatus template** structure
- Fully offline, single-file SQLite database

## Excel template fidelity

The official workbook `templates/GLS_Labstatus_template.xlsx` is the visual master.

Export produces three sheets:

| Sheet | Purpose |
|-------|---------|
| **Dashboard** | Career summary, MWF/TTS totals, faculty workload table, visiting faculty, faculty × career matrix – same information blocks that appear at the bottom of each monthly sheet in the template |
| **Allocations** | Flat filterable list of every allocation (audit / import friendly) |
| **Lab Status Grid** | Multi-row-per-lab grid that mirrors the monthly sheets: time slots across the top (B…G), each split into M/W/F and T/T/S blocks of 6 columns, with faculty / batch / module / date / career+students stacked vertically exactly as in the template |

Data is generated dynamically from the SQLite database; layout, headers, column grouping and summary tables follow the template.

## Quick Start

```bash
pip install -r requirements.txt
python seed_sample_data.py      # optional – loads demo rows from August-26 sheet
python main.py
```

## How it works

1. **Lookups** (Labs, Faculty, Careers, Time Slots, Day Groups) are managed once.
2. **Allocations** are the only transactional data you enter. One teacher can have many allocations (multiple careers / batches) — each save creates a new row and never overwrites another.
3. Course/series (AI, PMTZ, DM …) is derived automatically from the batch code prefix (e.g. `AI-202412C2` → AI). There is no separate Course Title field.
4. Every dashboard number and Excel sheet is calculated live from the `allocations` table.

## Project layout

```
aptech_timetable/
├── main.py                 # Entry point
├── database.py             # SQLite schema + connection
├── models.py               # CRUD
├── calculations.py         # All aggregates & reports
├── export_excel.py         # Template-aligned Excel exporter
├── seed_sample_data.py     # Demo data (from August-26 sheet)
├── ui/
│   └── main_window.py      # Tkinter GUI
├── templates/
│   └── GLS_Labstatus_template.xlsx   # Official master template
├── data/                   # timetable.db lives here
├── Output/                 # Generated reports
├── requirements.txt
└── README.md
```

## Building a Windows .exe

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pyinstaller --noconfirm --windowed --name AptechTimetable main.py
```

## Notes on the template

- Monthly sheets in the official file (May-25 … August-26) use a dense merged-cell layout.
- The exporter recreates that layout programmatically so it stays correct when labs, slots or allocations change.
- The original template file is shipped under `templates/` for reference and future exact-clone enhancements.

---

Internal use – Aptech Computer Education / GLS.
