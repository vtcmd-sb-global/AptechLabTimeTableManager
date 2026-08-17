#!/usr/bin/env python3
"""
Generate Excel reports that mirror the GLS Labstatus template structure.

- Dashboard sheet: career summary, MWF/TTS totals, faculty workload, matrix
- Allocations sheet: flat audit list
- Lab Status Grid sheet: multi-row-per-lab layout matching the monthly sheets
  in the official GLS Labstatus workbook (time slots across columns,
  M/W/F + T/T/S sub-columns, faculty / batch / module / date / career rows)
"""

from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from models import (
    list_allocations, list_labs, list_time_slots, list_day_groups,
    derive_course_from_batch,
)
from calculations import full_dashboard


# ---------------------------------------------------------------------------
# Styles (closely aligned with template visual weight)
# ---------------------------------------------------------------------------
THIN = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
MED = Border(
    left=Side(style="medium"),
    right=Side(style="medium"),
    top=Side(style="medium"),
    bottom=Side(style="medium"),
)

HEADER_FILL = PatternFill("solid", fgColor="1A365D")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
SECTION_FILL = PatternFill("solid", fgColor="2B6CB0")
SECTION_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=12)
SUBHEAD_FILL = PatternFill("solid", fgColor="BEE3F8")
SUBHEAD_FONT = Font(bold=True, name="Calibri", size=10)
LAB_FILL = PatternFill("solid", fgColor="EBF8FF")
CAREER_FILL = PatternFill("solid", fgColor="FEFCBF")
TOTAL_FILL = PatternFill("solid", fgColor="C6F6D5")

COURSE_FONT = Font(bold=True, color="C53030", name="Calibri", size=9)
NORMAL = Font(name="Calibri", size=9)
BOLD = Font(bold=True, name="Calibri", size=9)
SMALL = Font(name="Calibri", size=8)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
TOP_CENTER = Alignment(horizontal="center", vertical="top", wrap_text=True)


def _style_header(cell, fill=HEADER_FILL, font=HEADER_FONT):
    cell.fill = fill
    cell.font = font
    cell.alignment = CENTER
    cell.border = THIN


def _style_cell(cell, font=NORMAL, align=CENTER, fill=None):
    cell.font = font
    cell.alignment = align
    cell.border = THIN
    if fill:
        cell.fill = fill


def _apply_border_range(ws, r1, c1, r2, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = THIN


# ---------------------------------------------------------------------------
# Main export
# ---------------------------------------------------------------------------

def export_workbook(output_path: Optional[Path] = None) -> Path:
    if output_path is None:
        stamp = datetime.now().strftime("%Y-%m-%d")
        output_path = Path(__file__).parent / "Output" / f"Aptech_LabStatus_{stamp}.xlsx"
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    dash = full_dashboard()
    allocs = list_allocations(active_only=True)

    _build_dashboard(wb.active, dash)
    _build_allocations(wb.create_sheet("Allocations"), allocs)
    _build_lab_status_grid(wb.create_sheet("Lab Status Grid"), allocs, dash)

    wb.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# Sheet 1 – Dashboard
# ---------------------------------------------------------------------------

def _build_dashboard(ws, dash: dict):
    ws.title = "Dashboard"

    ws.merge_cells("A1:H1")
    ws["A1"] = f"Aptech Lab Status & Faculty Workload – {datetime.now().strftime('%d %B %Y')}"
    ws["A1"].font = Font(bold=True, size=16, color="1A365D", name="Calibri")
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    # Career / Semester Summary
    row = 3
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    ws.cell(row=row, column=1, value="CAREER / SEMESTER SUMMARY")
    _style_header(ws.cell(row=row, column=1), SECTION_FILL, SECTION_FONT)
    for c in range(2, 6):
        _style_header(ws.cell(row=row, column=c), SECTION_FILL, SECTION_FONT)

    row = 4
    for col, h in enumerate(["Semester", "No. of Batches", "No. of Students", "%Age"], 1):
        _style_header(ws.cell(row=row, column=col, value=h), SUBHEAD_FILL, SUBHEAD_FONT)

    for item in dash["career_summary"]:
        row += 1
        vals = [
            item["code"],
            item["batches"],
            item["students"],
            f"{item['pct'] * 100:.1f}%" if item["code"] != "Total" else "",
        ]
        for col, v in enumerate(vals, 1):
            fill = TOTAL_FILL if item["code"] == "Total" else None
            _style_cell(ws.cell(row=row, column=col, value=v),
                        BOLD if item["code"] == "Total" else NORMAL, fill=fill)

    # MWF / TTS Totals
    row += 2
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    ws.cell(row=row, column=1, value="MWF / TTS TOTALS")
    _style_header(ws.cell(row=row, column=1), SECTION_FILL, SECTION_FONT)

    row += 1
    for col, h in enumerate(["Day Group", "Batches", "Students"], 1):
        _style_header(ws.cell(row=row, column=col, value=h), SUBHEAD_FILL, SUBHEAD_FONT)

    for dg, totals in dash["day_group_totals"].items():
        row += 1
        for col, v in enumerate([dg, totals["batches"], totals["students"]], 1):
            _style_cell(ws.cell(row=row, column=col, value=v))

    # STC / SP
    row += 2
    stc = dash["stc_sp"]
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    ws.cell(row=row, column=1, value="STC / SP")
    _style_header(ws.cell(row=row, column=1), SECTION_FILL, SECTION_FONT)

    row += 1
    for label, key in [
        ("STC Batches", "stc_batches"),
        ("STC Students", "stc_students"),
        ("SP Batches", "sp_batches"),
        ("SP Students", "sp_students"),
    ]:
        ws.cell(row=row, column=1, value=label).font = BOLD
        ws.cell(row=row, column=2, value=stc[key])
        row += 1

    # Faculty Workload (matches template INITIAL / STATUS / Faculty ACCP / Load …)
    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.cell(row=row, column=1, value="FACULTY WORKLOAD (matches template summary)")
    _style_header(ws.cell(row=row, column=1), SECTION_FILL, SECTION_FONT)

    row += 1
    headers = ["Initial", "Status", "Faculty ACCP", "Load", "Career", "STC", "No. Students", "TTS/MWF"]
    for col, h in enumerate(headers, 1):
        _style_header(ws.cell(row=row, column=col, value=h), SUBHEAD_FILL, SUBHEAD_FONT)

    for item in dash["faculty_workload"]:
        row += 1
        vals = [
            item["initials"],
            item["status"],
            item["full_name"],
            item["load"],
            item["career_summary"],
            item["stc"],
            item["students"],
            item["timings"],
        ]
        for col, v in enumerate(vals, 1):
            _style_cell(
                ws.cell(row=row, column=col, value=v),
                align=LEFT if col in (3, 5, 8) else CENTER,
            )

    # Visiting Faculty
    row += 2
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    ws.cell(row=row, column=1, value="VISITING FACULTY")
    _style_header(ws.cell(row=row, column=1), SECTION_FILL, SECTION_FONT)

    row += 1
    for col, h in enumerate(["Faculty", "Visiting No.", "Time"], 1):
        _style_header(ws.cell(row=row, column=col, value=h), SUBHEAD_FILL, SUBHEAD_FONT)

    for item in dash["visiting_faculty"]:
        row += 1
        for col, v in enumerate([item["name"], item["visiting_no"], item["time"]], 1):
            _style_cell(
                ws.cell(row=row, column=col, value=v),
                align=LEFT if col in (1, 3) else CENTER,
            )

    # Faculty × Career matrix
    row += 2
    codes, matrix_rows = dash["faculty_career_matrix"]
    cols_needed = 2 + len(codes)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max(cols_needed, 4))
    ws.cell(row=row, column=1, value="FACULTY × CAREER DISTRIBUTION")
    _style_header(ws.cell(row=row, column=1), SECTION_FILL, SECTION_FONT)

    row += 1
    headers = ["Faculty"] + list(codes) + ["Total"]
    for col, h in enumerate(headers, 1):
        _style_header(ws.cell(row=row, column=col, value=h), SUBHEAD_FILL, SUBHEAD_FONT)

    for mrow in matrix_rows:
        row += 1
        vals = [mrow["faculty"]] + [mrow.get(c, 0) for c in codes] + [mrow["Total"]]
        for col, v in enumerate(vals, 1):
            _style_cell(
                ws.cell(row=row, column=col, value=v),
                align=LEFT if col == 1 else CENTER,
            )

    widths = [12, 10, 22, 10, 14, 8, 12, 42]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------------------------------------------------------------------------
# Sheet 2 – Allocations (flat list)
# ---------------------------------------------------------------------------

def _build_allocations(ws, allocs: list):
    headers = [
        "Lab", "Day Group", "Time Slot", "Faculty", "Status",
        "Batch Code", "Module", "Career", "Students",
        "Module Start", "Actual Start", "Admission Open", "Notes",
    ]
    for col, h in enumerate(headers, 1):
        _style_header(ws.cell(row=1, column=col, value=h))

    for r_idx, a in enumerate(allocs, 2):
        vals = [
            a.lab_code,
            a.day_group_code,
            a.time_slot_label,
            f"{a.faculty_initials} – {a.faculty_name}",
            a.faculty_status,
            a.batch_code,
            a.module_name,
            a.career_code,
            a.students,
            a.module_start_date,
            a.actual_start_date,
            "Yes" if a.is_admission_open else "",
            a.notes or "",
        ]
        for c_idx, v in enumerate(vals, 1):
            _style_cell(
                ws.cell(row=r_idx, column=c_idx, value=v),
                NORMAL,
                LEFT if c_idx in (4, 7, 13) else CENTER,
            )

    for i in range(1, 14):
        ws.column_dimensions[get_column_letter(i)].width = 14
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["G"].width = 16
    ws.column_dimensions["M"].width = 20
    ws.auto_filter.ref = f"A1:M{max(1, len(allocs) + 1)}"
    ws.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# Sheet 3 – Lab Status Grid (template-faithful layout)
# ---------------------------------------------------------------------------

def _build_lab_status_grid(ws, allocs: list, dash: dict):
    """
    Layout mirrors the monthly sheets in GLS Labstatus:

      Lab | B (9:00-11:00)          | C (11:00-1:00)           | …
          | M/W/F      | T/T/S      | M/W/F      | T/T/S      | …
          | (6 cols)   | (6 cols)   | …

    Each lab occupies a block of rows:
      r+0  PC count + (SERIES) Faculty initials
      r+1  Batch code
      r+2  Lab label (L1…) + Module name
      r+3  Module start date  (or "Admission Open")
      r+4  Career code + student count (in offset columns)
      r+5  Actual / CI date (optional)
      r+6  blank separator
    """
    labs = list_labs()
    slots = list_time_slots()
    days = list_day_groups()

    day_mwf = next((d for d in days if d.code.upper() in ("MWF", "M/W/F")), None)
    day_tts = next((d for d in days if d.code.upper() in ("TTS", "T/T/S", "T/TS")), None)
    if day_tts is None:
        day_tts = next((d for d in days if "T" in d.code.upper()), None)

    # Index allocations: (lab_id, slot_id, day_id) → list[Allocation]
    grid: Dict[Tuple[int, int, int], list] = defaultdict(list)
    for a in allocs:
        grid[(a.lab_id, a.time_slot_id, a.day_group_id)].append(a)

    # ---- Title ----
    total_cols = 1 + len(slots) * 12  # 6 cols MWF + 6 cols TTS per slot
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=min(total_cols, 40))
    ws["A1"] = f"Batch Status Report – {datetime.now().strftime('%Y-%m-%d')}"
    ws["A1"].font = Font(bold=True, size=14, color="1A365D", name="Calibri")
    ws.row_dimensions[1].height = 24

    # ---- Header rows 2–3 ----
    # Row 2: time-slot labels (merged across 12 columns)
    # Row 3: M/W/F | T/T/S sub-headers (each spanning 6 columns)
    ws.cell(row=2, column=1, value="Lab")
    _style_header(ws.cell(row=2, column=1))
    ws.cell(row=3, column=1, value="")
    _style_header(ws.cell(row=3, column=1), SUBHEAD_FILL, SUBHEAD_FONT)

    # Map (slot_id, day_id) → starting column of the 6-col block
    slot_day_start: Dict[Tuple[int, int], int] = {}
    col = 2

    for slot in slots:
        start_col = col
        # Slot header spanning full 12 columns
        cell = ws.cell(row=2, column=col, value=slot.label)
        _style_header(cell)
        for c in range(col, col + 12):
            _style_header(ws.cell(row=2, column=c))
        try:
            ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + 11)
        except Exception:
            pass

        # Day sub-headers
        for day_obj, label in ((day_mwf, "M/W/F"), (day_tts, "T/T/S")):
            if day_obj is None:
                col += 6
                continue
            cell = ws.cell(row=3, column=col, value=label)
            _style_header(cell, SUBHEAD_FILL, SUBHEAD_FONT)
            for c in range(col, col + 6):
                _style_header(ws.cell(row=3, column=c), SUBHEAD_FILL, SUBHEAD_FONT)
            try:
                ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + 5)
            except Exception:
                pass
            slot_day_start[(slot.id, day_obj.id)] = col
            col += 6

    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 18

    # ---- Lab blocks ----
    # 7 rows per lab (matching typical template density)
    ROWS_PER_LAB = 7
    data_start = 4

    def _series(a) -> str:
        return (a.course_title or derive_course_from_batch(a.batch_code) or "").strip()

    def _write_block(ws, base_row: int, start_c: int, items: list):
        """Write one allocation (or first of several) into a 6-col day block."""
        if not items:
            # Leave empty but bordered
            for r in range(base_row, base_row + 6):
                for c in range(start_c, start_c + 6):
                    ws.cell(r, c).border = THIN
            return

        a = items[0]  # primary; extras noted with separator if needed
        series = _series(a)
        fac_line = f"({series})  {a.faculty_initials}" if series else a.faculty_initials

        # r+0 faculty
        cell = ws.cell(base_row, start_c, value=fac_line)
        cell.font = COURSE_FONT
        cell.alignment = TOP_CENTER
        cell.border = THIN

        # r+1 batch
        cell = ws.cell(base_row + 1, start_c, value=a.batch_code or "")
        cell.font = BOLD
        cell.alignment = CENTER
        cell.border = THIN

        # r+2 module
        cell = ws.cell(base_row + 2, start_c, value=a.module_name or "")
        cell.font = NORMAL
        cell.alignment = CENTER
        cell.border = THIN

        # r+3 date or Admission Open
        if a.is_admission_open and not a.module_start_date:
            date_val = "Admission Open"
        else:
            date_val = a.module_start_date or ""
        cell = ws.cell(base_row + 3, start_c, value=date_val)
        cell.font = SMALL
        cell.alignment = CENTER
        cell.border = THIN

        # r+4 career + students (students in col+4 to mirror template F/L/X offsets)
        cell = ws.cell(base_row + 4, start_c, value=a.career_code or "")
        cell.font = BOLD
        cell.alignment = CENTER
        cell.border = THIN
        cell.fill = CAREER_FILL

        stu_cell = ws.cell(base_row + 4, start_c + 4, value=a.students if a.students else "")
        stu_cell.font = BOLD
        stu_cell.alignment = CENTER
        stu_cell.border = THIN
        stu_cell.fill = CAREER_FILL

        # r+5 actual start / extra
        extra = a.actual_start_date or ""
        if len(items) > 1:
            extra = (extra + " | +" + str(len(items) - 1) + " more").strip(" |")
        cell = ws.cell(base_row + 5, start_c, value=extra)
        cell.font = SMALL
        cell.alignment = CENTER
        cell.border = THIN

        # Fill remaining cells of the 6-col block with borders
        for r in range(base_row, base_row + 6):
            for c in range(start_c, start_c + 6):
                if ws.cell(r, c).border.left.style is None:
                    ws.cell(r, c).border = THIN

    for lab_idx, lab in enumerate(labs):
        base = data_start + lab_idx * ROWS_PER_LAB

        # Lab identity column
        # Row 0: PC count
        cell = ws.cell(base, 1, value=lab.pcs)
        _style_cell(cell, BOLD, CENTER, LAB_FILL)
        # Row 2: Lx label
        cell = ws.cell(base + 2, 1, value=lab.lab_code)
        _style_cell(cell, BOLD, CENTER, LAB_FILL)
        for r in range(base, base + ROWS_PER_LAB):
            if ws.cell(r, 1).value is None:
                ws.cell(r, 1).border = THIN
                ws.cell(r, 1).fill = LAB_FILL

        # Day blocks
        for (sid, did), start_c in slot_day_start.items():
            items = grid.get((lab.id, sid, did), [])
            _write_block(ws, base, start_c, items)

        for r in range(base, base + ROWS_PER_LAB):
            ws.row_dimensions[r].height = 16

    last_data_row = data_start + len(labs) * ROWS_PER_LAB - 1

    # ---- Column widths ----
    ws.column_dimensions["A"].width = 8
    for c in range(2, col):
        ws.column_dimensions[get_column_letter(c)].width = 5.5

    ws.freeze_panes = "B4"

    # ---- Bottom summary section (mirrors template faculty + career tables) ----
    summary_row = last_data_row + 3

    # Career totals line
    ws.merge_cells(start_row=summary_row, start_column=1, end_row=summary_row, end_column=8)
    ws.cell(row=summary_row, column=1, value="CAREER / SEMESTER TOTALS (from live data)")
    _style_header(ws.cell(row=summary_row, column=1), SECTION_FILL, SECTION_FONT)

    summary_row += 1
    for col, h in enumerate(["Semester", "Batches", "Students", "%Age"], 1):
        _style_header(ws.cell(row=summary_row, column=col, value=h), SUBHEAD_FILL, SUBHEAD_FONT)

    for item in dash["career_summary"]:
        summary_row += 1
        vals = [
            item["code"],
            item["batches"],
            item["students"],
            f"{item['pct'] * 100:.1f}%" if item["code"] != "Total" else "",
        ]
        for c, v in enumerate(vals, 1):
            fill = TOTAL_FILL if item["code"] == "Total" else None
            _style_cell(ws.cell(row=summary_row, column=c, value=v),
                        BOLD if item["code"] == "Total" else NORMAL, fill=fill)

    # Faculty load table
    summary_row += 2
    ws.merge_cells(start_row=summary_row, start_column=1, end_row=summary_row, end_column=8)
    ws.cell(row=summary_row, column=1, value="FACULTY LOAD SUMMARY")
    _style_header(ws.cell(row=summary_row, column=1), SECTION_FILL, SECTION_FONT)

    summary_row += 1
    headers = ["Initial", "Status", "Faculty", "Load", "Career", "STC", "Students", "Timings"]
    for col, h in enumerate(headers, 1):
        _style_header(ws.cell(row=summary_row, column=col, value=h), SUBHEAD_FILL, SUBHEAD_FONT)

    for item in dash["faculty_workload"]:
        summary_row += 1
        vals = [
            item["initials"], item["status"], item["full_name"],
            item["load"], item["career_summary"], item["stc"],
            item["students"], item["timings"],
        ]
        for c, v in enumerate(vals, 1):
            _style_cell(
                ws.cell(row=summary_row, column=c, value=v),
                align=LEFT if c in (3, 5, 8) else CENTER,
            )

    # STC / SP footer
    summary_row += 2
    stc = dash["stc_sp"]
    ws.cell(row=summary_row, column=1, value="STC TOTAL").font = BOLD
    ws.cell(row=summary_row, column=2, value=stc["stc_students"])
    ws.cell(row=summary_row, column=3, value="SP TOTAL").font = BOLD
    ws.cell(row=summary_row, column=4, value=stc["sp_students"])
    summary_row += 1
    total_students = sum(
        item["students"] for item in dash["career_summary"] if item["code"] != "Total"
    )
    ws.cell(row=summary_row, column=1, value="TOTAL STUDENTS").font = BOLD
    ws.cell(row=summary_row, column=2, value=total_students)
    ws.cell(row=summary_row, column=2).fill = TOTAL_FILL


if __name__ == "__main__":
    path = export_workbook()
    print(f"Wrote: {path}")
