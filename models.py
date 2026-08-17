#!/usr/bin/env python3
"""
CRUD helpers and domain objects for Aptech Timetable.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple
from database import db_session


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class Lab:
    id: int
    lab_code: str
    pcs: int
    notes: Optional[str] = None
    active: bool = True


@dataclass
class Faculty:
    id: int
    initials: str
    full_name: str
    status: str          # P / V
    active: bool = True


@dataclass
class Career:
    id: int
    code: str
    name: Optional[str]
    is_stc: bool = False
    is_sp: bool = False
    sort_order: int = 100


@dataclass
class TimeSlot:
    id: int
    code: str
    label: str
    start_time: Optional[str]
    end_time: Optional[str]
    sort_order: int = 0


@dataclass
class DayGroup:
    id: int
    code: str
    label: str
    sort_order: int = 0


@dataclass
class Allocation:
    id: int
    faculty_id: int
    career_id: int
    lab_id: int
    time_slot_id: int
    day_group_id: int
    batch_code: Optional[str]
    course_title: Optional[str]
    module_name: Optional[str]
    students: int
    module_start_date: Optional[str]
    actual_start_date: Optional[str]
    notes: Optional[str]
    is_admission_open: bool
    is_active: bool
    # joined display fields
    faculty_initials: str = ""
    faculty_name: str = ""
    faculty_status: str = ""
    career_code: str = ""
    lab_code: str = ""
    lab_pcs: int = 0
    time_slot_label: str = ""
    day_group_code: str = ""


# ---------------------------------------------------------------------------
# Lookups
# ---------------------------------------------------------------------------

def list_labs(active_only: bool = True) -> List[Lab]:
    with db_session() as conn:
        q = "SELECT * FROM labs"
        if active_only:
            q += " WHERE active = 1"
        q += " ORDER BY lab_code"
        rows = conn.execute(q).fetchall()
    return [Lab(**dict(r)) for r in rows]


def list_faculty(active_only: bool = True) -> List[Faculty]:
    with db_session() as conn:
        q = "SELECT * FROM faculty"
        if active_only:
            q += " WHERE active = 1"
        q += " ORDER BY initials"
        rows = conn.execute(q).fetchall()
    return [Faculty(**dict(r)) for r in rows]


def list_careers() -> List[Career]:
    with db_session() as conn:
        rows = conn.execute(
            "SELECT * FROM careers ORDER BY sort_order, code"
        ).fetchall()
    return [
        Career(
            id=r["id"], code=r["code"], name=r["name"],
            is_stc=bool(r["is_stc"]), is_sp=bool(r["is_sp"]),
            sort_order=r["sort_order"],
        )
        for r in rows
    ]


def list_time_slots() -> List[TimeSlot]:
    with db_session() as conn:
        rows = conn.execute(
            "SELECT * FROM time_slots ORDER BY sort_order"
        ).fetchall()
    return [TimeSlot(**dict(r)) for r in rows]


def list_day_groups() -> List[DayGroup]:
    with db_session() as conn:
        rows = conn.execute(
            "SELECT * FROM day_groups ORDER BY sort_order"
        ).fetchall()
    return [DayGroup(**dict(r)) for r in rows]


# ---------------------------------------------------------------------------
# Faculty / Lab / Career helpers
# ---------------------------------------------------------------------------

def add_faculty(initials: str, full_name: str, status: str = "P") -> int:
    with db_session() as conn:
        cur = conn.execute(
            "INSERT INTO faculty (initials, full_name, status) VALUES (?, ?, ?)",
            (initials.strip().upper(), full_name.strip(), status.upper()),
        )
        return cur.lastrowid


def update_faculty(fid: int, initials: str, full_name: str, status: str) -> None:
    with db_session() as conn:
        conn.execute(
            "UPDATE faculty SET initials=?, full_name=?, status=? WHERE id=?",
            (initials.strip().upper(), full_name.strip(), status.upper(), fid),
        )


def add_lab(lab_code: str, pcs: int, notes: str = "") -> int:
    with db_session() as conn:
        cur = conn.execute(
            "INSERT INTO labs (lab_code, pcs, notes) VALUES (?, ?, ?)",
            (lab_code.strip().upper(), pcs, notes),
        )
        return cur.lastrowid


def add_career(code: str, name: str = "", is_stc: bool = False, is_sp: bool = False) -> int:
    with db_session() as conn:
        cur = conn.execute(
            "INSERT INTO careers (code, name, is_stc, is_sp) VALUES (?, ?, ?, ?)",
            (code.strip().upper(), name or code, int(is_stc), int(is_sp)),
        )
        return cur.lastrowid


# ---------------------------------------------------------------------------
# Allocations – the heart of the system
# ---------------------------------------------------------------------------

_ALLOC_SELECT = """
SELECT
    a.*,
    f.initials   AS faculty_initials,
    f.full_name  AS faculty_name,
    f.status     AS faculty_status,
    c.code       AS career_code,
    l.lab_code   AS lab_code,
    l.pcs        AS lab_pcs,
    ts.label     AS time_slot_label,
    dg.code      AS day_group_code
FROM allocations a
JOIN faculty    f  ON f.id  = a.faculty_id
JOIN careers    c  ON c.id  = a.career_id
JOIN labs       l  ON l.id  = a.lab_id
JOIN time_slots ts ON ts.id = a.time_slot_id
JOIN day_groups dg ON dg.id = a.day_group_id
"""


def _row_to_allocation(r) -> Allocation:
    d = dict(r)
    return Allocation(
        id=d["id"],
        faculty_id=d["faculty_id"],
        career_id=d["career_id"],
        lab_id=d["lab_id"],
        time_slot_id=d["time_slot_id"],
        day_group_id=d["day_group_id"],
        batch_code=d["batch_code"],
        course_title=d["course_title"],
        module_name=d["module_name"],
        students=d["students"] or 0,
        module_start_date=d["module_start_date"],
        actual_start_date=d["actual_start_date"],
        notes=d["notes"],
        is_admission_open=bool(d["is_admission_open"]),
        is_active=bool(d["is_active"]),
        faculty_initials=d["faculty_initials"],
        faculty_name=d["faculty_name"],
        faculty_status=d["faculty_status"],
        career_code=d["career_code"],
        lab_code=d["lab_code"],
        lab_pcs=d["lab_pcs"],
        time_slot_label=d["time_slot_label"],
        day_group_code=d["day_group_code"],
    )


def list_allocations(active_only: bool = True) -> List[Allocation]:
    with db_session() as conn:
        q = _ALLOC_SELECT
        if active_only:
            q += " WHERE a.is_active = 1"
        q += " ORDER BY l.lab_code, ts.sort_order, dg.sort_order"
        rows = conn.execute(q).fetchall()
    return [_row_to_allocation(r) for r in rows]


def get_allocation(aid: int) -> Optional[Allocation]:
    with db_session() as conn:
        r = conn.execute(_ALLOC_SELECT + " WHERE a.id = ?", (aid,)).fetchone()
    return _row_to_allocation(r) if r else None


def add_allocation(data: Dict[str, Any]) -> int:
    """
    data keys: faculty_id, career_id, lab_id, time_slot_id, day_group_id,
               batch_code, course_title, module_name, students,
               module_start_date, actual_start_date, notes, is_admission_open
    """
    with db_session() as conn:
        cur = conn.execute(
            """
            INSERT INTO allocations (
                faculty_id, career_id, lab_id, time_slot_id, day_group_id,
                batch_code, course_title, module_name, students,
                module_start_date, actual_start_date, notes, is_admission_open
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                data["faculty_id"],
                data["career_id"],
                data["lab_id"],
                data["time_slot_id"],
                data["day_group_id"],
                data.get("batch_code"),
                data.get("course_title"),
                data.get("module_name"),
                data.get("students", 0),
                data.get("module_start_date"),
                data.get("actual_start_date"),
                data.get("notes"),
                int(data.get("is_admission_open", False)),
            ),
        )
        return cur.lastrowid


def update_allocation(aid: int, data: Dict[str, Any]) -> None:
    with db_session() as conn:
        conn.execute(
            """
            UPDATE allocations SET
                faculty_id = ?, career_id = ?, lab_id = ?,
                time_slot_id = ?, day_group_id = ?,
                batch_code = ?, course_title = ?, module_name = ?,
                students = ?, module_start_date = ?, actual_start_date = ?,
                notes = ?, is_admission_open = ?,
                updated_at = datetime('now')
            WHERE id = ?
            """,
            (
                data["faculty_id"],
                data["career_id"],
                data["lab_id"],
                data["time_slot_id"],
                data["day_group_id"],
                data.get("batch_code"),
                data.get("course_title"),
                data.get("module_name"),
                data.get("students", 0),
                data.get("module_start_date"),
                data.get("actual_start_date"),
                data.get("notes"),
                int(data.get("is_admission_open", False)),
                aid,
            ),
        )


def deactivate_allocation(aid: int) -> None:
    with db_session() as conn:
        conn.execute(
            "UPDATE allocations SET is_active = 0, updated_at = datetime('now') WHERE id = ?",
            (aid,),
        )


def delete_allocation(aid: int) -> None:
    with db_session() as conn:
        conn.execute("DELETE FROM allocations WHERE id = ?", (aid,))


def clear_all_allocations() -> int:
    """Permanently delete every allocation row. Returns how many were removed.
    Lookups (labs, faculty, careers, slots, day groups) are left untouched.
    """
    with db_session() as conn:
        cur = conn.execute("SELECT COUNT(*) FROM allocations")
        n = cur.fetchone()[0]
        conn.execute("DELETE FROM allocations")
        return n


# ---------------------------------------------------------------------------
# Conflict / capacity checks
# ---------------------------------------------------------------------------

def find_conflicts(
    lab_id: int,
    time_slot_id: int,
    day_group_id: int,
    exclude_id: Optional[int] = None,
) -> List[Allocation]:
    """Return existing active allocations that occupy the same lab + time + day group."""
    with db_session() as conn:
        q = _ALLOC_SELECT + """
            WHERE a.is_active = 1
              AND a.lab_id = ?
              AND a.time_slot_id = ?
              AND a.day_group_id = ?
        """
        params: list = [lab_id, time_slot_id, day_group_id]
        if exclude_id:
            q += " AND a.id != ?"
            params.append(exclude_id)
        rows = conn.execute(q, params).fetchall()
    return [_row_to_allocation(r) for r in rows]


def find_faculty_conflicts(
    faculty_id: int,
    time_slot_id: int,
    day_group_id: int,
    exclude_id: Optional[int] = None,
) -> List[Allocation]:
    """Return active allocations where the same faculty is already booked in this time + day."""
    with db_session() as conn:
        q = _ALLOC_SELECT + """
            WHERE a.is_active = 1
              AND a.faculty_id = ?
              AND a.time_slot_id = ?
              AND a.day_group_id = ?
        """
        params: list = [faculty_id, time_slot_id, day_group_id]
        if exclude_id:
            q += " AND a.id != ?"
            params.append(exclude_id)
        rows = conn.execute(q, params).fetchall()
    return [_row_to_allocation(r) for r in rows]


def derive_course_from_batch(batch_code: Optional[str]) -> Optional[str]:
    """Extract series prefix from batch code, e.g. AI-202412C2 → AI, PMTZ-202405B → PMTZ."""
    if not batch_code:
        return None
    batch_code = batch_code.strip()
    if "-" in batch_code:
        return batch_code.split("-", 1)[0].upper()
    return batch_code.upper() if batch_code else None


def check_capacity(lab_id: int, students: int) -> Tuple[bool, int]:
    """Returns (ok, available_pcs)."""
    with db_session() as conn:
        row = conn.execute("SELECT pcs FROM labs WHERE id = ?", (lab_id,)).fetchone()
    if not row:
        return False, 0
    pcs = row["pcs"]
    return students <= pcs, pcs
