#!/usr/bin/env python3
"""
CRUD helpers and domain objects for Aptech BSR.
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
    created_at: Optional[str] = None


@dataclass
class Module:
    id: int
    name: str
    category: str = "CAREER"  # SUBJECT | STC | SP
    sort_order: int = 100
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
    faculty_id: Optional[int]
    career_id: Optional[int]
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
    report_month: str = ""


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
    return [
        Faculty(
            id=r["id"], initials=r["initials"], full_name=r["full_name"],
            status=r["status"], active=bool(r["active"]),
            created_at=r["created_at"] if "created_at" in r.keys() else None,
        )
        for r in rows
    ]



def list_modules(active_only: bool = True, category: str | None = None) -> List[Module]:
    with db_session() as conn:
        sql = "SELECT * FROM modules WHERE 1=1"
        params: list = []
        if active_only:
            sql += " AND active = 1"
        if category:
            sql += " AND category = ?"
            params.append(category)
        sql += " ORDER BY sort_order, name"
        rows = conn.execute(sql, params).fetchall()
    return [
        Module(
            id=r["id"], name=r["name"], category=r["category"],
            sort_order=r["sort_order"], active=bool(r["active"]),
        )
        for r in rows
    ]


def add_module(name: str, category: str = "CAREER") -> int:
    with db_session() as conn:
        cur = conn.execute(
            "INSERT INTO modules (name, category) VALUES (?, ?)",
            (name.strip(), category.strip().upper() or "CAREER"),
        )
        return cur.lastrowid


def update_module(module_id: int, name: str, category: str = "CAREER") -> None:
    with db_session() as conn:
        conn.execute(
            "UPDATE modules SET name = ?, category = ? WHERE id = ?",
            (name.strip(), category.strip().upper() or "CAREER", module_id),
        )


def delete_module(module_id: int) -> None:
    with db_session() as conn:
        conn.execute("DELETE FROM modules WHERE id = ?", (module_id,))



def get_monthly_bsr_metrics(year_month: str) -> Optional[dict]:
    with db_session() as conn:
        row = conn.execute(
            "SELECT * FROM monthly_bsr_metrics WHERE year_month = ?",
            (year_month,),
        ).fetchone()
    return dict(row) if row else None


def list_monthly_bsr_metrics() -> List[dict]:
    with db_session() as conn:
        rows = conn.execute(
            "SELECT * FROM monthly_bsr_metrics ORDER BY year_month DESC"
        ).fetchall()
    return [dict(r) for r in rows]


def save_monthly_bsr_metrics(year_month: str, data: dict) -> None:
    """Insert or update metrics snapshot for a month (YYYY-MM)."""
    with db_session() as conn:
        existing = conn.execute(
            "SELECT id FROM monthly_bsr_metrics WHERE year_month = ?",
            (year_month,),
        ).fetchone()
        fields = (
            "stc_beg", "stc_end", "sp_beg", "sp_end",
            "total_batches", "total_students",
            "mwf_batches", "mwf_students", "tts_batches", "tts_students",
            "report_path", "notes",
        )
        vals = [data.get(f, 0) for f in fields]
        if existing:
            sets = ", ".join(f"{f} = ?" for f in fields)
            conn.execute(
                f"UPDATE monthly_bsr_metrics SET {sets} WHERE year_month = ?",
                (*vals, year_month),
            )
        else:
            cols = ", ".join(["year_month"] + list(fields))
            placeholders = ", ".join(["?"] * (1 + len(fields)))
            conn.execute(
                f"INSERT INTO monthly_bsr_metrics ({cols}) VALUES ({placeholders})",
                (year_month, *vals),
            )


def previous_year_month(year_month: str) -> str:
    y, m = map(int, year_month.split("-"))
    if m == 1:
        return f"{y - 1}-12"
    return f"{y}-{m - 1:02d}"


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


def get_faculty(fid: int) -> Optional[Faculty]:
    with db_session() as conn:
        r = conn.execute("SELECT * FROM faculty WHERE id = ?", (fid,)).fetchone()
    if not r:
        return None
    d = dict(r)
    return Faculty(
        id=d["id"],
        initials=d["initials"],
        full_name=d["full_name"],
        status=d["status"],
        active=bool(d.get("active", 1)),
        created_at=d.get("created_at"),
    )


def set_faculty_active(fid: int, active: bool) -> None:
    """Soft-activate / soft-deactivate a faculty member. Allocations are never deleted."""
    with db_session() as conn:
        conn.execute(
            "UPDATE faculty SET active = ? WHERE id = ?",
            (1 if active else 0, fid),
        )


def count_allocations_for_faculty(fid: int, active_only: bool = False) -> int:
    with db_session() as conn:
        q = "SELECT COUNT(*) FROM allocations WHERE faculty_id = ?"
        params: list = [fid]
        if active_only:
            q += " AND is_active = 1"
        return conn.execute(q, params).fetchone()[0]


def delete_faculty(fid: int) -> tuple:
    """
    Hard-delete only if the faculty has ZERO allocation history.
    Returns (ok: bool, message: str). Prefer set_faculty_active(False) otherwise.
    """
    n = count_allocations_for_faculty(fid, active_only=False)
    if n > 0:
        return False, (
            f"This faculty has {n} allocation record(s). "
            "Deactivate instead of deleting to preserve history."
        )
    with db_session() as conn:
        conn.execute("DELETE FROM faculty WHERE id = ?", (fid,))
    return True, "Faculty permanently deleted."




def add_lab(lab_code: str, pcs: int, notes: str = "") -> int:
    with db_session() as conn:
        cur = conn.execute(
            "INSERT INTO labs (lab_code, pcs, notes) VALUES (?, ?, ?)",
            (lab_code.strip().upper(), pcs, notes),
        )
        return cur.lastrowid



def update_career(career_id: int, code: str, name: str = "", is_stc: bool = False, is_sp: bool = False) -> None:
    with db_session() as conn:
        conn.execute(
            "UPDATE careers SET code = ?, name = ?, is_stc = ?, is_sp = ? WHERE id = ?",
            (code.strip().upper(), name or code, int(is_stc), int(is_sp), career_id),
        )


def count_allocations_for_career(career_id: int) -> int:
    with db_session() as conn:
        return conn.execute(
            "SELECT COUNT(*) FROM allocations WHERE career_id = ? AND is_active = 1",
            (career_id,),
        ).fetchone()[0]


def delete_career(career_id: int) -> None:
    """Delete only if no active allocations reference it."""
    if count_allocations_for_career(career_id) > 0:
        raise ValueError("Cannot delete career that has active allocations. Reassign or clear those first.")
    with db_session() as conn:
        conn.execute("DELETE FROM careers WHERE id = ?", (career_id,))


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
LEFT JOIN faculty    f  ON f.id  = a.faculty_id
LEFT JOIN careers    c  ON c.id  = a.career_id
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
        report_month=(d.get("report_month") or ""),
    )


def list_allocations(active_only: bool = True, report_month: Optional[str] = None) -> List[Allocation]:
    with db_session() as conn:
        q = _ALLOC_SELECT
        clauses = []
        params: list = []
        if active_only:
            clauses.append("a.is_active = 1")
        if report_month:
            clauses.append("a.report_month = ?")
            params.append(report_month)
        if clauses:
            q += " WHERE " + " AND ".join(clauses)
        q += " ORDER BY l.lab_code, ts.sort_order, dg.sort_order"
        rows = conn.execute(q, params).fetchall()
    return [_row_to_allocation(r) for r in rows]


def list_available_report_months() -> List[str]:
    """Months with allocation data first (newest first), then current month if empty."""
    with db_session() as conn:
        rows = conn.execute(
            """
            SELECT report_month AS ym, COUNT(*) AS n
            FROM allocations
            WHERE report_month IS NOT NULL AND report_month != ''
              AND is_active = 1
            GROUP BY report_month
            HAVING n > 0
            ORDER BY report_month DESC
            """
        ).fetchall()
    months = [r["ym"] for r in rows if r["ym"]]
    from datetime import datetime
    cur = datetime.now().strftime("%Y-%m")
    if cur not in months:
        months.append(cur)
    return months


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
                module_start_date, actual_start_date, notes, is_admission_open,
                report_month
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                data.get("report_month"),
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
                notes = ?, is_admission_open = ?, report_month = ?,
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
                data.get("report_month"),
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
