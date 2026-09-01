#!/usr/bin/env python3
"""
Aptech Monthly BSR - Database layer
"""

from __future__ import annotations

import sqlite3
from pathlib import Path
from contextlib import contextmanager
from typing import Iterator

# ---------------------------------------------------------------------------
# Persistent database path - always under the application folder
# ---------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parent
DB_DIR = ROOT / "database"
DB_PATH = DB_DIR / "monthly_bsr.db"

# Optional legacy paths (used only to migrate data once if present)
_LEGACY_PATHS = [
    DB_DIR / "lab_status.db",          # previous name before rename to monthly_bsr.db
    Path("/tmp") / "aptech_timetable.db",
    ROOT / "data" / "timetable.db",
]


def get_connection() -> sqlite3.Connection:
    DB_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH), timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    # DELETE journal so the single .db file is easy to copy/backup
    conn.execute("PRAGMA journal_mode = DELETE")
    return conn


@contextmanager
def db_session() -> Iterator[sqlite3.Connection]:
    conn = get_connection()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


SCHEMA = """
CREATE TABLE IF NOT EXISTS labs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lab_code TEXT NOT NULL UNIQUE,
    pcs INTEGER NOT NULL DEFAULT 12,
    notes TEXT,
    active INTEGER NOT NULL DEFAULT 1
);
CREATE TABLE IF NOT EXISTS faculty (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    initials TEXT NOT NULL UNIQUE,
    full_name TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'P',
    active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS careers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL UNIQUE,
    name TEXT,
    is_stc INTEGER NOT NULL DEFAULT 0,
    is_sp INTEGER NOT NULL DEFAULT 0,
    sort_order INTEGER NOT NULL DEFAULT 100
);
CREATE TABLE IF NOT EXISTS modules (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    category TEXT NOT NULL DEFAULT 'CAREER',
    sort_order INTEGER NOT NULL DEFAULT 100,
    active INTEGER NOT NULL DEFAULT 1
);
CREATE TABLE IF NOT EXISTS time_slots (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL UNIQUE,
    label TEXT NOT NULL,
    start_time TEXT,
    end_time TEXT,
    sort_order INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS day_groups (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL UNIQUE,
    label TEXT NOT NULL,
    sort_order INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS monthly_bsr_metrics (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    year_month TEXT NOT NULL UNIQUE,
    stc_beg INTEGER NOT NULL DEFAULT 0,
    stc_end INTEGER NOT NULL DEFAULT 0,
    sp_beg INTEGER NOT NULL DEFAULT 0,
    sp_end INTEGER NOT NULL DEFAULT 0,
    total_batches INTEGER NOT NULL DEFAULT 0,
    total_students INTEGER NOT NULL DEFAULT 0,
    mwf_batches INTEGER NOT NULL DEFAULT 0,
    mwf_students INTEGER NOT NULL DEFAULT 0,
    tts_batches INTEGER NOT NULL DEFAULT 0,
    tts_students INTEGER NOT NULL DEFAULT 0,
    report_path TEXT,
    created_at TEXT DEFAULT (datetime('now')),
    notes TEXT
);
CREATE TABLE IF NOT EXISTS allocations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    faculty_id INTEGER REFERENCES faculty(id),
    career_id INTEGER REFERENCES careers(id),
    lab_id INTEGER NOT NULL REFERENCES labs(id),
    time_slot_id INTEGER NOT NULL REFERENCES time_slots(id),
    day_group_id INTEGER NOT NULL REFERENCES day_groups(id),
    batch_code TEXT,
    course_title TEXT,
    module_name TEXT,
    students INTEGER NOT NULL DEFAULT 0,
    module_start_date TEXT,
    actual_start_date TEXT,
    notes TEXT,
    is_admission_open INTEGER NOT NULL DEFAULT 0,
    is_active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS idx_alloc_faculty ON allocations(faculty_id);
CREATE INDEX IF NOT EXISTS idx_alloc_career ON allocations(career_id);
CREATE INDEX IF NOT EXISTS idx_alloc_lab ON allocations(lab_id);
CREATE INDEX IF NOT EXISTS idx_alloc_timeslot ON allocations(time_slot_id);
CREATE INDEX IF NOT EXISTS idx_alloc_daygroup ON allocations(day_group_id);
CREATE INDEX IF NOT EXISTS idx_alloc_active ON allocations(is_active);
"""



def _migrate_from_legacy() -> None:
    """If the new DB is empty and a legacy DB exists, copy it once."""
    if DB_PATH.exists() and DB_PATH.stat().st_size > 0:
        return
    for legacy in _LEGACY_PATHS:
        if legacy.exists() and legacy.stat().st_size > 0:
            import shutil
            DB_DIR.mkdir(parents=True, exist_ok=True)
            shutil.copy2(legacy, DB_PATH)
            print(f"Migrated existing database from {legacy} → {DB_PATH}")
            return


def _ensure_columns(conn: sqlite3.Connection) -> None:
    """Add columns introduced after the first release (safe on existing DBs)."""
    from datetime import datetime
    fac_cols = {r[1] for r in conn.execute("PRAGMA table_info(faculty)").fetchall()}
    if "created_at" not in fac_cols:
        conn.execute(
            "ALTER TABLE faculty ADD COLUMN created_at TEXT DEFAULT (datetime('now'))"
        )
    alloc_cols = {r[1] for r in conn.execute("PRAGMA table_info(allocations)").fetchall()}
    if "report_month" not in alloc_cols:
        conn.execute("ALTER TABLE allocations ADD COLUMN report_month TEXT")
        cur = datetime.now().strftime("%Y-%m")
        conn.execute(
            "UPDATE allocations SET report_month = ? WHERE report_month IS NULL OR report_month = ''",
            (cur,),
        )
    # Allow NULL faculty_id / career_id so Admission Open = Yes rows can omit them
    _migrate_alloc_nullable_fk(conn)


def _migrate_alloc_nullable_fk(conn: sqlite3.Connection) -> None:
    """Rebuild allocations so faculty_id and career_id may be NULL (Admission Open)."""
    # Detect if already nullable by checking a dummy insert is unnecessary –
    # SQLite has no simple NOT NULL introspection on existing tables after CREATE.
    # Use a one-time flag table.
    conn.execute(
        "CREATE TABLE IF NOT EXISTS _schema_flags (flag TEXT PRIMARY KEY)"
    )
    row = conn.execute(
        "SELECT 1 FROM _schema_flags WHERE flag = 'alloc_nullable_fk'"
    ).fetchone()
    if row:
        return

    conn.executescript(
        """
        CREATE TABLE allocations_new (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            faculty_id INTEGER REFERENCES faculty(id),
            career_id INTEGER REFERENCES careers(id),
            lab_id INTEGER NOT NULL REFERENCES labs(id),
            time_slot_id INTEGER NOT NULL REFERENCES time_slots(id),
            day_group_id INTEGER NOT NULL REFERENCES day_groups(id),
            batch_code TEXT,
            course_title TEXT,
            module_name TEXT,
            students INTEGER NOT NULL DEFAULT 0,
            module_start_date TEXT,
            actual_start_date TEXT,
            notes TEXT,
            is_admission_open INTEGER NOT NULL DEFAULT 0,
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now')),
            report_month TEXT
        );
        INSERT INTO allocations_new (
            id, faculty_id, career_id, lab_id, time_slot_id, day_group_id,
            batch_code, course_title, module_name, students,
            module_start_date, actual_start_date, notes,
            is_admission_open, is_active, created_at, updated_at, report_month
        )
        SELECT
            id, faculty_id, career_id, lab_id, time_slot_id, day_group_id,
            batch_code, course_title, module_name, students,
            module_start_date, actual_start_date, notes,
            is_admission_open, is_active, created_at, updated_at, report_month
        FROM allocations;
        DROP TABLE allocations;
        ALTER TABLE allocations_new RENAME TO allocations;
        CREATE INDEX IF NOT EXISTS idx_alloc_faculty ON allocations(faculty_id);
        CREATE INDEX IF NOT EXISTS idx_alloc_career ON allocations(career_id);
        CREATE INDEX IF NOT EXISTS idx_alloc_lab ON allocations(lab_id);
        CREATE INDEX IF NOT EXISTS idx_alloc_timeslot ON allocations(time_slot_id);
        CREATE INDEX IF NOT EXISTS idx_alloc_daygroup ON allocations(day_group_id);
        CREATE INDEX IF NOT EXISTS idx_alloc_active ON allocations(is_active);
        """
    )
    conn.execute(
        "INSERT INTO _schema_flags (flag) VALUES ('alloc_nullable_fk')"
    )


def init_db() -> None:
    """Create schema if needed. Never deletes existing data."""
    _migrate_from_legacy()
    with db_session() as conn:
        conn.executescript(SCHEMA)
        _ensure_columns(conn)
    _seed_lookups()
    _migrate_day_groups()


def _seed_lookups() -> None:
    """Insert standard lookup values if tables are empty (first run only)."""
    with db_session() as conn:
        if conn.execute("SELECT COUNT(*) FROM labs").fetchone()[0] == 0:
            labs = [
                ("L1", 12), ("L2", 14), ("L3", 14), ("L4", 11),
                ("L5", 17), ("L6", 8), ("L7", 12), ("L8", 14),
            ]
            conn.executemany(
                "INSERT INTO labs (lab_code, pcs) VALUES (?, ?)", labs
            )

        if conn.execute("SELECT COUNT(*) FROM time_slots").fetchone()[0] == 0:
            slots = [
                ("B", "B (9:00-11:00)", "09:00", "11:00", 1),
                ("C", "C (11:00-1:00)", "11:00", "13:00", 2),
                ("D", "D (1:00-3:00)", "13:00", "15:00", 3),
                ("E", "E (3:00-5:00)", "15:00", "17:00", 4),
                ("F", "F (5:00-7:00)", "17:00", "19:00", 5),
                ("G", "G (7:00-9:00)", "19:00", "21:00", 6),
            ]
            conn.executemany(
                "INSERT INTO time_slots (code, label, start_time, end_time, sort_order) "
                "VALUES (?, ?, ?, ?, ?)",
                slots,
            )

        if conn.execute("SELECT COUNT(*) FROM day_groups").fetchone()[0] == 0:
            groups = [
                ("MWF", "M/W/F", 1),
                ("TTS", "T/T/S", 2),
                ("Saturday", "Saturday", 3),
            ]
            conn.executemany(
                "INSERT INTO day_groups (code, label, sort_order) VALUES (?, ?, ?)",
                groups,
            )

        if conn.execute("SELECT COUNT(*) FROM careers").fetchone()[0] == 0:
            careers = [
                ("CPISM", "CPISM", 0, 0, 10),
                ("DISM", "DISM", 0, 0, 20),
                ("HDSE I", "HDSE I", 0, 0, 30),
                ("HDSE II", "HDSE II", 0, 0, 40),
                ("ADSE I", "ADSE I", 0, 0, 50),
                ("ADSE II", "ADSE II", 0, 0, 60),
                ("CDMA", "CDMA", 0, 0, 70),
                ("AID", "AID", 0, 0, 80),
                ("CS", "CS", 0, 0, 90),
                ("STC", "Short Term Course", 1, 0, 100),
                ("SP", "Special Program", 0, 1, 110),
                ("SEM I", "Semester I", 0, 0, 120),
                ("OST", "OST", 0, 0, 130),
            ]
            conn.executemany(
                "INSERT INTO careers (code, name, is_stc, is_sp, sort_order) "
                "VALUES (?, ?, ?, ?, ?)",
                careers,
            )


                # Modules catalog – insert any missing defaults (idempotent)
        if conn.execute(
            "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='modules'"
        ).fetchone()[0]:
            # Migrate old SUBJECT category → CAREER
            conn.execute("UPDATE modules SET category = 'CAREER' WHERE UPPER(category) = 'SUBJECT'")
            modules = [
                # Common program / career modules
                ("FSA", "CAREER", 10),
                ("MUI", "CAREER", 20),
                ("PWD", "CAREER", 30),
                ("FBDS", "CAREER", 40),
                ("AZURE", "CAREER", 50),
                ("XML/JS", "CAREER", 60),
                ("SQL", "CAREER", 70),
                ("E-Mail", "CAREER", 80),
                ("GIT", "CAREER", 90),
                ("SEO", "CAREER", 100),
                # Short courses / STC topics
                ("MS Office", "STC", 200),
                ("Web Designing", "STC", 210),
                ("ASP.Net/C#", "STC", 220),
                ("Java", "STC", 230),
                ("Flutter and Dart", "STC", 240),
                ("C Lang.", "STC", 250),
                ("PHP", "STC", 260),
                ("Data Science", "STC", 270),
                ("Mern Stack", "STC", 280),
                ("Data Analytics", "STC", 290),
                ("Wordpress", "STC", 300),
                ("Advance Excel", "STC", 310),
                ("Shopify", "STC", 320),
            ]
            existing = {
                (r[0] or "").strip().lower()
                for r in conn.execute("SELECT name FROM modules").fetchall()
            }
            for name, cat, order in modules:
                if name.lower() not in existing:
                    conn.execute(
                        "INSERT INTO modules (name, category, sort_order) VALUES (?, ?, ?)",
                        (name, cat, order),
                    )

        # Remove short-course names that were incorrectly seeded as careers
        bad_codes = (
            "MSOFFICE", "WEBDES", "ASPNET", "JAVA", "FLUTTER", "CLANG",
            "PHP", "DATASCI", "MERN", "DATAAN", "WP", "EXCEL", "SHOPIFY",
        )
        conn.execute(
            f"DELETE FROM careers WHERE code IN ({','.join('?'*len(bad_codes))}) "
            "AND id NOT IN (SELECT career_id FROM allocations)",
            bad_codes,
        )

        # Faculty - seed only on first run so user-added teachers are never wiped
        if conn.execute("SELECT COUNT(*) FROM faculty").fetchone()[0] == 0:
            faculty = [
                ("MS", "Shahbaz", "V"),
                ("AA", "Aman Ali", "V"),
                ("HM", "Hafiz Moiz", "V"),
                ("T", "Talha Qureshi", "P"),
                ("A", "Aousaja", "P"),
                ("BF", "M Bilal Faroqui", "P"),
                ("RA", "Riaz Ahmed", "P"),
                ("MAR", "M Ashar Rehan", "P"),
                ("Z", "Syed Zaid", "V"),
                ("KA", "Kashif Ali", "P"),
                ("AW", "Abdul Wahab", "P"),
                ("H", "Hadiqa", "V"),
            ]
            conn.executemany(
                "INSERT INTO faculty (initials, full_name, status) VALUES (?, ?, ?)",
                faculty,
            )


def _migrate_day_groups() -> None:
    """Replace Sunday (off day) with Saturday for existing databases."""
    with db_session() as conn:
        has_sun = conn.execute(
            "SELECT id FROM day_groups WHERE code = 'Sunday' OR code = 'SUNDAY'"
        ).fetchone()
        has_sat = conn.execute(
            "SELECT id FROM day_groups WHERE code = 'Saturday' OR code = 'SATURDAY'"
        ).fetchone()
        if has_sun and not has_sat:
            # Rename Sunday → Saturday (keeps allocation FKs intact)
            conn.execute(
                "UPDATE day_groups SET code = 'Saturday', label = 'Saturday' WHERE id = ?",
                (has_sun[0],),
            )
        elif has_sun and has_sat:
            # Move any allocations from Sunday to Saturday, then drop Sunday
            conn.execute(
                "UPDATE allocations SET day_group_id = ? WHERE day_group_id = ?",
                (has_sat[0], has_sun[0]),
            )
            conn.execute("DELETE FROM day_groups WHERE id = ?", (has_sun[0],))
        elif not has_sat and not has_sun:
            # Ensure Saturday exists if table was partially seeded
            n = conn.execute("SELECT COUNT(*) FROM day_groups").fetchone()[0]
            if n > 0:
                conn.execute(
                    "INSERT INTO day_groups (code, label, sort_order) VALUES ('Saturday', 'Saturday', 3)"
                )



def get_db_path() -> Path:
    """Public helper so the UI can show where the database lives."""
    return DB_PATH


if __name__ == "__main__":
    init_db()
    print(f"Database ready at: {DB_PATH}")
    print(f"  exists = {DB_PATH.exists()}  size = {DB_PATH.stat().st_size if DB_PATH.exists() else 0} bytes")
