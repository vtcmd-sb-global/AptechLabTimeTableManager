#!/usr/bin/env python3
"""
Aptech Lab Timetable – Database layer
Single source of truth (SQLite).
"""

from __future__ import annotations

import sqlite3
from pathlib import Path
from contextlib import contextmanager
from typing import Iterator

# Use /tmp in restricted environments; change to local data/ on a normal Windows machine
DB_PATH = Path("/tmp") / "aptech_timetable.db"
# For production / Windows packaging you can switch back to:
# DB_PATH = Path(__file__).parent / "data" / "timetable.db"


def get_connection() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH), timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
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
-- Labs
CREATE TABLE IF NOT EXISTS labs (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    lab_code    TEXT NOT NULL UNIQUE,          -- L1, L2, …
    pcs         INTEGER NOT NULL DEFAULT 12,
    notes       TEXT,
    active      INTEGER NOT NULL DEFAULT 1
);

-- Faculty
CREATE TABLE IF NOT EXISTS faculty (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    initials    TEXT NOT NULL UNIQUE,          -- MS, AA, HM, …
    full_name   TEXT NOT NULL,
    status      TEXT NOT NULL DEFAULT 'P',     -- P = Permanent, V = Visiting
    active      INTEGER NOT NULL DEFAULT 1
);

-- Careers / Programs
CREATE TABLE IF NOT EXISTS careers (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    code        TEXT NOT NULL UNIQUE,          -- CPISM, DISM, HDSE I, …
    name        TEXT,
    is_stc      INTEGER NOT NULL DEFAULT 0,    -- Short Term Course
    is_sp       INTEGER NOT NULL DEFAULT 0,    -- Special Program
    sort_order  INTEGER NOT NULL DEFAULT 100
);

-- Time slots
CREATE TABLE IF NOT EXISTS time_slots (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    code        TEXT NOT NULL UNIQUE,          -- B, C, D, E, F, G
    label       TEXT NOT NULL,                 -- B (9:00-11:00)
    start_time  TEXT,                          -- 09:00
    end_time    TEXT,                          -- 11:00
    sort_order  INTEGER NOT NULL DEFAULT 0
);

-- Day groups
CREATE TABLE IF NOT EXISTS day_groups (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    code        TEXT NOT NULL UNIQUE,          -- MWF, TTS, Sunday
    label       TEXT NOT NULL,                 -- M/W/F, T/T/S
    sort_order  INTEGER NOT NULL DEFAULT 0
);

-- The only transactional table – single source of truth
CREATE TABLE IF NOT EXISTS allocations (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    faculty_id          INTEGER NOT NULL REFERENCES faculty(id),
    career_id           INTEGER NOT NULL REFERENCES careers(id),
    lab_id              INTEGER NOT NULL REFERENCES labs(id),
    time_slot_id        INTEGER NOT NULL REFERENCES time_slots(id),
    day_group_id        INTEGER NOT NULL REFERENCES day_groups(id),
    batch_code          TEXT,                  -- AI-202412C2
    course_title        TEXT,                  -- AI, PMTZ, CDMA …
    module_name         TEXT,                  -- FSA, AZURE, PHP …
    students            INTEGER NOT NULL DEFAULT 0,
    module_start_date   TEXT,                  -- YYYY-MM-DD
    actual_start_date   TEXT,
    notes               TEXT,
    is_admission_open   INTEGER NOT NULL DEFAULT 0,
    is_active           INTEGER NOT NULL DEFAULT 1,
    created_at          TEXT DEFAULT (datetime('now')),
    updated_at          TEXT DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_alloc_faculty   ON allocations(faculty_id);
CREATE INDEX IF NOT EXISTS idx_alloc_career    ON allocations(career_id);
CREATE INDEX IF NOT EXISTS idx_alloc_lab       ON allocations(lab_id);
CREATE INDEX IF NOT EXISTS idx_alloc_timeslot  ON allocations(time_slot_id);
CREATE INDEX IF NOT EXISTS idx_alloc_daygroup  ON allocations(day_group_id);
CREATE INDEX IF NOT EXISTS idx_alloc_active    ON allocations(is_active);

-- Monthly snapshot meta (optional – for STC/SP beginning/end of month)
CREATE TABLE IF NOT EXISTS monthly_meta (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    year_month          TEXT NOT NULL UNIQUE,  -- 2026-08
    stc_begin           INTEGER DEFAULT 0,
    stc_end             INTEGER DEFAULT 0,
    sp_begin            INTEGER DEFAULT 0,
    sp_end              INTEGER DEFAULT 0,
    notes               TEXT
);
"""


def init_db() -> None:
    with db_session() as conn:
        conn.executescript(SCHEMA)
    _seed_lookups()


def _seed_lookups() -> None:
    """Insert standard lookup values if tables are empty."""
    with db_session() as conn:
        # Labs
        if conn.execute("SELECT COUNT(*) FROM labs").fetchone()[0] == 0:
            labs = [
                ("L1", 12), ("L2", 14), ("L3", 14), ("L4", 11),
                ("L5", 17), ("L6", 8), ("L7", 12), ("L8", 14),
            ]
            conn.executemany(
                "INSERT INTO labs (lab_code, pcs) VALUES (?, ?)", labs
            )

        # Time slots
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

        # Day groups
        if conn.execute("SELECT COUNT(*) FROM day_groups").fetchone()[0] == 0:
            groups = [
                ("MWF", "M/W/F", 1),
                ("TTS", "T/T/S", 2),
                ("Sunday", "Sunday", 3),
            ]
            conn.executemany(
                "INSERT INTO day_groups (code, label, sort_order) VALUES (?, ?, ?)",
                groups,
            )

        # Careers
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

        # Faculty (from the August-26 sheet)
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


if __name__ == "__main__":
    init_db()
    print(f"Database initialized at: {DB_PATH}")
