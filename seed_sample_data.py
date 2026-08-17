#!/usr/bin/env python3
"""
Seed a few realistic allocations taken from the August-26 sheet
so the application is immediately usable for testing.
"""

from database import init_db, db_session
from models import list_faculty, list_labs, list_careers, list_time_slots, list_day_groups, add_allocation


def seed():
    init_db()

    with db_session() as conn:
        count = conn.execute("SELECT COUNT(*) FROM allocations").fetchone()[0]
        if count > 0:
            print(f"Allocations already exist ({count}). Skipping seed.")
            return

    faculty = {f.initials: f.id for f in list_faculty()}
    labs = {l.lab_code: l.id for l in list_labs()}
    careers = {c.code: c.id for c in list_careers()}
    slots = {s.code: s.id for s in list_time_slots()}
    days = {d.code: d.id for d in list_day_groups()}

    samples = [
        # lab, day, slot, faculty, career, course, batch, module, students, start
        ("L1", "MWF", "B", "MS", "HDSE II", "AI", "AI-202412C2", "FSA", 4, "2024-09-10"),
        ("L1", "TTS", "B", "T", "ADSE I", "PMTZ", "PMTZ-202405B", "FSA", 5, "2024-05-04"),
        ("L1", "TTS", "C", "AA", "ADSE II", "PMTZ", "PMTZ-202401B", "FBDS", 14, "2024-01-22"),
        ("L1", "TTS", "D", "AA", "CDMA", "CDMA", "DM-202601C", "E-Mail", 5, "2026-01-24"),
        ("L3", "MWF", "B", "T", "HDSE II", "AI", "AI-202501B", "AZURE", 6, "2025-01-13"),
        ("L3", "TTS", "B", "Z", "DISM", "AI", "AI-202601B", "XML/JS", 3, "2026-01-27"),
        ("L3", "MWF", "C", "T", "HDSE II", "AI", "AI-202409C", "FSA", 9, "2024-09-30"),
        ("L3", "TTS", "C", "T", "HDSE I", "AI", "AI-202505B", "MUI", 7, "2025-05-12"),
        ("L5", "MWF", "B", "AW", "CPISM", "AI", "AI-202607B", "PWD", 5, "2026-07-13"),
        ("L5", "TTS", "C", "Z", "HDSE I", "AI", "AI-202506C", "SQL", 7, "2025-07-22"),
        ("L5", "TTS", "D", "T", "CPISM", "AI", "AI-202606D", "PWD", 0, "2026-06-04"),
        ("L2", "TTS", "B", "MS", "ADSE I", "PMTZ", "PMTZ-202412C", "MUI", 8, "2024-09-10"),
    ]

    for lab, day, slot, fac, career, course, batch, module, students, start in samples:
        if fac not in faculty or lab not in labs or career not in careers:
            print(f"Skipping {batch} – missing lookup")
            continue
        add_allocation({
            "faculty_id": faculty[fac],
            "career_id": careers[career],
            "lab_id": labs[lab],
            "time_slot_id": slots[slot],
            "day_group_id": days[day],
            "batch_code": batch,
            "course_title": course,
            "module_name": module,
            "students": students,
            "module_start_date": start,
            "is_admission_open": students == 0,
        })
        print(f"  + {batch}  ({fac} / {lab} / {day} {slot})")

    print("Sample data seeded successfully.")


if __name__ == "__main__":
    seed()
