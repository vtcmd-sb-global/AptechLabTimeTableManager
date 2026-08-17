#!/usr/bin/env python3
"""
All automatic calculations and report aggregations.
Everything is derived from the allocations table – never stored separately.
"""

from __future__ import annotations

from collections import defaultdict
from typing import Any, Dict, List, Tuple
from models import Allocation, list_allocations, list_careers, list_faculty


def get_active_allocations() -> List[Allocation]:
    return list_allocations(active_only=True)


# ---------------------------------------------------------------------------
# Faculty Workload
# ---------------------------------------------------------------------------

def faculty_workload() -> List[Dict[str, Any]]:
    """
    Returns one row per faculty:
      initials, status, full_name, load (#batches), career_summary,
      stc_count, total_students, timings
    """
    allocs = get_active_allocations()
    by_fac: Dict[int, List[Allocation]] = defaultdict(list)
    for a in allocs:
        by_fac[a.faculty_id].append(a)

    faculty = {f.id: f for f in list_faculty(active_only=False)}
    result = []

    for fid, items in sorted(by_fac.items(), key=lambda x: faculty[x[0]].initials if x[0] in faculty else ""):
        fac = faculty.get(fid)
        if not fac:
            continue

        load = len(items)
        students = sum(a.students for a in items)
        stc = sum(1 for a in items if a.career_code == "STC")

        # Career summary like "4-26" or date strings seen in original
        career_parts = []
        for a in items:
            career_parts.append(a.career_code)
        career_summary = f"{load}-{students}" if load else "-"

        # Timings string
        timing_set = set()
        for a in items:
            # Simplify label e.g. "B (9:00-11:00)" → "9-11"
            label = a.time_slot_label
            if "(" in label and ")" in label:
                inner = label[label.find("(")+1 : label.find(")")]
                timing_set.add(f"{inner} {a.day_group_code}")
            else:
                timing_set.add(f"{label} {a.day_group_code}")
        timings = " / ".join(sorted(timing_set))

        result.append({
            "faculty_id": fid,
            "initials": fac.initials,
            "status": fac.status,
            "full_name": fac.full_name,
            "load": load,
            "career_summary": career_summary,
            "stc": stc,
            "students": students,
            "timings": timings,
            "allocations": items,
        })

    # Also include faculty with zero load (optional)
    seen = {r["faculty_id"] for r in result}
    for f in list_faculty(active_only=True):
        if f.id not in seen:
            result.append({
                "faculty_id": f.id,
                "initials": f.initials,
                "status": f.status,
                "full_name": f.full_name,
                "load": 0,
                "career_summary": "-",
                "stc": 0,
                "students": 0,
                "timings": "",
                "allocations": [],
            })

    result.sort(key=lambda r: r["initials"])
    return result


# ---------------------------------------------------------------------------
# Faculty × Career matrix
# ---------------------------------------------------------------------------

def faculty_career_matrix() -> Tuple[List[str], List[Dict[str, Any]]]:
    """
    Returns (career_codes_ordered, rows)
    Each row: faculty_name + counts per career + total
    """
    careers = [c.code for c in list_careers() if c.code not in ("STC", "SP")]
    # Prefer the classic order seen in the sheet
    preferred = ["CPISM", "DISM", "HDSE I", "HDSE II", "ADSE I", "ADSE II",
                 "CDMA", "AID", "CS"]
    ordered = [c for c in preferred if c in careers]
    ordered += [c for c in careers if c not in ordered]

    allocs = get_active_allocations()
    # faculty_id → career_code → count
    matrix: Dict[int, Dict[str, int]] = defaultdict(lambda: defaultdict(int))
    fac_names: Dict[int, str] = {}

    for a in allocs:
        if a.career_code in ("STC", "SP"):
            continue
        matrix[a.faculty_id][a.career_code] += 1
        fac_names[a.faculty_id] = a.faculty_name

    rows = []
    for fid, counts in sorted(matrix.items(), key=lambda x: fac_names.get(x[0], "")):
        row = {"faculty": fac_names.get(fid, "?"), "faculty_id": fid}
        total = 0
        for c in ordered:
            n = counts.get(c, 0)
            row[c] = n
            total += n
        row["Total"] = total
        rows.append(row)

    return ordered, rows


# ---------------------------------------------------------------------------
# Career / Semester summary
# ---------------------------------------------------------------------------

def career_summary() -> List[Dict[str, Any]]:
    """
    No. of Batches, No. of Students, % for each career.
    """
    allocs = get_active_allocations()
    by_career: Dict[str, List[Allocation]] = defaultdict(list)
    for a in allocs:
        by_career[a.career_code].append(a)

    total_students = sum(a.students for a in allocs) or 1
    total_batches = len(allocs)

    careers = list_careers()
    result = []
    for c in careers:
        items = by_career.get(c.code, [])
        batches = len(items)
        students = sum(a.students for a in items)
        pct = students / total_students if total_students else 0
        result.append({
            "code": c.code,
            "name": c.name or c.code,
            "batches": batches,
            "students": students,
            "pct": pct,
            "is_stc": c.is_stc,
            "is_sp": c.is_sp,
        })

    # Grand total
    result.append({
        "code": "Total",
        "name": "Total",
        "batches": total_batches,
        "students": sum(a.students for a in allocs),
        "pct": 1.0,
        "is_stc": False,
        "is_sp": False,
    })
    return result


# ---------------------------------------------------------------------------
# MWF / TTS totals
# ---------------------------------------------------------------------------

def day_group_totals() -> Dict[str, Dict[str, int]]:
    """
    {
      "MWF": {"batches": 20, "students": 200},
      "TTS": {"batches": 26, "students": 165},
      ...
    }
    """
    allocs = get_active_allocations()
    totals: Dict[str, Dict[str, int]] = defaultdict(lambda: {"batches": 0, "students": 0})
    for a in allocs:
        totals[a.day_group_code]["batches"] += 1
        totals[a.day_group_code]["students"] += a.students
    return dict(totals)


# ---------------------------------------------------------------------------
# Visiting Faculty list
# ---------------------------------------------------------------------------

def visiting_faculty() -> List[Dict[str, Any]]:
    wl = faculty_workload()
    return [
        {
            "name": r["full_name"],
            "initials": r["initials"],
            "visiting_no": r["load"],
            "time": r["timings"],
        }
        for r in wl if r["status"].upper() == "V" and r["load"] > 0
    ]


# ---------------------------------------------------------------------------
# STC / SP counts
# ---------------------------------------------------------------------------

def stc_sp_counts() -> Dict[str, int]:
    allocs = get_active_allocations()
    stc_batches = sum(1 for a in allocs if a.career_code == "STC")
    stc_students = sum(a.students for a in allocs if a.career_code == "STC")
    sp_batches = sum(1 for a in allocs if a.career_code == "SP")
    sp_students = sum(a.students for a in allocs if a.career_code == "SP")
    return {
        "stc_batches": stc_batches,
        "stc_students": stc_students,
        "sp_batches": sp_batches,
        "sp_students": sp_students,
    }


# ---------------------------------------------------------------------------
# Lab utilization
# ---------------------------------------------------------------------------

def lab_utilization() -> List[Dict[str, Any]]:
    allocs = get_active_allocations()
    by_lab: Dict[str, List[Allocation]] = defaultdict(list)
    for a in allocs:
        by_lab[a.lab_code].append(a)

    result = []
    for lab_code, items in sorted(by_lab.items()):
        pcs = items[0].lab_pcs if items else 0
        total_students = sum(a.students for a in items)
        batches = len(items)
        result.append({
            "lab_code": lab_code,
            "pcs": pcs,
            "batches": batches,
            "total_students": total_students,
            "utilization_pct": (total_students / (pcs * batches)) if pcs and batches else 0,
        })
    return result


# ---------------------------------------------------------------------------
# Full dashboard snapshot (for GUI + Excel)
# ---------------------------------------------------------------------------

def full_dashboard() -> Dict[str, Any]:
    return {
        "faculty_workload": faculty_workload(),
        "faculty_career_matrix": faculty_career_matrix(),
        "career_summary": career_summary(),
        "day_group_totals": day_group_totals(),
        "visiting_faculty": visiting_faculty(),
        "stc_sp": stc_sp_counts(),
        "lab_utilization": lab_utilization(),
        "total_batches": len(get_active_allocations()),
        "total_students": sum(a.students for a in get_active_allocations()),
    }
