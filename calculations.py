#!/usr/bin/env python3
"""All automatic calculations derived from allocations (optionally filtered by report_month)."""

from __future__ import annotations

from collections import defaultdict
from typing import Any, Dict, List, Optional, Tuple

from models import list_allocations, list_careers, list_faculty, list_modules


def get_active_allocations(report_month: Optional[str] = None):
    return list_allocations(active_only=True, report_month=report_month)


def faculty_workload(report_month: Optional[str] = None) -> List[Dict[str, Any]]:
    allocs = get_active_allocations(report_month)
    by_fac: Dict[int, list] = defaultdict(list)
    for a in allocs:
        by_fac[a.faculty_id].append(a)

    rows = []
    for fac in list_faculty(active_only=False):
        items = by_fac.get(fac.id, [])
        load = len(items)
        students = sum(int(a.students or 0) for a in items)
        stc = sum(int(a.students or 0) for a in items if (a.career_code or "").upper() == "STC")
        days = sorted({(a.day_group_code or "") for a in items if a.day_group_code})
        slots = sorted({(a.time_slot_label or "") for a in items if a.time_slot_label})
        timings = ", ".join(filter(None, ["/".join(days), " | ".join(slots)])) if items else "-"
        career_summary = f"{load}-{students}" if load else "-"
        rows.append({
            "initials": fac.initials,
            "status": fac.status,
            "full_name": fac.full_name,
            "load": load,
            "career_summary": career_summary,
            "stc_count": stc,
            "students": students,
            "timings": timings or "-",
            "active": fac.active,
        })
    rows.sort(key=lambda r: (-r["load"], r["initials"]))
    return rows


def faculty_career_matrix(report_month: Optional[str] = None) -> Tuple[List[str], List[Dict[str, Any]]]:
    careers = [c.code for c in list_careers() if not c.is_stc and not c.is_sp]
    if not careers:
        careers = ["CPISM", "DISM", "HDSE I", "HDSE II", "ADSE I", "ADSE II", "CDMA", "AID", "CS"]
    allocs = get_active_allocations(report_month)
    by_fac: Dict[str, Dict[str, int]] = defaultdict(lambda: defaultdict(int))
    name_map = {}
    for a in allocs:
        init = a.faculty_initials or "?"
        name_map[init] = a.faculty_name or init
        code = (a.career_code or "").upper()
        by_fac[init][code] += 1

    rows = []
    for init, counts in sorted(by_fac.items()):
        row = {"initials": init, "full_name": name_map.get(init, init)}
        total = 0
        for c in careers:
            n = counts.get(c, 0)
            row[c] = n
            total += n
        row["STC"] = counts.get("STC", 0)
        row["total"] = total + row["STC"]
        rows.append(row)
    return careers, rows


def career_summary(report_month: Optional[str] = None) -> List[Dict[str, Any]]:
    allocs = get_active_allocations(report_month)
    by_c: Dict[str, Dict[str, int]] = defaultdict(lambda: {"batches": 0, "students": 0})
    for a in allocs:
        code = a.career_code or "Unknown"
        by_c[code]["batches"] += 1
        by_c[code]["students"] += int(a.students or 0)
    total_students = sum(v["students"] for v in by_c.values()) or 1
    rows = []
    for code, v in sorted(by_c.items(), key=lambda x: -x[1]["students"]):
        rows.append({
            "code": code,
            "batches": v["batches"],
            "students": v["students"],
            "pct": v["students"] / total_students,
            "is_stc": code.upper() == "STC",
        })
    return rows


def day_group_totals(report_month: Optional[str] = None) -> Dict[str, Dict[str, int]]:
    allocs = get_active_allocations(report_month)
    out: Dict[str, Dict[str, int]] = defaultdict(lambda: {"batches": 0, "students": 0})
    for a in allocs:
        dg = a.day_group_code or "Other"
        # Normalize Saturday into TTS-like bucket for BSR summary when needed
        key = dg
        if dg.upper() in ("MWF", "M/W/F"):
            key = "MWF"
        elif dg.upper() in ("TTS", "T/T/S", "SATURDAY"):
            key = "TTS" if dg.upper() != "SATURDAY" else "TTS"
        out[key]["batches"] += 1
        out[key]["students"] += int(a.students or 0)
    return dict(out)


def visiting_faculty(report_month: Optional[str] = None) -> List[Dict[str, Any]]:
    allocs = get_active_allocations(report_month)
    by_init: Dict[str, Dict[str, Any]] = {}
    for a in allocs:
        if (a.faculty_status or "").upper() != "V":
            continue
        init = a.faculty_initials or "?"
        if init not in by_init:
            by_init[init] = {
                "name": a.faculty_name or init,
                "count": 0,
                "timings": set(),
            }
        by_init[init]["count"] += 1
        if a.time_slot_label:
            by_init[init]["timings"].add(a.time_slot_label)
    rows = []
    for init, d in sorted(by_init.items()):
        rows.append({
            "initials": init,
            "name": d["name"],
            "count": d["count"],
            "time": " / ".join(sorted(d["timings"])) if d["timings"] else "",
        })
    return rows


def stc_sp_counts(report_month: Optional[str] = None) -> Dict[str, int]:
    allocs = get_active_allocations(report_month)
    stc_b = stc_s = sp_b = sp_s = 0
    for a in allocs:
        code = (a.career_code or "").upper()
        if code == "STC":
            stc_b += 1
            stc_s += int(a.students or 0)
        elif code == "SP":
            sp_b += 1
            sp_s += int(a.students or 0)
    return {
        "stc_batches": stc_b,
        "stc_students": stc_s,
        "sp_batches": sp_b,
        "sp_students": sp_s,
    }


def lab_utilization(report_month: Optional[str] = None) -> List[Dict[str, Any]]:
    allocs = get_active_allocations(report_month)
    by_lab: Dict[str, Dict[str, int]] = defaultdict(lambda: {"batches": 0, "students": 0})
    for a in allocs:
        lab = a.lab_code or "?"
        by_lab[lab]["batches"] += 1
        by_lab[lab]["students"] += int(a.students or 0)
    return [
        {"lab": k, "batches": v["batches"], "students": v["students"]}
        for k, v in sorted(by_lab.items())
    ]


def stc_module_breakdown(report_month: Optional[str] = None) -> List[Dict[str, Any]]:
    allocs = get_active_allocations(report_month)
    counts: Dict[str, int] = defaultdict(int)
    for a in allocs:
        if (a.career_code or "").upper() != "STC":
            continue
        name = (a.module_name or "").strip() or "Unnamed STC"
        counts[name] += int(a.students or 0)
    try:
        for m in list_modules(active_only=True, category="STC"):
            counts.setdefault(m.name, 0)
    except Exception:
        pass
    items = [{"name": k, "students": v} for k, v in counts.items()]
    items.sort(key=lambda x: (-x["students"], x["name"].lower()))
    return items


def sp_module_breakdown(report_month: Optional[str] = None) -> List[Dict[str, Any]]:
    allocs = get_active_allocations(report_month)
    counts: Dict[str, int] = defaultdict(int)
    for a in allocs:
        if (a.career_code or "").upper() != "SP":
            continue
        name = (a.module_name or "").strip() or "Unnamed SP"
        counts[name] += int(a.students or 0)
    try:
        for m in list_modules(active_only=True, category="SP"):
            counts.setdefault(m.name, 0)
    except Exception:
        pass
    items = [{"name": k, "students": v} for k, v in counts.items()]
    items.sort(key=lambda x: (-x["students"], x["name"].lower()))
    return items


def full_dashboard(report_month: Optional[str] = None) -> Dict[str, Any]:
    allocs = get_active_allocations(report_month)
    return {
        "faculty_workload": faculty_workload(report_month),
        "faculty_career_matrix": faculty_career_matrix(report_month),
        "stc_module_breakdown": stc_module_breakdown(report_month),
        "sp_module_breakdown": sp_module_breakdown(report_month),
        "career_summary": career_summary(report_month),
        "day_group_totals": day_group_totals(report_month),
        "visiting_faculty": visiting_faculty(report_month),
        "stc_sp": stc_sp_counts(report_month),
        "lab_utilization": lab_utilization(report_month),
        "total_batches": len(allocs),
        "total_students": sum(int(a.students or 0) for a in allocs),
    }


def build_bsr_month_metrics(year_month: str) -> Dict[str, Any]:
    """BSR figures for a month; STC/SP beginning from prior month snapshot when available."""
    from models import get_monthly_bsr_metrics, previous_year_month

    dash = full_dashboard(report_month=year_month)
    stc = dash["stc_sp"]
    dgt = dash["day_group_totals"]
    mwf = dgt.get("MWF", {"batches": 0, "students": 0})
    tts = dgt.get("TTS", {"batches": 0, "students": 0})

    stc_end = int(stc.get("stc_students") or 0)
    sp_end = int(stc.get("sp_students") or 0)

    # Prefer stored snapshot end values if present (imported historical)
    stored = get_monthly_bsr_metrics(year_month)
    if stored:
        if stored.get("stc_end"):
            stc_end = int(stored["stc_end"])
        if stored.get("sp_end"):
            sp_end = int(stored["sp_end"])
        if stored.get("mwf_batches"):
            mwf = {"batches": int(stored["mwf_batches"]), "students": int(stored.get("mwf_students") or 0)}
        if stored.get("tts_batches"):
            tts = {"batches": int(stored["tts_batches"]), "students": int(stored.get("tts_students") or 0)}

    prev = get_monthly_bsr_metrics(previous_year_month(year_month))
    if prev:
        stc_beg = int(prev.get("stc_end") or 0)
        sp_beg = int(prev.get("sp_end") or 0)
    else:
        stc_beg = int((stored or {}).get("stc_beg") or stc_end)
        sp_beg = int((stored or {}).get("sp_beg") or sp_end)

    return {
        "year_month": year_month,
        "stc_beg": stc_beg,
        "stc_end": stc_end,
        "sp_beg": sp_beg,
        "sp_end": sp_end,
        "total_batches": dash.get("total_batches", 0),
        "total_students": dash.get("total_students", 0),
        "mwf_batches": mwf.get("batches", 0),
        "mwf_students": mwf.get("students", 0),
        "tts_batches": tts.get("batches", 0),
        "tts_students": tts.get("students", 0),
        "dashboard": dash,
    }
