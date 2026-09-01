#!/usr/bin/env python3
"""
Main application window – Aptech Monthly BSR
"""

from __future__ import annotations

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from datetime import datetime, date

# Ensure project root on path
ROOT = Path(__file__).parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

OUTPUT_DIR = ROOT / "Output"

from database import init_db
from ui.excel_preview import ExcelPreviewPanel
from ui.splash import SplashScreen
from models import (
    list_allocations, list_faculty, list_labs, list_careers,
    list_time_slots, list_day_groups,
    add_allocation, update_allocation, deactivate_allocation,
    delete_allocation, clear_all_allocations,
    find_conflicts, find_faculty_conflicts, check_capacity, get_allocation,
    derive_course_from_batch,
    add_faculty, update_faculty, get_faculty, set_faculty_active,
    delete_faculty, count_allocations_for_faculty,
    add_lab, add_career, update_career, delete_career, count_allocations_for_career,
    list_modules, add_module, update_module, delete_module,
    list_available_report_months, list_monthly_bsr_metrics,
)
from database import get_db_path
from calculations import full_dashboard, faculty_workload, career_summary, day_group_totals
from export_excel import export_workbook

# Optional calendar widget
try:
    from tkcalendar import DateEntry  # type: ignore
    HAS_CALENDAR = True
except Exception:
    HAS_CALENDAR = False
    DateEntry = None  # type: ignore



class ModuleDialog(tk.Toplevel):
    def __init__(self, master, module_id=None, on_saved=None):
        super().__init__(master)
        self.module_id = module_id
        self.on_saved = on_saved
        self.title("Edit Module" if module_id else "Add Module / Subject")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        try:
            self.configure(bg=COLORS.get("bg", "#f0f4f8"))
        except Exception:
            pass
        f = ttk.Frame(self, padding=14)
        f.pack(fill=tk.BOTH, expand=True)
        pad = dict(padx=4, pady=4)
        ttk.Label(f, text="Name *").grid(row=0, column=0, sticky="w", **pad)
        self.name_var = tk.StringVar()
        self.name_entry = ttk.Entry(f, textvariable=self.name_var, width=28)
        self.name_entry.grid(row=0, column=1, **pad)
        self._ph = "e.g. FSA / MS Office / Java"
        self.name_var.set(self._ph)
        self.name_entry.configure(foreground="#888")
        self.name_entry.bind("<FocusIn>", self._clear_ph)
        self.name_entry.bind("<FocusOut>", self._restore_ph)

        ttk.Label(f, text="Category").grid(row=1, column=0, sticky="w", **pad)
        self.cat_var = tk.StringVar(value="CAREER")
        ttk.Combobox(
            f, textvariable=self.cat_var, values=("CAREER", "STC", "SP"),
            state="readonly", width=25,
        ).grid(row=1, column=1, **pad)
        if module_id:
            mod = next((m for m in list_modules(active_only=False) if m.id == module_id), None)
            if mod:
                self.name_var.set(mod.name)
                self.name_entry.configure(foreground="#000")
                cat = (mod.category or "CAREER").upper()
                if cat == "SUBJECT":
                    cat = "CAREER"
                self.cat_var.set(cat)
        btn = ttk.Frame(f)
        btn.grid(row=2, column=0, columnspan=2, pady=12)
        ttk.Button(btn, text="Save", command=self._save).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn, text="Cancel", command=self.destroy).pack(side=tk.LEFT, padx=6)

    def _clear_ph(self, _e=None):
        if self.name_var.get() == self._ph:
            self.name_var.set("")
            self.name_entry.configure(foreground="#000")

    def _restore_ph(self, _e=None):
        if not self.name_var.get().strip():
            self.name_var.set(self._ph)
            self.name_entry.configure(foreground="#888")

    def _save(self):
        name = self.name_var.get().strip()
        if not name or name == self._ph:
            messagebox.showerror("Validation", "Name is required.")
            return
        cat = (self.cat_var.get() or "CAREER").upper()
        if cat == "SUBJECT":
            cat = "CAREER"
        try:
            if self.module_id:
                update_module(self.module_id, name, cat)
            else:
                add_module(name, cat)
            if self.on_saved:
                self.on_saved()
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e))


class CareerDialog(tk.Toplevel):
    def __init__(self, master, career_id=None, on_saved=None):
        super().__init__(master)
        self.career_id = career_id
        self.on_saved = on_saved
        self.title("Edit Career" if career_id else "Add Career / Course")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        f = ttk.Frame(self, padding=14)
        f.pack(fill=tk.BOTH, expand=True)

        pad = dict(padx=4, pady=4)
        ttk.Label(f, text="Code *").grid(row=0, column=0, sticky="w", **pad)
        self.code_var = tk.StringVar()
        ttk.Entry(f, textvariable=self.code_var, width=24).grid(row=0, column=1, **pad)

        ttk.Label(f, text="Name").grid(row=1, column=0, sticky="w", **pad)
        self.name_var = tk.StringVar()
        ttk.Entry(f, textvariable=self.name_var, width=24).grid(row=1, column=1, **pad)

        ttk.Label(f, text="Type").grid(row=2, column=0, sticky="w", **pad)
        self.type_var = tk.StringVar(value="Career")
        ttk.Combobox(
            f, textvariable=self.type_var, values=("Career", "STC", "SP"),
            state="readonly", width=21,
        ).grid(row=2, column=1, **pad)

        if career_id:
            from models import list_careers
            car = next((c for c in list_careers() if c.id == career_id), None)
            if car:
                self.code_var.set(car.code)
                self.name_var.set(car.name or "")
                self.type_var.set("STC" if car.is_stc else ("SP" if car.is_sp else "Career"))

        btn = ttk.Frame(f)
        btn.grid(row=3, column=0, columnspan=2, pady=12)
        ttk.Button(btn, text="Save", command=self._save).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn, text="Cancel", command=self.destroy).pack(side=tk.LEFT, padx=6)

    def _save(self):
        code = self.code_var.get().strip()
        if not code:
            messagebox.showerror("Validation", "Code is required.")
            return
        name = self.name_var.get().strip() or code
        typ = self.type_var.get()
        is_stc = typ == "STC"
        is_sp = typ == "SP"
        try:
            if self.career_id:
                update_career(self.career_id, code, name, is_stc, is_sp)
            else:
                add_career(code, name, is_stc, is_sp)
            if self.on_saved:
                self.on_saved()
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e))



class AllocationForm(tk.Toplevel):
    """Modal form to add / edit a single allocation."""

    def __init__(self, master, allocation_id=None, on_save=None):
        super().__init__(master)
        self.allocation_id = allocation_id
        self.on_save = on_save
        self.title("Edit Allocation" if allocation_id else "Add New Allocation")
        self.geometry("540x560")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        try:
            self.configure(bg=COLORS["bg"])
        except Exception:
            pass

        self.faculty = list_faculty()
        self.labs = list_labs()
        self.careers = list_careers()
        self.slots = list_time_slots()
        self.days = list_day_groups()

        self._build()
        if allocation_id:
            self._load(allocation_id)

    def _build(self):
        pad = {"padx": 8, "pady": 4}
        f = ttk.Frame(self, padding=12)
        f.pack(fill=tk.BOTH, expand=True)

        # Faculty
        self.fac_lbl = ttk.Label(f, text="Faculty *")
        self.fac_lbl.grid(row=0, column=0, sticky="w", **pad)
        self.fac_var = tk.StringVar(value="— Select Faculty —")
        fac_values = ["— Select Faculty —"] + [f"{x.initials} – {x.full_name}" for x in self.faculty]
        self.fac_combo = ttk.Combobox(f, textvariable=self.fac_var, values=fac_values, width=40, state="readonly")
        self.fac_combo.grid(row=0, column=1, columnspan=2, sticky="ew", **pad)

        # Career
        self.career_lbl = ttk.Label(f, text="Career / Program *")
        self.career_lbl.grid(row=1, column=0, sticky="w", **pad)
        self.career_var = tk.StringVar(value="— Select Career —")
        self.career_combo = ttk.Combobox(
            f, textvariable=self.career_var,
            values=["— Select Career —"] + [c.code for c in self.careers], width=40, state="readonly"
        )
        self.career_combo.grid(row=1, column=1, columnspan=2, sticky="ew", **pad)

        # Lab
        self.lab_lbl = ttk.Label(f, text="Lab *")
        self.lab_lbl.grid(row=2, column=0, sticky="w", **pad)
        self.lab_var = tk.StringVar(value="— Select Lab —")
        self.lab_combo = ttk.Combobox(
            f, textvariable=self.lab_var,
            values=["— Select Lab —"] + [f"{l.lab_code} ({l.pcs} PCs)" for l in self.labs],
            width=40, state="readonly"
        )
        self.lab_combo.grid(row=2, column=1, columnspan=2, sticky="ew", **pad)

        # Time slot
        self.slot_lbl = ttk.Label(f, text="Time Slot *")
        self.slot_lbl.grid(row=3, column=0, sticky="w", **pad)
        self.slot_var = tk.StringVar(value="— Select Time Slot —")
        self.slot_combo = ttk.Combobox(
            f, textvariable=self.slot_var,
            values=["— Select Time Slot —"] + [s.label for s in self.slots],
            width=40, state="readonly"
        )
        self.slot_combo.grid(row=3, column=1, columnspan=2, sticky="ew", **pad)

        # Day group
        self.day_lbl = ttk.Label(f, text="Day Group *")
        self.day_lbl.grid(row=4, column=0, sticky="w", **pad)
        self.day_var = tk.StringVar(value="— Select Day Group —")
        self.day_combo = ttk.Combobox(
            f, textvariable=self.day_var,
            values=["— Select Day Group —"] + [d.code for d in self.days],
            width=40, state="readonly"
        )
        self.day_combo.grid(row=4, column=1, columnspan=2, sticky="ew", **pad)

        # Batch code
        ttk.Label(f, text="Batch Code").grid(row=5, column=0, sticky="w", **pad)
        self.batch_var = tk.StringVar()
        self.batch_entry = ttk.Entry(f, textvariable=self.batch_var, width=42)
        self.batch_entry.grid(row=5, column=1, columnspan=2, sticky="ew", **pad)
        self._add_placeholder(self.batch_entry, self.batch_var, "e.g. AI-202607D")

        # Module / Subject – from Modules catalog (dynamic)
        ttk.Label(f, text="Module / Subject").grid(row=6, column=0, sticky="w", **pad)
        self.module_var = tk.StringVar()
        try:
            mod_opts = [m.name for m in list_modules(active_only=True)]
        except Exception:
            mod_opts = []
        self.module_combo = ttk.Combobox(
            f, textvariable=self.module_var, values=mod_opts, width=40,
        )
        self.module_combo.grid(row=6, column=1, columnspan=2, sticky="ew", **pad)
        self.module_entry = self.module_combo
        self._module_ph = "— Select Module / Subject —"
        self.module_var.set(self._module_ph)
        self.module_combo.configure(foreground="#888")
        def _mod_in(_e=None):
            if self.module_var.get() == self._module_ph:
                self.module_var.set("")
                self.module_combo.configure(foreground="#000")
        def _mod_out(_e=None):
            if not self.module_var.get().strip():
                self.module_var.set(self._module_ph)
                self.module_combo.configure(foreground="#888")
        self.module_combo.bind("<FocusIn>", _mod_in)
        self.module_combo.bind("<FocusOut>", _mod_out)
        self.module_combo.bind("<<ComboboxSelected>>", lambda e: self.module_combo.configure(foreground="#000"))

        # Students
        self.students_lbl = ttk.Label(f, text="No. of Students *")
        self.students_lbl.grid(row=7, column=0, sticky="w", **pad)
        self.students_var = tk.StringVar()
        self.students_entry = ttk.Entry(f, textvariable=self.students_var, width=12)
        self.students_entry.grid(row=7, column=1, sticky="w", **pad)
        self._add_placeholder(self.students_entry, self.students_var, "e.g. 10")

        # Module Start Date – simple text date entry
        ttk.Label(f, text="Module Start Date").grid(row=8, column=0, sticky="w", **pad)
        self.mod_start_var = tk.StringVar()
        self.mod_start_entry = ttk.Entry(f, textvariable=self.mod_start_var, width=16)
        self.mod_start_entry.grid(row=8, column=1, sticky="w", **pad)
        self._add_placeholder(self.mod_start_entry, self.mod_start_var, "yyyy-mm-dd")
        ttk.Label(f, text="e.g. 2026-07-17", foreground="#666").grid(
            row=8, column=2, sticky="w"
        )

        # Batch Start Date – simple text date entry
        ttk.Label(f, text="Batch Start Date").grid(row=9, column=0, sticky="w", **pad)
        self.act_start_var = tk.StringVar()
        self.act_start_entry = ttk.Entry(f, textvariable=self.act_start_var, width=16)
        self.act_start_entry.grid(row=9, column=1, sticky="w", **pad)
        self._add_placeholder(self.act_start_entry, self.act_start_var, "yyyy-mm-dd")
        ttk.Label(f, text="e.g. 2026-07-13", foreground="#666").grid(
            row=9, column=2, sticky="w"
        )

        # Admission Open
        ttk.Label(f, text="Admission Open").grid(row=10, column=0, sticky="w", **pad)
        self.adm_var = tk.StringVar(value="No")
        self.adm_combo = ttk.Combobox(
            f, textvariable=self.adm_var, values=["Yes", "No"],
            state="readonly", width=8,
        )
        self.adm_combo.grid(row=10, column=1, sticky="w", **pad)
        self.adm_hint = ttk.Label(
            f, text="", foreground="#666", wraplength=320
        )
        self.adm_hint.grid(row=10, column=2, sticky="w", **pad)
        self.adm_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_admission_changed())
        self._on_admission_changed()

        # Notes
        ttk.Label(f, text="Notes").grid(row=11, column=0, sticky="nw", **pad)
        self.notes_txt = tk.Text(f, height=3, width=40, fg="#888")
        self.notes_txt.grid(row=11, column=1, columnspan=2, sticky="ew", **pad)
        self._notes_placeholder = "Optional notes (e.g. Short Course, special instructions)"
        self.notes_txt.insert("1.0", self._notes_placeholder)
        self.notes_txt.bind("<FocusIn>", self._notes_focus_in)
        self.notes_txt.bind("<FocusOut>", self._notes_focus_out)

        # Buttons
        btn_f = ttk.Frame(f)
        btn_f.grid(row=12, column=0, columnspan=3, pady=16)
        ttk.Button(btn_f, text="Save", command=self._save, width=12).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_f, text="Cancel", command=self.destroy, width=12).pack(side=tk.LEFT, padx=6)

    def _add_placeholder(self, entry, var, placeholder: str):
        """Show grey placeholder text that clears on focus."""
        var.set(placeholder)
        entry.configure(foreground="#888")

        def on_in(_e=None):
            if var.get() == placeholder:
                var.set("")
                entry.configure(foreground="#000")

        def on_out(_e=None):
            if not var.get().strip():
                var.set(placeholder)
                entry.configure(foreground="#888")

        entry.bind("<FocusIn>", on_in)
        entry.bind("<FocusOut>", on_out)
        # Keep real value tracking for save
        entry._placeholder = placeholder

    def _notes_focus_in(self, _e=None):
        if self.notes_txt.get("1.0", "end-1c") == self._notes_placeholder:
            self.notes_txt.delete("1.0", tk.END)
            self.notes_txt.configure(fg="#000")

    def _notes_focus_out(self, _e=None):
        if not self.notes_txt.get("1.0", "end-1c").strip():
            self.notes_txt.insert("1.0", self._notes_placeholder)
            self.notes_txt.configure(fg="#888")

    def _entry_value(self, var, entry=None):
        """Return entry text, treating placeholder as empty."""
        val = var.get().strip()
        ph = getattr(entry, "_placeholder", None) if entry is not None else None
        if ph and val == ph:
            return ""
        return val

    def _load(self, aid: int):
        a = get_allocation(aid)
        if not a:
            return
        self.fac_var.set(f"{a.faculty_initials} – {a.faculty_name}")
        self.career_var.set(a.career_code)
        self.lab_var.set(f"{a.lab_code} ({a.lab_pcs} PCs)")
        self.slot_var.set(a.time_slot_label)
        self.day_var.set(a.day_group_code)
        if a.batch_code:
            self.batch_var.set(a.batch_code)
            if hasattr(self, "batch_entry"):
                self.batch_entry.configure(foreground="#000")
        if a.module_name:
            self.module_var.set(a.module_name)
            self.module_combo.configure(foreground="#000")
        else:
            self.module_var.set(getattr(self, "_module_ph", "— Select Module / Subject —"))
            self.module_combo.configure(foreground="#888")
        self.students_var.set(str(a.students))
        if hasattr(self, "students_entry"):
            self.students_entry.configure(foreground="#000")
        if a.module_start_date:
            self.mod_start_var.set(str(a.module_start_date)[:10])
            self.mod_start_entry.configure(foreground="#000")
        if a.actual_start_date:
            self.act_start_var.set(str(a.actual_start_date)[:10])
            self.act_start_entry.configure(foreground="#000")
        self.adm_var.set("Yes" if a.is_admission_open else "No")
        self._on_admission_changed()
        if a.notes:
            self.notes_txt.delete("1.0", tk.END)
            self.notes_txt.insert("1.0", a.notes)
            self.notes_txt.configure(fg="#000")

    def _on_admission_changed(self):
        """Toggle required markers and hint when Admission Open Yes/No changes."""
        is_yes = self.adm_var.get().strip().lower() in ("yes", "y", "1", "true")
        if is_yes:
            self.fac_lbl.configure(text="Faculty")
            self.career_lbl.configure(text="Career / Program")
            self.students_lbl.configure(text="No. of Students")
            self.lab_lbl.configure(text="Lab *")
            self.slot_lbl.configure(text="Time Slot *")
            self.day_lbl.configure(text="Day Group *")
            self.adm_hint.configure(
                text="Yes → only Lab, Time Slot & Day required"
            )
        else:
            self.fac_lbl.configure(text="Faculty *")
            self.career_lbl.configure(text="Career / Program *")
            self.students_lbl.configure(text="No. of Students *")
            self.lab_lbl.configure(text="Lab *")
            self.slot_lbl.configure(text="Time Slot *")
            self.day_lbl.configure(text="Day Group *")
            self.adm_hint.configure(text="")

    def _resolve_ids(self):
        fac_txt = (self.fac_var.get() or "").strip()
        if fac_txt.startswith("—") or fac_txt.lower().startswith("select"):
            faculty_id = None
        else:
            faculty_id = next(
                (f.id for f in self.faculty if f"{f.initials} – {f.full_name}" == fac_txt),
                None,
            )

        career_txt = (self.career_var.get() or "").strip()
        if career_txt.startswith("—") or career_txt.lower().startswith("select"):
            career_id = None
        else:
            career_id = next((c.id for c in self.careers if c.code == career_txt), None)

        lab_txt = (self.lab_var.get() or "").strip()
        if lab_txt.startswith("—") or lab_txt.lower().startswith("select"):
            lab_id = None
        else:
            lab_id = next(
                (l.id for l in self.labs if f"{l.lab_code} ({l.pcs} PCs)" == lab_txt),
                None,
            )

        slot_txt = (self.slot_var.get() or "").strip()
        if slot_txt.startswith("—") or slot_txt.lower().startswith("select"):
            slot_id = None
        else:
            slot_id = next((s.id for s in self.slots if s.label == slot_txt), None)

        day_txt = (self.day_var.get() or "").strip()
        if day_txt.startswith("—") or day_txt.lower().startswith("select"):
            day_id = None
        else:
            day_id = next((d.id for d in self.days if d.code == day_txt), None)

        return faculty_id, career_id, lab_id, slot_id, day_id

    def _save(self):
        faculty_id, career_id, lab_id, slot_id, day_id = self._resolve_ids()
        is_adm_open = self.adm_var.get().strip().lower() in ("yes", "y", "1", "true")

        # Lab + Time Slot + Day are always required (needed to place the block).
        # When Admission Open = Yes, Faculty / Career / Batch / Module / Students
        # become optional. When Admission Open = No, they remain mandatory.
        if not all([lab_id, slot_id, day_id]):
            messagebox.showerror(
                "Missing data",
                "Lab, Time Slot and Day Group are required.",
                parent=self,
            )
            return

        if not is_adm_open:
            if not all([faculty_id, career_id]):
                messagebox.showerror(
                    "Missing data",
                    "Faculty and Career are required when Admission Open is No.",
                    parent=self,
                )
                return

        try:
            students = int(self._entry_value(self.students_var, getattr(self, "students_entry", None)) or 0)
        except ValueError:
            messagebox.showerror("Invalid", "Students must be a number.", parent=self)
            return

        batch_code = self._entry_value(self.batch_var, getattr(self, "batch_entry", None)) or None
        # Course/series is derived from batch code (matches Excel template style)
        course_title = derive_course_from_batch(batch_code)

        # Capacity / conflict checks only when there is real student load or faculty
        if students > 0:
            ok, pcs = check_capacity(lab_id, students)
            if not ok:
                if not messagebox.askyesno(
                    "Capacity Warning",
                    f"Students ({students}) exceed lab capacity ({pcs} PCs).\nSave anyway?",
                    parent=self,
                ):
                    return

        if not is_adm_open:
            lab_conflicts = find_conflicts(lab_id, slot_id, day_id, exclude_id=self.allocation_id)
            if lab_conflicts:
                names = ", ".join(
                    f"{c.faculty_initials}/{c.batch_code or c.course_title or '?'}"
                    for c in lab_conflicts
                )
                if not messagebox.askyesno(
                    "Lab Schedule Conflict",
                    f"This lab + time + day is already used by:\n{names}\n\n"
                    "Save anyway? (existing allocation is NOT overwritten)",
                    parent=self,
                ):
                    return

        if faculty_id and not is_adm_open:
            fac_conflicts = find_faculty_conflicts(
                faculty_id, slot_id, day_id, exclude_id=self.allocation_id
            )
            if fac_conflicts:
                names = ", ".join(
                    f"{c.lab_code}/{c.batch_code or c.career_code or '?'}"
                    for c in fac_conflicts
                )
                if not messagebox.askyesno(
                    "Faculty Time Conflict",
                    f"This teacher is already allocated at the same time + day in:\n{names}\n\n"
                    "Save anyway? (a new allocation will be added; nothing is overwritten)",
                    parent=self,
                ):
                    return

        data = {
            "faculty_id": faculty_id,          # may be None when Admission Open = Yes
            "career_id": career_id,            # may be None when Admission Open = Yes
            "lab_id": lab_id,
            "time_slot_id": slot_id,
            "day_group_id": day_id,
            "batch_code": batch_code,
            "course_title": course_title,
            "module_name": (None if self.module_var.get().strip() in ("", getattr(self, "_module_ph", "— Select Module / Subject —")) else self.module_var.get().strip()),
            "students": students,
            "module_start_date": (self.mod_start_var.get().strip() or None) if self.mod_start_var.get().strip() not in ("", "yyyy-mm-dd") else None,
            "actual_start_date": (self.act_start_var.get().strip() or None) if self.act_start_var.get().strip() not in ("", "yyyy-mm-dd") else None,
            "notes": (lambda n: None if (not n or n == getattr(self, "_notes_placeholder", "")) else n)(
                self.notes_txt.get("1.0", "end-1c").strip()
            ),
            "is_admission_open": is_adm_open,
        }

        try:
            if self.allocation_id:
                update_allocation(self.allocation_id, data)
            else:
                add_allocation(data)
            if self.on_save:
                self.on_save()
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)


class FacultyForm(tk.Toplevel):
    """Modal form to add / edit a faculty member."""

    def __init__(self, master, faculty_id=None, on_save=None):
        super().__init__(master)
        self.faculty_id = faculty_id
        self.on_save = on_save
        self.title("Edit Faculty" if faculty_id else "Add Faculty")
        self.geometry("420x220")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        pad = {"padx": 8, "pady": 6}
        f = ttk.Frame(self, padding=12)
        f.pack(fill=tk.BOTH, expand=True)

        ttk.Label(f, text="Initials *").grid(row=0, column=0, sticky="w", **pad)
        self.init_var = tk.StringVar(value="e.g. AW")
        self.init_entry = ttk.Entry(f, textvariable=self.init_var, width=20, foreground="#888")
        self.init_entry.grid(row=0, column=1, sticky="w", **pad)
        ttk.Label(f, text="e.g. AW", foreground="#666").grid(row=0, column=2, sticky="w")

        ttk.Label(f, text="Full Name *").grid(row=1, column=0, sticky="w", **pad)
        self.name_var = tk.StringVar(value="e.g. Abdul Wahab")
        self.name_entry = ttk.Entry(f, textvariable=self.name_var, width=36, foreground="#888")
        self.name_entry.grid(row=1, column=1, columnspan=2, sticky="ew", **pad)

        ttk.Label(f, text="Status *").grid(row=2, column=0, sticky="w", **pad)
        self.status_var = tk.StringVar(value="P")
        ttk.Combobox(
            f, textvariable=self.status_var,
            values=["P – Permanent", "V – Visiting"],
            state="readonly", width=18,
        ).grid(row=2, column=1, sticky="w", **pad)

        def _ph_in(entry, var, ph):
            def handler(_e=None):
                if var.get() == ph:
                    var.set("")
                    entry.configure(foreground="#000")
            return handler
        def _ph_out(entry, var, ph):
            def handler(_e=None):
                if not var.get().strip():
                    var.set(ph)
                    entry.configure(foreground="#888")
            return handler
        self.init_entry.bind("<FocusIn>", _ph_in(self.init_entry, self.init_var, "e.g. AW"))
        self.init_entry.bind("<FocusOut>", _ph_out(self.init_entry, self.init_var, "e.g. AW"))
        self.name_entry.bind("<FocusIn>", _ph_in(self.name_entry, self.name_var, "e.g. Abdul Wahab"))
        self.name_entry.bind("<FocusOut>", _ph_out(self.name_entry, self.name_var, "e.g. Abdul Wahab"))

        btn = ttk.Frame(f)
        btn.grid(row=3, column=0, columnspan=3, pady=14)
        ttk.Button(btn, text="Save", command=self._save, width=12).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn, text="Cancel", command=self.destroy, width=12).pack(side=tk.LEFT)

        if faculty_id:
            fac = get_faculty(faculty_id)
            if fac:
                self.init_var.set(fac.initials)
                self.name_var.set(fac.full_name)
                self.status_var.set("P – Permanent" if fac.status.upper() == "P" else "V – Visiting")
                if hasattr(self, "init_entry"):
                    self.init_entry.configure(foreground="#000")
                if hasattr(self, "name_entry"):
                    self.name_entry.configure(foreground="#000")

    def _save(self):
        initials = self.init_var.get().strip().upper()
        if initials in ("E.G. AW", "E.G.AW"):
            initials = ""
        name = self.name_var.get().strip()
        if name.startswith("e.g."):
            name = ""
        status_raw = self.status_var.get()
        status = "P" if status_raw.startswith("P") else "V"
        if not initials or not name:
            messagebox.showwarning("Required", "Initials and Full Name are required.", parent=self)
            return
        try:
            if self.faculty_id:
                update_faculty(self.faculty_id, initials, name, status)
            else:
                add_faculty(initials, name, status)
            if self.on_save:
                self.on_save()
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)




# ── Modern UI theme ──────────────────────────────────────────────
COLORS = {
    # Soft futuristic — light, eye-friendly, unisex
    "bg": "#f4f6fb",
    "surface": "#ffffff",
    "header": "#1a1f36",
    "header_text": "#f0f4ff",
    "accent": "#5b6cff",
    "accent_hover": "#4a58e0",
    "accent_soft": "#e8ebff",
    "success": "#2d9f78",
    "danger": "#e05a6a",
    "muted": "#7a8499",
    "border": "#dce3f0",
    "row_alt": "#f8f9fd",
    "select": "#e4e8ff",
    "select_text": "#1a1f36",
    "tree_header": "#2a3150",
    "month_banner": "#eef1ff",
    "month_text": "#2a3150",
    "glow": "#c5ccff",
    "card_shadow": "#e6eaf5",
}


class WindowsStyleProgressBar(tk.Canvas):
    """Windows-like determinate progress bar with continuous glossy shine.

    Keeps the app on the clam colour theme while drawing a bar that mimics
    the Windows copy/delete progress look: soft blue fill + moving highlight.
    """

    def __init__(self, master, height: int = 16, **kwargs):
        self._height = max(12, int(height))
        super().__init__(
            master,
            height=self._height,
            bg="#e6e6e6",
            highlightthickness=1,
            highlightbackground="#adadad",
            bd=0,
            **kwargs,
        )
        self._value = 0.0
        self._maximum = 100.0
        self._running = False
        self._phase = 0.0
        self._after_id = None
        self.bind("<Configure>", lambda _e: self._draw())

    def __setitem__(self, key, value):
        if key == "value":
            self.set(value)
        elif key == "maximum":
            self._maximum = float(value) or 100.0
            self._draw()
        else:
            super().__setitem__(key, value)

    def __getitem__(self, key):
        if key == "value":
            return self._value
        if key == "maximum":
            return self._maximum
        return super().__getitem__(key)

    def set(self, value: float) -> None:
        self._value = max(0.0, min(float(value), self._maximum))
        self._draw()

    def start(self, _interval: int = 16) -> None:
        if self._running:
            return
        self._running = True
        self._phase = 0.0
        self._tick()

    def stop(self) -> None:
        self._running = False
        if self._after_id is not None:
            try:
                self.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        self._draw()

    def _tick(self) -> None:
        if not self._running:
            return
        # Continuous left→right loop (Windows-style sheen)
        self._phase = (self._phase + 0.022) % 1.0
        self._draw()
        self._after_id = self.after(20, self._tick)

    def _draw(self) -> None:
        self.delete("all")
        w = max(int(self.winfo_width()), 4)
        h = max(int(self.winfo_height()), self._height)

        # Trough (light grey like Windows)
        self.create_rectangle(0, 0, w, h, fill="#e6e6e6", outline="")
        # subtle inner border
        self.create_line(0, 0, w, 0, fill="#f5f5f5")
        self.create_line(0, h - 1, w, h - 1, fill="#cfcfcf")

        frac = 0.0 if self._maximum <= 0 else self._value / self._maximum
        fw = int(w * frac)
        if fw < 2:
            return

        # Windows-style green/blue progress fill with vertical gloss
        # (close to Windows 10 accent progress)
        self.create_rectangle(0, 0, fw, h, fill="#06b025", outline="")
        # lighter top half (gloss)
        self.create_rectangle(0, 0, fw, max(2, h // 2), fill="#3ddc65", outline="")
        # bright top edge
        self.create_rectangle(0, 0, fw, max(1, h // 5), fill="#7aef95", outline="")
        # darker bottom edge
        self.create_rectangle(0, max(0, h - max(2, h // 4)), fw, h, fill="#05941f", outline="")

        if not self._running or fw < 8:
            return

        # Moving soft white sheen across the fill (Windows highlight)
        band = max(36, int(fw * 0.45))
        # phase 0..1 maps to travel from -band to fw
        x0 = int(self._phase * (fw + band)) - band
        x1 = x0 + band

        # Clip sheen to filled region by drawing only overlapping part
        left = max(0, x0)
        right = min(fw, x1)
        if right <= left:
            return

        # Layered stipple = soft translucent look without true alpha
        self.create_rectangle(left, 0, right, h, fill="#ffffff", stipple="gray25", outline="")
        mid_pad = max(4, (right - left) // 4)
        if right - left > mid_pad * 2:
            self.create_rectangle(
                left + mid_pad, 0, right - mid_pad, h,
                fill="#ffffff", stipple="gray50", outline="",
            )
        core = max(4, (right - left) // 6)
        cx = (left + right) // 2
        self.create_rectangle(
            max(left, cx - core), 0, min(right, cx + core), h,
            fill="#ffffff", stipple="gray75", outline="",
        )


def apply_app_theme(root: tk.Tk) -> ttk.Style:
    """Apply the app's custom colour theme (clam) across the UI."""
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    c = COLORS
    root.configure(bg=c["bg"])

    font_ui = ("Segoe UI", 10)
    font_ui_bold = ("Segoe UI", 10, "bold")
    font_title = ("Segoe UI", 15, "bold")
    font_small = ("Segoe UI", 9)

    style.configure(".", font=font_ui, background=c["bg"], foreground="#1e293b")
    style.configure("TFrame", background=c["bg"])
    style.configure("Surface.TFrame", background=c["surface"])
    style.configure("Header.TFrame", background=c["header"])
    style.configure("TLabel", background=c["bg"], foreground="#1e293b", font=font_ui)
    style.configure("Header.TLabel", background=c["header"], foreground=c["header_text"], font=font_title)
    style.configure("SubHeader.TLabel", background=c["header"], foreground="#a8b0d0", font=font_small)
    style.configure("Muted.TLabel", background=c["bg"], foreground=c["muted"], font=font_small)
    style.configure("Hint.TLabel", background=c["bg"], foreground=c["muted"], font=font_small)

    style.configure(
        "TButton",
        font=font_ui,
        padding=(12, 6),
        background="#e8ecf6",
        foreground="#2a3150",
        borderwidth=0,
        focuscolor="",
    )
    style.map(
        "TButton",
        background=[("active", "#d5dcf0"), ("disabled", "#eef1f8")],
        foreground=[("disabled", "#a0a8bc")],
    )

    # Primary action (Generate BSR)
    style.configure(
        "Primary.TButton",
        font=font_ui_bold,
        padding=(14, 7),
        background=c["accent"],
        foreground="#ffffff",
        borderwidth=0,
    )
    style.map(
        "Primary.TButton",
        background=[("active", c["accent_hover"]), ("disabled", "#9aa5b5")],
        foreground=[("disabled", "#e5e7eb")],
    )

    # Success / danger
    style.configure("Success.TButton", font=font_ui, padding=(12, 6), background=c["success"], foreground="#fff")
    style.map("Success.TButton", background=[("active", "#047857")])
    style.configure("Danger.TButton", font=font_ui, padding=(12, 6), background=c["danger"], foreground="#fff")
    style.map("Danger.TButton", background=[("active", "#b91c1c")])

    style.configure(
        "TEntry",
        fieldbackground=c["surface"],
        foreground="#1e293b",
        padding=5,
        borderwidth=1,
        relief="solid",
    )
    style.configure(
        "TCombobox",
        fieldbackground=c["surface"],
        background=c["surface"],
        foreground="#1e293b",
        padding=4,
        arrowsize=14,
    )
    style.map(
        "TCombobox",
        fieldbackground=[("readonly", c["surface"])],
        foreground=[("readonly", "#1e293b")],
    )

    style.configure(
        "Treeview",
        background=c["surface"],
        fieldbackground=c["surface"],
        foreground="#1e293b",
        rowheight=28,
        font=font_ui,
        borderwidth=0,
    )
    style.configure(
        "Treeview.Heading",
        background=c["tree_header"],
        foreground="#ffffff",
        font=font_ui_bold,
        relief="flat",
        padding=(6, 6),
    )
    style.map(
        "Treeview",
        background=[("selected", c["select"])],
        foreground=[("selected", c["select_text"])],
    )
    style.map("Treeview.Heading", background=[("active", "#274b75")])

    style.configure("TNotebook", background=c["bg"], borderwidth=0, tabmargins=(6, 6, 6, 0))
    style.configure(
        "TNotebook.Tab",
        background="#e2e8f0",
        foreground="#475569",
        padding=(16, 8),
        font=font_ui_bold,
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", c["surface"]), ("active", "#f1f5f9")],
        foreground=[("selected", c["accent"]), ("active", "#1e293b")],
        expand=[("selected", (0, 0, 0, 2))],
    )

    style.configure(
        "Horizontal.TProgressbar",
        troughcolor="#e8ebff",
        background=c["accent"],
        thickness=12,
        borderwidth=0,
    )
    style.configure(
        "Export.Horizontal.TProgressbar",
        troughcolor="#e8ebff",
        background=c["accent"],
        thickness=14,
        borderwidth=0,
    )
    style.configure("TLabelframe", background=c["bg"], foreground="#1e293b")
    style.configure("TLabelframe.Label", background=c["bg"], foreground=c["accent"], font=font_ui_bold)
    style.configure("TScrollbar", background="#e2e8f0", troughcolor=c["bg"], borderwidth=0, arrowsize=12)
    style.configure("Status.TLabel", background="#e8ecf6", foreground=c["muted"], font=font_small, padding=(8, 4))

    return style


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        # Hide main window until splash finishes (professional startup)
        try:
            self.withdraw()
        except Exception:
            pass

        self.title("Aptech Monthly BSR – Batch Status Report")
        self.geometry("1180x720")
        self.minsize(980, 640)
        try:
            self.iconbitmap(default="")
        except Exception:
            pass

        splash = None
        try:
            splash = SplashScreen(self, app_name="Aptech Monthly BSR", version="1.0")
            splash.set_status("Connecting to database…")
            splash.pulse()

            init_db()
            splash.set_status("Loading application theme…")
            splash.pulse()

            self.style = apply_app_theme(self)
            splash.set_status("Building interface…")
            splash.pulse()

            self._build_ui()
            splash.set_status("Loading report months…")
            splash.pulse()

            # Default to latest month that actually has allocation rows
            try:
                months = list_available_report_months()
                self.bsr_month_combo.configure(values=self._month_choices())
                chosen = None
                for m in months:
                    if list_allocations(active_only=True, report_month=m):
                        chosen = m
                        break
                if chosen:
                    self.bsr_month_var.set(chosen)
                elif months:
                    self.bsr_month_var.set(months[0])
            except Exception:
                pass

            splash.set_status("Refreshing data…")
            splash.pulse()
            self.refresh()
            splash.set_status("Ready")
            splash.pulse()
        except Exception:
            # Ensure splash never traps the user on failure
            if splash is not None:
                try:
                    splash.close_splash()
                except Exception:
                    pass
            try:
                self.deiconify()
            except Exception:
                pass
            raise
        finally:
            if splash is not None:
                try:
                    splash.close_splash()
                except Exception:
                    pass
            try:
                self.deiconify()
                self.lift()
                self.focus_force()
            except Exception:
                pass

    def _build_ui(self):
        # ── Dark header ──
        header = tk.Frame(self, bg=COLORS["header"], height=68)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        left = tk.Frame(header, bg=COLORS["header"])
        left.pack(side=tk.LEFT, fill=tk.Y, padx=18, pady=10)
        tk.Label(
            left, text="Aptech Monthly BSR",
            bg=COLORS["header"], fg="#ffffff",
            font=("Segoe UI", 16, "bold"),
        ).pack(anchor="w")
        tk.Label(
            left, text="Batch Status Report  ·  Faculty Workload  ·  Monthly Generation",
            bg=COLORS["header"], fg="#a8b0d0",
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(2, 0))

        right = tk.Frame(header, bg=COLORS["header"])
        right.pack(side=tk.RIGHT, padx=16, pady=14)
        tk.Label(
            right, text="Report Month",
            bg=COLORS["header"], fg="#a8b0d0", font=("Segoe UI", 9),
        ).pack(side=tk.LEFT, padx=(0, 6))
        self.bsr_month_var = tk.StringVar(value=datetime.now().strftime("%Y-%m"))
        self.bsr_month_combo = ttk.Combobox(
            right, textvariable=self.bsr_month_var, width=10,
            values=self._month_choices(), state="readonly",
        )
        self.bsr_month_combo.pack(side=tk.LEFT)
        self.bsr_month_combo.bind("<<ComboboxSelected>>", self._on_month_changed)

        # ── Toolbar ──
        top = ttk.Frame(self, padding=(10, 8))
        top.pack(fill=tk.X)

        self.export_btn = ttk.Button(
            top, text="✦  Generate Monthly BSR", style="Primary.TButton", command=self._export
        )
        self.export_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.open_folder_btn = ttk.Button(
            top, text="Open Output Folder", command=self._open_output_folder
        )
        self.open_folder_btn.pack(side=tk.LEFT, padx=4)
        self.open_folder_btn.pack_forget()

        ttk.Button(top, text="Refresh", command=self.refresh).pack(side=tk.LEFT, padx=4)

        ttk.Button(top, text="＋ Add Allocation", command=self._add).pack(side=tk.RIGHT, padx=4)
        ttk.Button(top, text="Edit Selected", command=self._edit).pack(side=tk.RIGHT, padx=4)
        ttk.Button(top, text="Deactivate", command=self._deactivate).pack(side=tk.RIGHT, padx=4)
        ttk.Button(top, text="Delete Selected", command=self._delete_one).pack(side=tk.RIGHT, padx=4)
        ttk.Button(top, text="Clear All Data", command=self._clear_all).pack(side=tk.RIGHT, padx=8)

        # Windows-style progress bars (custom draw keeps app colour theme intact)
        self.progress = WindowsStyleProgressBar(self, height=12)
        self._exporting = False

        # Export progress card (under toolbar, visible while generating)
        self.export_loader = tk.Frame(
            self, bg=COLORS["surface"],
            highlightbackground=COLORS["glow"], highlightthickness=1,
        )
        el_inner = tk.Frame(self.export_loader, bg=COLORS["surface"])
        el_inner.pack(fill=tk.X, padx=16, pady=10)
        self.export_loader_title = tk.Label(
            el_inner, text="Generating Monthly BSR…",
            bg=COLORS["surface"], fg=COLORS["month_text"],
            font=("Segoe UI", 11, "bold"), anchor="w",
        )
        self.export_loader_title.pack(fill=tk.X)
        self.export_loader_sub = tk.Label(
            el_inner, text="Preparing workbook — please wait",
            bg=COLORS["surface"], fg=COLORS["muted"],
            font=("Segoe UI", 9), anchor="w",
        )
        self.export_loader_sub.pack(fill=tk.X, pady=(2, 8))
        self.export_progress = WindowsStyleProgressBar(el_inner, height=16)
        self.export_progress.pack(fill=tk.X)
        self.export_pct_label = tk.Label(
            el_inner, text="0%", bg=COLORS["surface"], fg=COLORS.get("text", "#1a202c"),
            font=("Segoe UI", 9, "bold"),
        )
        self.export_pct_label.pack(anchor="e", pady=(4, 0))
        # hidden until export starts

        # Notebook
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=12, pady=(4, 8))

        # --- Allocations tab ---
        alloc_frame = ttk.Frame(self.nb, padding=4)
        self.nb.add(alloc_frame, text="  Allocations  ")

        # Highlighted month banner (selected report month)
        self.month_banner = tk.Frame(alloc_frame, bg=COLORS["month_banner"], height=36)
        self.month_banner.pack(fill=tk.X, pady=(0, 6))
        self.month_banner.pack_propagate(False)
        self.month_banner_label = tk.Label(
            self.month_banner,
            text="Report Month: —",
            bg=COLORS["month_banner"],
            fg=COLORS["month_text"],
            font=("Segoe UI", 11, "bold"),
            anchor="w",
        )
        self.month_banner_label.pack(side=tk.LEFT, padx=12, pady=6)
        self.month_banner_count = tk.Label(
            self.month_banner,
            text="",
            bg=COLORS["month_banner"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
            anchor="e",
        )
        self.month_banner_count.pack(side=tk.RIGHT, padx=12, pady=6)

        # Grid host (loader is placed only over this area)
        self.grid_host = tk.Frame(alloc_frame, bg=COLORS["surface"])
        self.grid_host.pack(fill=tk.BOTH, expand=True)

        cols = ("lab", "day", "slot", "faculty", "batch", "module", "career", "students", "mod_start", "batch_start")
        self.tree = ttk.Treeview(self.grid_host, columns=cols, show="headings", selectmode="browse")
        headings = {
            "lab": "Lab", "day": "Days", "slot": "Time", "faculty": "Faculty",
            "batch": "Batch Code", "module": "Module",
            "career": "Career", "students": "Students",
            "mod_start": "Module Start", "batch_start": "Batch Start",
        }
        widths = {
            "lab": 55, "day": 55, "slot": 110, "faculty": 140,
            "batch": 130, "module": 100,
            "career": 75, "students": 65,
            "mod_start": 100, "batch_start": 100,
        }
        for c in cols:
            self.tree.heading(c, text=headings[c])
            self.tree.column(c, width=widths[c], anchor="center")

        scroll = ttk.Scrollbar(self.grid_host, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<Double-1>", lambda e: self._edit())

        # --- Dashboard tab ---
        dash_frame = ttk.Frame(self.nb, padding=8)
        self.nb.add(dash_frame, text="  Dashboard  ")

        self.dash_text = tk.Text(dash_frame, wrap="word", font=("Consolas", 10), state="disabled",
                              bg="#ffffff", fg="#1e293b", insertbackground="#1e293b",
                              relief="flat", padx=12, pady=10, highlightthickness=1,
                              highlightbackground="#cbd5e1", highlightcolor="#2563eb")
        self.dash_text.pack(fill=tk.BOTH, expand=True)

        # --- Faculty Management tab ---
        fac_frame = ttk.Frame(self.nb, padding=4)
        self.nb.add(fac_frame, text="  Faculty  ")

        fac_top = ttk.Frame(fac_frame)
        fac_top.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(
            fac_top,
            text="Faculty is stored in the database. Add new teachers here – they appear in allocations immediately.",
            foreground="#444",
        ).pack(side=tk.LEFT)
        ttk.Button(fac_top, text="＋ Add Faculty", command=self._add_faculty).pack(side=tk.RIGHT, padx=4)
        ttk.Button(fac_top, text="Edit", command=self._edit_faculty).pack(side=tk.RIGHT, padx=4)
        ttk.Button(fac_top, text="Deactivate", command=self._deactivate_faculty).pack(side=tk.RIGHT, padx=4)
        ttk.Button(fac_top, text="Reactivate", command=self._reactivate_faculty).pack(side=tk.RIGHT, padx=4)
        ttk.Button(fac_top, text="Delete", command=self._delete_faculty).pack(side=tk.RIGHT, padx=4)

        fac_cols = ("initials", "name", "status", "active", "created", "allocs")
        self.fac_tree = ttk.Treeview(fac_frame, columns=fac_cols, show="headings", selectmode="browse")
        fac_heads = {
            "initials": "Initials", "name": "Full Name", "status": "Type",
            "active": "Active", "created": "Created", "allocs": "Allocations",
        }
        fac_widths = {"initials": 80, "name": 220, "status": 90, "active": 70, "created": 140, "allocs": 100}
        for c in fac_cols:
            self.fac_tree.heading(c, text=fac_heads[c])
            self.fac_tree.column(c, width=fac_widths[c], anchor="center" if c != "name" else "w")
        fac_scroll = ttk.Scrollbar(fac_frame, orient=tk.VERTICAL, command=self.fac_tree.yview)
        self.fac_tree.configure(yscrollcommand=fac_scroll.set)
        self.fac_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        fac_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.fac_tree.bind("<Double-1>", lambda e: self._edit_faculty())

        # --- Careers / Modules tab ---
        car_frame = ttk.Frame(self.nb, padding=4)
        self.nb.add(car_frame, text="  Careers / Modules  ")

        # Careers (programs)
        car_top = ttk.Frame(car_frame)
        car_top.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(
            car_top,
            text="Careers / Programs (CPISM, DISM, HDSE I, ADSE II, STC, SP…)",
            foreground="#444",
        ).pack(side=tk.LEFT)
        ttk.Button(car_top, text="＋ Add Career", command=self._add_career).pack(side=tk.RIGHT, padx=2)
        ttk.Button(car_top, text="Edit Career", command=self._edit_career).pack(side=tk.RIGHT, padx=2)
        ttk.Button(car_top, text="Delete Career", command=self._delete_career).pack(side=tk.RIGHT, padx=2)

        car_cols = ("code", "name", "type", "sort")
        self.car_tree = ttk.Treeview(car_frame, columns=car_cols, show="headings", selectmode="browse", height=8)
        for c, h, w in (
            ("code", "Code", 100), ("name", "Name", 180), ("type", "Type", 80), ("sort", "Order", 60)
        ):
            self.car_tree.heading(c, text=h)
            self.car_tree.column(c, width=w, anchor="center")
        self.car_tree.pack(fill=tk.BOTH, expand=True, pady=(0, 6))
        self.car_tree.bind("<Double-1>", lambda e: self._edit_career())

        # Modules (subjects / short-course topics)
        mod_top = ttk.Frame(car_frame)
        mod_top.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(
            mod_top,
            text="Modules / Subjects (FSA, MS Office, Java, Web Designing…)",
            foreground="#444",
        ).pack(side=tk.LEFT)
        ttk.Button(mod_top, text="＋ Add Module", command=self._add_module).pack(side=tk.RIGHT, padx=2)
        ttk.Button(mod_top, text="Edit Module", command=self._edit_module).pack(side=tk.RIGHT, padx=2)
        ttk.Button(mod_top, text="Delete Module", command=self._delete_module).pack(side=tk.RIGHT, padx=2)

        mod_cols = ("name", "category")
        self.mod_tree = ttk.Treeview(car_frame, columns=mod_cols, show="headings", selectmode="browse", height=8)
        self.mod_tree.heading("name", text="Module Name")
        self.mod_tree.heading("category", text="Category")
        self.mod_tree.column("name", width=280, anchor="w")
        self.mod_tree.column("category", width=100, anchor="center")
        self.mod_tree.pack(fill=tk.BOTH, expand=True)
        self.mod_tree.bind("<Double-1>", lambda e: self._edit_module())

        # --- Excel / BSR Preview tab ---
        preview_frame = ttk.Frame(self.nb, padding=2)
        self.nb.add(preview_frame, text="  Excel Preview  ")
        self.excel_preview = ExcelPreviewPanel(preview_frame)
        self.excel_preview.pack(fill=tk.BOTH, expand=True)
        # Auto-load latest BSR in Output/ if present
        try:
            self._try_load_latest_preview()
        except Exception:
            pass

        # DB path info bar
        self.db_info = tk.StringVar()
        ttk.Label(self, textvariable=self.db_info, style="Muted.TLabel").pack(fill=tk.X, padx=12, pady=(2, 0))

        # Status bar
        self.status = tk.StringVar(value="Ready")
        ttk.Label(self, textvariable=self.status, style="Status.TLabel", anchor="w").pack(fill=tk.X, side=tk.BOTTOM)

    def refresh(self):
        # Allocations list — STRICTLY selected month only
        for i in self.tree.get_children():
            self.tree.delete(i)
        ym = (self.bsr_month_var.get() if hasattr(self, "bsr_month_var") else "") or ""
        ym = ym.strip()
        if not ym:
            ym = datetime.now().strftime("%Y-%m")
            if hasattr(self, "bsr_month_var"):
                self.bsr_month_var.set(ym)
        # Never pass None (that would return every month)
        allocs = list_allocations(active_only=True, report_month=ym)
        for a in allocs:
            self.tree.insert("", tk.END, iid=str(a.id), values=(
                a.lab_code, a.day_group_code, a.time_slot_label,
                f"{a.faculty_initials} ({a.faculty_name})",
                a.batch_code or "", a.module_name or "",
                a.career_code, a.students,
                a.module_start_date or "",
                a.actual_start_date or "",
            ))

        # Dashboard text — same month scope
        dash = full_dashboard(report_month=ym)
        lines = []
        lines.append(f"TOTAL BATCHES : {dash['total_batches']}")
        lines.append(f"TOTAL STUDENTS: {dash['total_students']}")
        lines.append("")
        lines.append("─── CAREER SUMMARY ───")
        for c in dash["career_summary"]:
            if c["code"] == "Total":
                lines.append(f"  {c['code']:<12}  Batches={c['batches']:<4}  Students={c['students']}")
            else:
                lines.append(f"  {c['code']:<12}  Batches={c['batches']:<4}  Students={c['students']:<5}  ({c['pct']*100:.1f}%)")
        lines.append("")
        lines.append("─── MWF / TTS TOTALS ───")
        for dg, t in dash["day_group_totals"].items():
            lines.append(f"  {dg}:  {t['batches']} batches,  {t['students']} students")
        lines.append("")
        lines.append("─── FACULTY WORKLOAD ───")
        for f in dash["faculty_workload"]:
            if f["load"] == 0:
                continue
            lines.append(f"  {f['initials']:<4} {f['status']}  {f['full_name']:<20}  Load={f['load']}  Students={f['students']}  {f['timings']}")
        lines.append("")
        stc = dash["stc_sp"]
        lines.append(f"STC: {stc['stc_batches']} batches / {stc['stc_students']} students")
        lines.append(f"SP : {stc['sp_batches']} batches / {stc['sp_students']} students")
        lines.append("")
        lines.append("─── PAST MONTHLY BSRs ───")
        try:
            for row in list_monthly_bsr_metrics()[:12]:
                lines.append(
                    f"  {row.get('year_month')}:  Students={row.get('total_students', 0)}  "
                    f"STC {row.get('stc_beg', 0)}→{row.get('stc_end', 0)}  "
                    f"SP {row.get('sp_beg', 0)}→{row.get('sp_end', 0)}"
                )
        except Exception:
            lines.append("  (no monthly snapshots yet)")

        self.dash_text.configure(state="normal")
        self.dash_text.delete("1.0", tk.END)
        self.dash_text.insert("1.0", "\n".join(lines))
        self.dash_text.configure(state="disabled")

        # Faculty list (include inactive so admin can reactivate)
        for i in self.fac_tree.get_children():
            self.fac_tree.delete(i)
        for fac in list_faculty(active_only=False):
            n_alloc = count_allocations_for_faculty(fac.id, active_only=False)
            self.fac_tree.insert(
                "", tk.END, iid=str(fac.id),
                values=(
                    fac.initials,
                    fac.full_name,
                    "Permanent" if fac.status.upper() == "P" else "Visiting",
                    "Yes" if fac.active else "No",
                    fac.created_at or "",
                    n_alloc,
                ),
            )

        # Careers list
        if hasattr(self, "car_tree"):
            for i in self.car_tree.get_children():
                self.car_tree.delete(i)
            for car in list_careers():
                typ = "STC" if car.is_stc else ("SP" if car.is_sp else "Career")
                self.car_tree.insert(
                    "", tk.END, iid=str(car.id),
                    values=(car.code, car.name or car.code, typ, car.sort_order),
                )

        # Modules list (drives Module/Subject dropdown in allocation form)
        if hasattr(self, "mod_tree"):
            for i in self.mod_tree.get_children():
                self.mod_tree.delete(i)
            try:
                for mod in list_modules(active_only=False):
                    cat = (mod.category or "CAREER").upper()
                    if cat == "SUBJECT":
                        cat = "CAREER"
                    self.mod_tree.insert(
                        "", tk.END, iid=f"m{mod.id}",
                        values=(mod.name, cat),
                    )
            except Exception:
                pass

        try:
            self.db_info.set(f"Database file: {get_db_path()}")
        except Exception:
            pass

        if hasattr(self, "month_banner_label"):
            self.month_banner_label.configure(text=f"Report Month:  {ym}")
            self.month_banner_count.configure(text=f"{len(allocs)} allocation(s)")
        self.status.set(f"Showing {ym} only  ·  {len(allocs)} allocation(s)")

    def _add(self):
        AllocationForm(self, on_save=self.refresh)

    def _edit(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Please select an allocation first.")
            return
        AllocationForm(self, allocation_id=int(sel[0]), on_save=self.refresh)

    def _deactivate(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Please select an allocation first.")
            return
        if messagebox.askyesno("Confirm", "Deactivate this allocation?\n(It will be hidden but kept in the database.)"):
            deactivate_allocation(int(sel[0]))
            self.refresh()

    def _delete_one(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Please select an allocation first.")
            return
        if messagebox.askyesno(
            "Delete permanently?",
            "Permanently delete this allocation?\nThis cannot be undone.",
            icon="warning",
        ):
            delete_allocation(int(sel[0]))
            self.refresh()
            self.status.set("Allocation deleted")

    def _clear_all(self):
        n = len(list_allocations(active_only=False))
        if n == 0:
            messagebox.showinfo("Nothing to clear", "There are no allocations in the database.")
            return
        ok = messagebox.askyesno(
            "Clear ALL allocations?",
            f"This will permanently delete all {n} allocation(s).\n\n"
            "Labs, faculty, careers and time slots are kept.\n"
            "You can then enter fresh data manually.\n\n"
            "This cannot be undone. Continue?",
            icon="warning",
        )
        if not ok:
            return
        # Second confirmation for safety
        ok2 = messagebox.askyesno(
            "Final confirmation",
            "Are you sure? All sample / existing schedule data will be wiped.",
            icon="warning",
        )
        if not ok2:
            return
        removed = clear_all_allocations()
        self.refresh()
        messagebox.showinfo("Cleared", f"Removed {removed} allocation(s).\nYou can now add fresh data.")
        self.status.set(f"Cleared {removed} allocation(s) — ready for fresh data")




    def _add_module(self):
        ModuleDialog(self, on_saved=self.refresh)

    def _edit_module(self):
        sel = self.mod_tree.selection()
        if not sel:
            messagebox.showinfo("Edit Module", "Select a module first.")
            return
        mid = int(str(sel[0]).lstrip("m"))
        ModuleDialog(self, module_id=mid, on_saved=self.refresh)

    def _delete_module(self):
        sel = self.mod_tree.selection()
        if not sel:
            messagebox.showinfo("Delete", "Select a module first.")
            return
        mid = int(str(sel[0]).lstrip("m"))
        if not messagebox.askyesno("Delete Module", "Delete this module?"):
            return
        try:
            delete_module(mid)
            self.refresh()
            self.status.set("Module deleted.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _on_month_changed(self, _event=None):
        """Load only the selected month's BSR data with a visible loader."""
        ym = (self.bsr_month_var.get() or "").strip()
        if not ym:
            return
        # Clear grid immediately so mixed months never stay visible
        try:
            for i in self.tree.get_children():
                self.tree.delete(i)
        except Exception:
            pass
        self._show_month_loader(ym)

        def work():
            import time
            err = None
            count = 0
            t0 = time.time()
            try:
                from models import list_allocations
                from calculations import full_dashboard
                # Force month-scoped fetch
                rows = list_allocations(active_only=True, report_month=ym)
                count = len(rows)
                full_dashboard(report_month=ym)
            except Exception as e:
                err = e
            # Minimum delay so the loader is noticeable
            elapsed = time.time() - t0
            if elapsed < 0.7:
                time.sleep(0.7 - elapsed)

            def done():
                self._hide_month_loader()
                if err:
                    self.status.set(f"Failed to load {ym}: {err}")
                    messagebox.showerror("Load failed", str(err))
                    return
                self.refresh()
                self.status.set(f"Showing {ym} only  ·  {count} allocation(s)")
            self.after(0, done)

        import threading
        threading.Thread(target=work, daemon=True).start()

    def _show_month_loader(self, ym: str):
        """Percentage loader only over the allocations data grid."""
        self._hide_month_loader()
        self.status.set(f"Loading {ym} BSR… 0%")
        if hasattr(self, "month_banner_label"):
            self.month_banner_label.configure(text=f"Report Month:  {ym}  ·  Loading…")
            self.month_banner_count.configure(text="")

        host = getattr(self, "grid_host", None) or self
        overlay = tk.Frame(host, bg="#f5f6f8")
        overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        card = tk.Frame(overlay, bg="#ffffff", padx=28, pady=22,
                        highlightbackground="#e5e7eb", highlightthickness=1)
        card.place(relx=0.5, rely=0.42, anchor="center")
        tk.Label(
            card, text=f"Loading {ym}",
            bg="#ffffff", fg="#1f2937",
            font=("Segoe UI", 12, "bold"),
        ).pack(pady=(0, 4))
        pct_lbl = tk.Label(
            card, text="0%",
            bg="#ffffff", fg="#1f2937",
            font=("Segoe UI", 11, "bold"),
        )
        pct_lbl.pack(pady=(2, 4))
        tk.Label(
            card, text="Please wait while this month's data is prepared…",
            bg="#ffffff", fg="#6b7280",
            font=("Segoe UI", 9),
        ).pack()
        bar = WindowsStyleProgressBar(card, height=14)
        bar.pack(pady=(12, 0), fill=tk.X)
        bar.set(0)
        bar.start()
        self._month_loader = overlay
        self._month_loader_bar = bar
        self._month_loader_pct = pct_lbl
        self._month_loader_pct_val = 0

        def _tick():
            if getattr(self, "_month_loader", None) is None:
                return
            cur = getattr(self, "_month_loader_pct_val", 0)
            if cur < 90:
                cur = min(90, cur + 7)
                self._month_loader_pct_val = cur
                try:
                    bar.set(cur)
                    pct_lbl.configure(text=f"{cur}%")
                    self.status.set(f"Loading {ym} BSR… {cur}%")
                except Exception:
                    pass
                self.after(120, _tick)

        self.after(80, _tick)
        self.update_idletasks()

    def _hide_month_loader(self):
        try:
            self.progress.stop()
            self.progress.pack_forget()
            self.progress.set(0)
        except Exception:
            pass
        bar = getattr(self, "_month_loader_bar", None)
        pct_lbl = getattr(self, "_month_loader_pct", None)
        if bar is not None:
            try:
                bar.set(100)
                if pct_lbl is not None:
                    pct_lbl.configure(text="100%")
                bar.stop()
            except Exception:
                pass
        ov = getattr(self, "_month_loader", None)
        if ov is not None:
            try:
                ov.place_forget()
                ov.destroy()
            except Exception:
                pass
            self._month_loader = None
        self._month_loader_bar = None
        self._month_loader_pct = None
        self._month_loader_pct_val = 0

    def _month_choices(self):
        """Months that have BSR data (newest first). Current month always listed for new entry."""
        try:
            available = list_available_report_months()
        except Exception:
            available = []
        with_data = []
        for m in available:
            try:
                if list_allocations(active_only=True, report_month=m):
                    with_data.append(m)
            except Exception:
                pass
        cur = datetime.now().strftime("%Y-%m")
        out = list(with_data)
        if cur not in out:
            out.append(cur)
        # also keep other available empty months at end
        for m in available:
            if m not in out:
                out.append(m)
        return out or [cur]


    def _add_career(self):

        CareerDialog(self, on_saved=self.refresh)

    def _edit_career(self):
        sel = self.car_tree.selection()
        if not sel:
            messagebox.showinfo("Edit Career", "Select a career/course first.")
            return
        CareerDialog(self, career_id=int(sel[0]), on_saved=self.refresh)

    def _delete_career(self):
        sel = self.car_tree.selection()
        if not sel:
            messagebox.showinfo("Delete", "Select a career/course first.")
            return
        cid = int(sel[0])
        n = count_allocations_for_career(cid)
        if n > 0:
            messagebox.showerror(
                "Cannot Delete",
                f"This career is used by {n} active allocation(s).\n"
                "Reassign or remove those allocations first.",
            )
            return
        if not messagebox.askyesno("Delete Career", "Delete this career/course permanently?"):
            return
        try:
            delete_career(cid)
            self.refresh()
            self.status.set("Career deleted.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _add_faculty(self):
        FacultyForm(self, on_save=self.refresh)

    def _edit_faculty(self):
        sel = self.fac_tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Please select a faculty member first.")
            return
        FacultyForm(self, faculty_id=int(sel[0]), on_save=self.refresh)

    def _deactivate_faculty(self):
        sel = self.fac_tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Please select a faculty member first.")
            return
        fid = int(sel[0])
        fac = get_faculty(fid)
        if not fac:
            return
        if not fac.active:
            messagebox.showinfo("Already inactive", f"{fac.full_name} is already inactive.")
            return
        n = count_allocations_for_faculty(fid, active_only=False)
        msg = (
            f"Deactivate {fac.initials} – {fac.full_name}?\n\n"
            "They will no longer appear in new allocation dropdowns.\n"
            "Existing allocation history is kept."
        )
        if n:
            msg += f"\n\n({n} allocation record(s) linked – history preserved.)"
        if messagebox.askyesno("Deactivate Faculty", msg):
            set_faculty_active(fid, False)
            self.refresh()
            self.status.set(f"Deactivated faculty {fac.initials}")

    def _reactivate_faculty(self):
        sel = self.fac_tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Please select a faculty member first.")
            return
        fid = int(sel[0])
        fac = get_faculty(fid)
        if not fac:
            return
        if fac.active:
            messagebox.showinfo("Already active", f"{fac.full_name} is already active.")
            return
        set_faculty_active(fid, True)
        self.refresh()
        self.status.set(f"Reactivated faculty {fac.initials}")

    def _delete_faculty(self):
        sel = self.fac_tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Please select a faculty member first.")
            return
        fid = int(sel[0])
        fac = get_faculty(fid)
        if not fac:
            return
        n = count_allocations_for_faculty(fid, active_only=False)
        if n > 0:
            messagebox.showwarning(
                "Cannot delete",
                f"This faculty has {n} allocation record(s).\n\n"
                "Use Deactivate instead to hide them while keeping history.",
            )
            return
        if not messagebox.askyesno(
            "Delete permanently?",
            f"Permanently delete {fac.initials} – {fac.full_name}?\n"
            "This faculty has no allocation history.\nThis cannot be undone.",
            icon="warning",
        ):
            return
        ok, msg = delete_faculty(fid)
        if not ok:
            messagebox.showwarning("Cannot delete", msg)
            return
        self.refresh()
        self.status.set(f"Deleted faculty {fac.initials}")

    def _set_export_progress(self, pct: int, message: str = ""):
        """Update determinate progress bars and labels (0–100). Safe from worker thread via after()."""
        pct = max(0, min(100, int(pct)))

        def _apply():
            try:
                self.export_progress.set(pct)
                self.progress.set(pct)
                self.export_pct_label.configure(text=f"{pct}%")
                if message:
                    self.export_loader_sub.configure(text=message)
                    self.status.set(message)
            except Exception:
                pass

        try:
            self.after(0, _apply)
        except Exception:
            _apply()

    def _show_export_loader(self, year_month: str):
        self.export_loader_title.configure(text=f"Generating BSR for {year_month}")
        self.export_loader_sub.configure(text="Starting… 0%")
        try:
            self.export_progress.set(0)
            self.export_pct_label.configure(text="0%")
            self.progress.set(0)
        except Exception:
            pass
        self.export_loader.pack(fill=tk.X, padx=12, pady=(0, 8), before=self.nb if hasattr(self, "nb") else None)
        try:
            self.progress.pack(fill=tk.X, side=tk.BOTTOM, padx=8, pady=2)
        except Exception:
            pass
        try:
            self.export_progress.start()
            self.progress.start()
        except Exception:
            pass
        self.update_idletasks()

    def _hide_export_loader(self):
        try:
            self.export_progress.stop()
            self.export_progress.set(0)
            self.export_loader.pack_forget()
        except Exception:
            pass
        try:
            self.progress.stop()
            self.progress.set(0)
            self.progress.pack_forget()
        except Exception:
            pass

    def _export(self):
        if self._exporting:
            return
        year_month = (self.bsr_month_var.get() if hasattr(self, "bsr_month_var") else None) or datetime.now().strftime("%Y-%m")
        self._exporting = True
        try:
            self.export_btn.configure(state="disabled")
        except Exception:
            pass
        self._show_export_loader(year_month)
        self.status.set(f"Generating BSR for {year_month}… 0%")

        def worker():
            import time
            err = None
            out = None
            t0 = time.time()
            try:
                OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
                out_path = OUTPUT_DIR / f"BSR_{year_month}.xlsx"

                def on_progress(pct, msg=""):
                    self._set_export_progress(pct, msg or f"Generating BSR… {pct}%")

                out = export_workbook(
                    out_path,
                    year_month=year_month,
                    progress_callback=on_progress,
                )
                self._set_export_progress(100, "Finalising…")
            except Exception as e:
                err = e
            elapsed = time.time() - t0
            if elapsed < 0.5:
                time.sleep(0.5 - elapsed)

            def done():
                self._hide_export_loader()
                try:
                    self.export_btn.configure(state="normal")
                except Exception:
                    pass
                self._exporting = False
                if err is not None:
                    messagebox.showerror("BSR generation failed", str(err))
                    self.status.set("BSR generation failed")
                    return
                try:
                    self.open_folder_btn.pack(side=tk.LEFT, padx=4)
                except Exception:
                    pass
                # Load into Excel Preview tab
                try:
                    if out and hasattr(self, "excel_preview"):
                        self.excel_preview.load_file(out)
                        # Switch to preview tab (last tab)
                        try:
                            for i in range(self.nb.index("end")):
                                if "Preview" in str(self.nb.tab(i, "text")):
                                    self.nb.select(i)
                                    break
                        except Exception:
                            pass
                except Exception:
                    pass
                messagebox.showinfo(
                    "Success",
                    f"Monthly BSR generated for {year_month}.\n\n"
                    f"Saved to:\n{out}\n\n"
                    f"Opened in the Excel Preview tab for review.\n"
                    f"Previous months are kept separately and are not overwritten.",
                )
                self.status.set(f"BSR {year_month} → {out}")
                self.refresh()

            self.after(0, done)

        threading.Thread(target=worker, daemon=True).start()


    def _try_load_latest_preview(self):
        """If Output/ has BSR_*.xlsx files, load the newest into the preview tab."""
        try:
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            files = sorted(
                OUTPUT_DIR.glob("BSR_*.xlsx"),
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )
            if files and hasattr(self, "excel_preview"):
                self.excel_preview.load_file(files[0])
        except Exception:
            pass

    def _open_output_folder(self):
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        path = str(OUTPUT_DIR.resolve())
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", path])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("Open folder", f"Could not open folder:\n{path}\n\n{e}")


def main():
    app = MainApp()
    app.mainloop()


if __name__ == "__main__":
    main()
