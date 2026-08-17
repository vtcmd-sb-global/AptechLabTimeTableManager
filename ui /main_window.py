#!/usr/bin/env python3
"""
Main application window – Aptech Lab Timetable Manager
"""

from __future__ import annotations

import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path

# Ensure project root on path
ROOT = Path(__file__).parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from database import init_db
from models import (
    list_allocations, list_faculty, list_labs, list_careers,
    list_time_slots, list_day_groups,
    add_allocation, update_allocation, deactivate_allocation,
    delete_allocation, clear_all_allocations,
    find_conflicts, find_faculty_conflicts, check_capacity, get_allocation,
    derive_course_from_batch,
    add_faculty, add_lab, add_career,
)
from calculations import full_dashboard, faculty_workload, career_summary, day_group_totals
from export_excel import export_workbook


class AllocationForm(tk.Toplevel):
    """Modal form to add / edit a single allocation."""

    def __init__(self, master, allocation_id=None, on_save=None):
        super().__init__(master)
        self.allocation_id = allocation_id
        self.on_save = on_save
        self.title("Edit Allocation" if allocation_id else "Add New Allocation")
        self.geometry("520x520")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

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
        ttk.Label(f, text="Faculty *").grid(row=0, column=0, sticky="w", **pad)
        self.fac_var = tk.StringVar()
        fac_values = [f"{x.initials} – {x.full_name}" for x in self.faculty]
        self.fac_combo = ttk.Combobox(f, textvariable=self.fac_var, values=fac_values, width=40, state="readonly")
        self.fac_combo.grid(row=0, column=1, columnspan=2, sticky="ew", **pad)

        # Career
        ttk.Label(f, text="Career / Program *").grid(row=1, column=0, sticky="w", **pad)
        self.career_var = tk.StringVar()
        self.career_combo = ttk.Combobox(
            f, textvariable=self.career_var,
            values=[c.code for c in self.careers], width=40, state="readonly"
        )
        self.career_combo.grid(row=1, column=1, columnspan=2, sticky="ew", **pad)

        # Lab
        ttk.Label(f, text="Lab *").grid(row=2, column=0, sticky="w", **pad)
        self.lab_var = tk.StringVar()
        self.lab_combo = ttk.Combobox(
            f, textvariable=self.lab_var,
            values=[f"{l.lab_code} ({l.pcs} PCs)" for l in self.labs], width=40, state="readonly"
        )
        self.lab_combo.grid(row=2, column=1, columnspan=2, sticky="ew", **pad)

        # Time slot
        ttk.Label(f, text="Time Slot *").grid(row=3, column=0, sticky="w", **pad)
        self.slot_var = tk.StringVar()
        self.slot_combo = ttk.Combobox(
            f, textvariable=self.slot_var,
            values=[s.label for s in self.slots], width=40, state="readonly"
        )
        self.slot_combo.grid(row=3, column=1, columnspan=2, sticky="ew", **pad)

        # Day group
        ttk.Label(f, text="Day Group *").grid(row=4, column=0, sticky="w", **pad)
        self.day_var = tk.StringVar()
        self.day_combo = ttk.Combobox(
            f, textvariable=self.day_var,
            values=[d.code for d in self.days], width=40, state="readonly"
        )
        self.day_combo.grid(row=4, column=1, columnspan=2, sticky="ew", **pad)

        # Batch code (course/series is derived from this, e.g. AI-202412C2 → AI)
        ttk.Label(f, text="Batch Code").grid(row=5, column=0, sticky="w", **pad)
        self.batch_var = tk.StringVar()
        ttk.Entry(f, textvariable=self.batch_var, width=42).grid(row=5, column=1, columnspan=2, sticky="ew", **pad)

        # Module
        ttk.Label(f, text="Module / Subject").grid(row=6, column=0, sticky="w", **pad)
        self.module_var = tk.StringVar()
        ttk.Entry(f, textvariable=self.module_var, width=42).grid(row=6, column=1, columnspan=2, sticky="ew", **pad)

        # Students
        ttk.Label(f, text="No. of Students *").grid(row=7, column=0, sticky="w", **pad)
        self.students_var = tk.StringVar(value="0")
        ttk.Entry(f, textvariable=self.students_var, width=12).grid(row=7, column=1, sticky="w", **pad)

        # Dates
        ttk.Label(f, text="Module Start Date").grid(row=8, column=0, sticky="w", **pad)
        self.mod_start_var = tk.StringVar()
        ttk.Entry(f, textvariable=self.mod_start_var, width=16).grid(row=8, column=1, sticky="w", **pad)
        ttk.Label(f, text="(YYYY-MM-DD)").grid(row=8, column=2, sticky="w")

        ttk.Label(f, text="Actual Start Date").grid(row=9, column=0, sticky="w", **pad)
        self.act_start_var = tk.StringVar()
        ttk.Entry(f, textvariable=self.act_start_var, width=16).grid(row=9, column=1, sticky="w", **pad)

        # Admission open
        self.adm_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="Admission Open", variable=self.adm_var).grid(
            row=10, column=1, sticky="w", **pad
        )

        # Notes
        ttk.Label(f, text="Notes").grid(row=11, column=0, sticky="nw", **pad)
        self.notes_txt = tk.Text(f, height=3, width=40)
        self.notes_txt.grid(row=11, column=1, columnspan=2, sticky="ew", **pad)

        # Buttons
        btn_f = ttk.Frame(f)
        btn_f.grid(row=12, column=0, columnspan=3, pady=16)
        ttk.Button(btn_f, text="Save", command=self._save, width=12).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_f, text="Cancel", command=self.destroy, width=12).pack(side=tk.LEFT, padx=6)

    def _load(self, aid: int):
        a = get_allocation(aid)
        if not a:
            return
        self.fac_var.set(f"{a.faculty_initials} – {a.faculty_name}")
        self.career_var.set(a.career_code)
        self.lab_var.set(f"{a.lab_code} ({a.lab_pcs} PCs)")
        self.slot_var.set(a.time_slot_label)
        self.day_var.set(a.day_group_code)
        self.batch_var.set(a.batch_code or "")
        self.module_var.set(a.module_name or "")
        self.students_var.set(str(a.students))
        self.mod_start_var.set(a.module_start_date or "")
        self.act_start_var.set(a.actual_start_date or "")
        self.adm_var.set(a.is_admission_open)
        if a.notes:
            self.notes_txt.insert("1.0", a.notes)

    def _resolve_ids(self):
        fac_txt = self.fac_var.get()
        faculty_id = next((f.id for f in self.faculty if f"{f.initials} – {f.full_name}" == fac_txt), None)
        career_id = next((c.id for c in self.careers if c.code == self.career_var.get()), None)
        lab_txt = self.lab_var.get()
        lab_id = next((l.id for l in self.labs if f"{l.lab_code} ({l.pcs} PCs)" == lab_txt), None)
        slot_id = next((s.id for s in self.slots if s.label == self.slot_var.get()), None)
        day_id = next((d.id for d in self.days if d.code == self.day_var.get()), None)
        return faculty_id, career_id, lab_id, slot_id, day_id

    def _save(self):
        faculty_id, career_id, lab_id, slot_id, day_id = self._resolve_ids()
        if not all([faculty_id, career_id, lab_id, slot_id, day_id]):
            messagebox.showerror("Missing data", "Please fill all required (*) fields.", parent=self)
            return
        try:
            students = int(self.students_var.get() or 0)
        except ValueError:
            messagebox.showerror("Invalid", "Students must be a number.", parent=self)
            return

        batch_code = self.batch_var.get().strip() or None
        # Course/series is derived from batch code (matches Excel template style)
        course_title = derive_course_from_batch(batch_code)

        # Capacity check
        ok, pcs = check_capacity(lab_id, students)
        if not ok:
            if not messagebox.askyesno(
                "Capacity Warning",
                f"Students ({students}) exceed lab capacity ({pcs} PCs).\nSave anyway?",
                parent=self,
            ):
                return

        # Lab occupancy conflict (same lab + time + day)
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

        # Faculty double-booking warning (same teacher at same time + day)
        # This does NOT block or overwrite – a teacher can still have multiple careers
        # in different slots; only same-time overlap is flagged.
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
            "faculty_id": faculty_id,
            "career_id": career_id,
            "lab_id": lab_id,
            "time_slot_id": slot_id,
            "day_group_id": day_id,
            "batch_code": batch_code,
            "course_title": course_title,
            "module_name": self.module_var.get().strip() or None,
            "students": students,
            "module_start_date": self.mod_start_var.get().strip() or None,
            "actual_start_date": self.act_start_var.get().strip() or None,
            "notes": self.notes_txt.get("1.0", "end").strip() or None,
            "is_admission_open": self.adm_var.get(),
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


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Aptech Lab Timetable & Faculty Workload Manager")
        self.geometry("1100x700")
        self.minsize(900, 600)

        init_db()
        self._build_ui()
        self.refresh()

    def _build_ui(self):
        # Top bar
        top = ttk.Frame(self, padding=8)
        top.pack(fill=tk.X)

        ttk.Label(
            top, text="Aptech Lab Timetable Manager",
            font=("Segoe UI", 14, "bold")
        ).pack(side=tk.LEFT)

        ttk.Button(top, text="＋ Add Allocation", command=self._add).pack(side=tk.RIGHT, padx=4)
        ttk.Button(top, text="Edit Selected", command=self._edit).pack(side=tk.RIGHT, padx=4)
        ttk.Button(top, text="Deactivate", command=self._deactivate).pack(side=tk.RIGHT, padx=4)
        ttk.Button(top, text="Delete Selected", command=self._delete_one).pack(side=tk.RIGHT, padx=4)
        ttk.Button(top, text="Clear All Data", command=self._clear_all).pack(side=tk.RIGHT, padx=8)
        ttk.Button(top, text="Export Excel Report", command=self._export).pack(side=tk.RIGHT, padx=8)
        ttk.Button(top, text="Refresh", command=self.refresh).pack(side=tk.RIGHT, padx=4)

        # Notebook
        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        # --- Allocations tab ---
        alloc_frame = ttk.Frame(nb, padding=4)
        nb.add(alloc_frame, text="  Allocations  ")

        cols = ("lab", "day", "slot", "faculty", "batch", "module", "career", "students", "start")
        self.tree = ttk.Treeview(alloc_frame, columns=cols, show="headings", selectmode="browse")
        headings = {
            "lab": "Lab", "day": "Days", "slot": "Time", "faculty": "Faculty",
            "batch": "Batch Code", "module": "Module",
            "career": "Career", "students": "Students", "start": "Start Date",
        }
        widths = {
            "lab": 60, "day": 60, "slot": 120, "faculty": 150,
            "batch": 140, "module": 110,
            "career": 80, "students": 70, "start": 100,
        }
        for c in cols:
            self.tree.heading(c, text=headings[c])
            self.tree.column(c, width=widths[c], anchor="center")

        scroll = ttk.Scrollbar(alloc_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<Double-1>", lambda e: self._edit())

        # --- Dashboard tab ---
        dash_frame = ttk.Frame(nb, padding=8)
        nb.add(dash_frame, text="  Dashboard  ")

        self.dash_text = tk.Text(dash_frame, wrap="word", font=("Consolas", 10), state="disabled", bg="#f7fafc")
        self.dash_text.pack(fill=tk.BOTH, expand=True)

        # Status bar
        self.status = tk.StringVar(value="Ready")
        ttk.Label(self, textvariable=self.status, relief=tk.SUNKEN, anchor="w").pack(fill=tk.X, side=tk.BOTTOM)

    def refresh(self):
        # Allocations list
        for i in self.tree.get_children():
            self.tree.delete(i)
        allocs = list_allocations(active_only=True)
        for a in allocs:
            self.tree.insert("", tk.END, iid=str(a.id), values=(
                a.lab_code, a.day_group_code, a.time_slot_label,
                f"{a.faculty_initials} ({a.faculty_name})",
                a.batch_code or "", a.module_name or "",
                a.career_code, a.students, a.module_start_date or "",
            ))

        # Dashboard text
        dash = full_dashboard()
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

        self.dash_text.configure(state="normal")
        self.dash_text.delete("1.0", tk.END)
        self.dash_text.insert("1.0", "\n".join(lines))
        self.dash_text.configure(state="disabled")

        self.status.set(f"Loaded {len(allocs)} active allocation(s)")

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

    def _export(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"Aptech_LabStatus_{__import__('datetime').datetime.now().strftime('%Y-%m-%d')}.xlsx",
        )
        if not path:
            return
        try:
            out = export_workbook(Path(path))
            messagebox.showinfo("Exported", f"Report saved to:\n{out}")
            self.status.set(f"Exported → {out}")
        except Exception as e:
            messagebox.showerror("Export failed", str(e))


def main():
    app = MainApp()
    app.mainloop()


if __name__ == "__main__":
    main()
