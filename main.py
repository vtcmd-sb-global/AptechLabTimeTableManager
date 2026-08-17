#!/usr/bin/env python3
"""
Aptech Lab Timetable & Faculty Workload Manager
Entry point.
"""

from database import init_db
from ui.main_window import main

if __name__ == "__main__":
    init_db()
    main()
