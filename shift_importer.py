# shift_importer.py
"""
Tailored importer for the Staff shift details .xlsx you uploaded.

What it does:
- Reads the sheet layout you provided: department blocks, then employee rows.
- For each employee row (empcode in col 0), reads name (col1) and shift cell (col2).
- Parses shift string(s) like "09:00 to 18:00" or "07:00-13:00;15:00-19:00".
- Creates a reusable shift pattern in shift_master (shift_code).
- Upserts employee with default_shift pointing to that shift_code.
- Optionally bulk-assigns that shift for a given year_month into employee_shift_assignments.

Usage:
    python shift_importer.py /path/to/Staff\ shift\ details\ .xlsx 2025-07

If year_month (YYYY-MM) provided, the importer will bulk-assign the parsed default shift
to every day of that month for each employee (so compute pipeline can use assignments).
"""

import pandas as pd
import re
import sys
from mongo_helpers import get_db, upsert_shift_master, upsert_employee, bulk_assign_shift_for_month
from datetime import datetime
from bson import ObjectId

def parse_shift_cell(cell_value):
    """
    Parse a shift cell and return list of intervals like [{'start':'HH:MM','end':'HH:MM'}, ...]
    - Supports separators ';' ',' '/'
    - Recognizes patterns: 'HH:MM to HH:MM', 'HH:MM-HH:MM', 'HH:MM – HH:MM'
    - Empty / OFF / WO / -- => returns []
    """
    if cell_value is None:
        return []
    s = str(cell_value).strip()
    if s == "" or s.upper() in ("OFF", "WO", "W/O", "HOL", "--", "-"):
        return []
    parts = re.split(r'[;,/]', s)
    intervals = []
    for p in parts:
        p = p.strip()
        # match patterns like 14:00 to 23:00 or 14:00-23:00
        m = re.search(r'(\d{1,2}:\d{2})\s*(?:to|[-–])\s*(\d{1,2}:\d{2})', p, flags=re.IGNORECASE)
        if m:
            start = m.group(1)
            end = m.group(2)
            intervals.append({"start": start, "end": end})
    return intervals

def make_shift_code(intervals):
    """
    Create a deterministic shift_code from intervals list so same pattern reuses same code.
    Example: [{'start':'09:00','end':'18:00'}] -> SH_09_00-18_00
    For split: SH_09_00-13_00__15_00-19_00
    """
    if not intervals:
        return "OFF"
    parts = []
    for it in intervals:
        s = it["start"].replace(":", "_")
        e = it["end"].replace(":", "_")
        parts.append(f"{s}-{e}")
    code = "SH_" + "__".join(parts)
    # replace any illegal chars
    code = re.sub(r'[^A-Za-z0-9_\-]', '_', code)
    return code

def importer(file_path, sheet_name=None, year_month=None):
    df = pd.read_excel(file_path, header=None, sheet_name=sheet_name, dtype=str)
    df = df.fillna("").astype(str)
    nrows, ncols = df.shape
    db = get_db()

    current_dept = None
    created_shifts = {}
    stats = {"employees":0, "shifts_created":0, "assigned_month":0}

    for i in range(nrows):
        row0 = df.iat[i,0].strip() if 0 < ncols else ""
        # detect Dept. Name rows: cell contains 'Dept' or 'Dept.' or 'Dept. Name'
        if row0.lower().startswith("dept"):
            # department name usually in column 1
            dept_name = df.iat[i,1].strip() if ncols > 1 else ""
            if dept_name:
                current_dept = dept_name
            else:
                current_dept = None
            continue

        # skip header rows that say Empcode or Name
        if row0.lower().startswith("empcode") or row0.lower().startswith("employee"):
            continue

        # If first cell looks numeric (empcode), treat as employee row
        first_cell = row0
        if first_cell and re.fullmatch(r'\d+', first_cell):
            empcode = first_cell
            name = df.iat[i,1].strip() if ncols > 1 else ""
            shift_cell = df.iat[i,2].strip() if ncols > 2 else ""

            # parse intervals from shift cell
            intervals = parse_shift_cell(shift_cell)
            shift_code = make_shift_code(intervals)

            # if shift_code is OFF (no intervals), we still upsert employee with no default shift
            if shift_code != "OFF":
                # create shift_master entry if not already
                if shift_code not in created_shifts:
                    # store description as original shift_cell for clarity
                    upsert_shift_master(db, shift_code, intervals, weekly_off=[], grace_minutes=10, break_minutes=60, description=shift_cell)
                    created_shifts[shift_code] = True
                    stats["shifts_created"] += 1
            else:
                # mark OFF with empty intervals in shift_master if not exists
                if "OFF" not in created_shifts:
                    upsert_shift_master(db, "OFF", [], weekly_off=[], grace_minutes=0, break_minutes=0, description="Off/Weekly off")
                    created_shifts["OFF"] = True
                    stats["shifts_created"] += 1

            # upsert employee with default_shift = shift_code (or None)
            default_shift = shift_code if shift_code != "OFF" else None
            upsert_employee(db, empcode, name, current_dept or "", default_shift=default_shift)
            stats["employees"] += 1

            # if year_month provided, bulk assign for month:
            if year_month:
                # assign default shift for all days of that month
                res = bulk_assign_shift_for_month(db, empcode, shift_code if shift_code!="OFF" else None, year_month)
                # we count assigned result only if non-empty
                stats["assigned_month"] += 1 if res is not None else 0

    print("Import complete.")
    print(f"Employees processed: {stats['employees']}")
    print(f"Shift patterns created: {stats['shifts_created']}")
    if year_month:
        print(f"Assigned default shift for month {year_month} to {stats['assigned_month']} employees.")
    return stats

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python shift_importer.py <path_to_shift_excel> [year-month]")
        sys.exit(1)
    path = sys.argv[1]
    ym = sys.argv[2] if len(sys.argv) > 2 else None
    print("Running importer on:", path, "year_month:", ym)
    stats = importer(path, year_month=ym)
    print(stats)
