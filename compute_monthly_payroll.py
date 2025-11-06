"""
compute_monthly_payroll.py
------------------------------------
Final version for MVA Payroll System.

✅ Reads only monthly attendance Excel file.
✅ Fetches shift details directly from MongoDB (`shift_master`).
✅ Matches each employee to their assigned/default shift.
✅ Computes presence, absences, FM status automatically.
✅ Updates payroll summary in MongoDB.
"""

import pandas as pd
from datetime import datetime
from mongo_helpers import get_db
from mva import parse_mva_blocks, calculate_fm_status


AUTO_UPDATE_EMPLOYEES = True   # Automatically update employee info
SYNC_EMPLOYEES = False         # Set True to remove employees missing in this import


def compute_monthly_payroll(attendance_file):
    print(f"📘 Loading attendance file: {attendance_file}")

    # --- Read attendance file ---
    att_df = pd.read_excel(attendance_file, header=None)
    db = get_db()
    payroll_coll = db["monthly_payroll_summary"]
    employees_coll = db["employees"]
    shift_coll = db["shift_master"]
    assign_coll = db["employee_shift_assignments"]

    # --- Parse attendance blocks ---
    blocks = parse_mva_blocks(att_df)
    print(f"🔍 Parsed {len(blocks)} employee attendance blocks.")

    # --- Prepare shift lookup from DB ---
    shifts = list(shift_coll.find({}))
    shift_lookup = {s["shift_code"]: s for s in shifts}

    # --- Prepare employee-to-shift mapping ---
    emp_shift_map = {}
    for emp in employees_coll.find({}, {"_id": 0, "employee_id": 1, "default_shift": 1}):
        emp_shift_map[str(emp["employee_id"]).strip()] = emp.get("default_shift")

    # Optionally remove employees missing in this attendance file
    if SYNC_EMPLOYEES:
        imported_ids = {str(b.get("empcode", "")).strip() for b in blocks if b.get("empcode")}
        existing_ids = {e["employee_id"] for e in employees_coll.find({}, {"employee_id": 1})}
        to_remove = existing_ids - imported_ids
        if to_remove:
            res = employees_coll.delete_many({"employee_id": {"$in": list(to_remove)}})
            print(f"🗑️ Removed {res.deleted_count} employees not present in this import.")

    computed = 0

    for b in blocks:
        emp_id = str(b.get("empcode", "")).strip()
        name = b.get("name", "")
        dept = b.get("dept", "")
        month = b.get("month", "")

        if not emp_id:
            continue

        # --- Determine shift ---
        assigned_shift = assign_coll.find_one({"employee_id": emp_id})
        shift_code = (
            assigned_shift.get("shift_code")
            if assigned_shift
            else emp_shift_map.get(emp_id)
        )
        shift_info = shift_lookup.get(shift_code, {})

        shift_in = ""
        shift_out = ""
        shift_type = shift_code or "N/A"
        role_type = shift_info.get("description", "").lower()  # detect 'split' role
        if shift_info and "intervals" in shift_info and shift_info["intervals"]:
            primary = shift_info["intervals"][0]
            shift_in = primary.get("start", "")
            shift_out = primary.get("end", "")

        # --- Attendance metrics ---
        daily = b.get("daily", [])
        statuses = [d.get("status", "").strip().upper() for d in daily if d.get("day")]

        present_days = sum(1 for s in statuses if s in ("P", "PR", "W", "WD"))
        half_days = sum(1 for s in statuses if s in ("HD", "H", "1/2"))
        leaves = sum(1 for s in statuses if s in ("L", "LV", "LEAVE"))
        week_offs = sum(1 for s in statuses if s in ("WO", "W/O", "OFF"))
        absents = sum(1 for s in statuses if s in ("A", "AB", "ABSENT"))
        total_working_days = len(statuses)

        # --- Additional Payroll Rules ---
        total_quarter_cut_days = 0  # ¼-day salary cuts
        total_late_entries = 0      # Late >5min but ≤30min count

        for day_data in daily:
            in_time = day_data.get("in_time", "").strip()
            out_time = day_data.get("out_time", "").strip()
            status = (day_data.get("status", "") or "").strip().upper()

            if not in_time or not out_time or status not in ("P", "PR", "W", "WD"):
                continue

            try:
                fmt = "%H:%M"
                in_dt = datetime.strptime(in_time, fmt)
                out_dt = datetime.strptime(out_time, fmt)

                # Handle overnight
                if out_dt < in_dt:
                    out_dt = out_dt.replace(day=in_dt.day + 1)

                worked_hours = (out_dt - in_dt).total_seconds() / 3600

                # Handle split shifts: must work >= 9 hours total
                if "split" in role_type and worked_hours < 9:
                    total_quarter_cut_days += 0.25  # less than 9 hrs → deduct ¼
                    continue

                # --- Late arrival checks ---
                expected_in = datetime.strptime(shift_in, fmt) if shift_in else None
                if expected_in:
                    late_mins = (in_dt - expected_in).total_seconds() / 60
                    if late_mins > 30:
                        # More than 30 mins late → direct ¼-day cut
                        total_quarter_cut_days += 0.25
                    elif late_mins > 5:
                        # More than 5 mins late (within 30 mins) → track count
                        total_late_entries += 1

            except Exception as e:
                # Ignore parsing errors safely
                continue

        # For every 3 late entries (>5min), apply 1 quarter-day cut
        total_quarter_cut_days += (total_late_entries // 3) * 0.25

        # --- Compute FM status ---
        fm_status = calculate_fm_status(
            present_days, half_days, leaves, absents, week_offs, total_working_days
        )

        # --- Adjust FM status if deductions exist ---
        if total_quarter_cut_days > 0:
            # Convert 0.25 → 1 day equivalent for salary adjustment tagging
            fm_status = f"{fm_status} -{round(total_quarter_cut_days, 2)}"

        # --- Save result ---
        doc = {
            "employee_id": emp_id,
            "employee_name": name,
            "department": dept,
            "month": month,
            "present_days": present_days,
            "half_days": half_days,
            "leaves": leaves,
            "week_offs": week_offs,
            "absent_days": absents,
            "total_working_days": total_working_days,
            "full_month_status": fm_status,
            "quarter_day_cuts": round(total_quarter_cut_days, 2),
            "late_count": total_late_entries,
            "shift_type": shift_type,
            "shift_in": shift_in,
            "shift_out": shift_out,
            "computed_at": datetime.utcnow()
        }

        payroll_coll.update_one(
            {"employee_id": emp_id, "month": month},
            {"$set": doc},
            upsert=True
        )

        if AUTO_UPDATE_EMPLOYEES:
            employees_coll.update_one(
                {"employee_id": emp_id},
                {"$set": {
                    "name": name,
                    "department": dept,
                    "default_shift": shift_code,
                    "shift_in": shift_in,
                    "shift_out": shift_out,
                    "updated_at": datetime.utcnow()
                }},
                upsert=True
            )

        computed += 1


    print(f"✅ Computed payroll for {computed} employees.")
    print("📦 Saved to MongoDB collection: monthly_payroll_summary")

    return {"processed": computed}


# Optional direct run for testing
if __name__ == "__main__":
    attendance_path = "monthperformance05112025114624.xlsx"
    compute_monthly_payroll(attendance_path)
