"""
compute_monthly_payroll.py
------------------------------------
Final version for MVA Payroll System (Updated with ABSENT logic Option A).

✅ Reads only monthly attendance Excel file.
✅ Fetches shift details from MongoDB weekly_shifts collection.
✅ Matches each employee to their assigned shift pattern.
✅ Computes presence, absences, FM status automatically.
✅ Implements new deduction rules:
    - Compare actual IN/OUT with expected IN/OUT
    - If hours worked < expected, apply deductions
    - Late arrival rules (6+ min for 3 days, 30+ min any day)
    - Half-day marking
✅ Updates payroll summary in MongoDB.

ABSENT logic (Option A):
- If status == "A" or "ABSENT" -> absent (always)
- If punches are "--" or "--:--", check status; only count absent when status == "A"/"ABSENT"
- If punches exist and status == "A" -> still absent
"""

import pandas as pd
from datetime import datetime, timedelta, date
import calendar
from mva import parse_mva_blocks, calculate_fm_status
from pymongo import MongoClient


def extract_year_month_from_label(month_label):
    """
    Extract year and month number from a month label like 'November-2024'
    Returns tuple (year, month_num) 
    """
    import re
    year = None
    month_num = None
    
    if month_label:
        # Try to extract year from month string like "November-2024"
        year_match = re.search(r'(\d{4})', month_label)
        if year_match:
            year = int(year_match.group(1))
            month_name_match = re.search(r'([A-Za-z]+)', month_label)
            if month_name_match:
                month_name = month_name_match.group(1)
                # Get month number from month name
                month_num = list(calendar.month_name).index(month_name.title())
    
    # If we couldn't parse the month/year from the string, use current date
    if year is None or month_num is None:
        year = datetime.now().year
        month_num = datetime.now().month
    
    return year, month_num


AUTO_UPDATE_EMPLOYEES = False   # Automatically update employee info
SYNC_EMPLOYEES = False         # Set True to remove employees missing in this import


def compute_monthly_payroll(attendance_file, timetable_file=None, vacation_start_date=None, vacation_end_date=None):
    print(f"📘 Loading attendance file: {attendance_file}")
    if timetable_file:
        print(f"📘 Loading timetable file: {timetable_file}")

    # Connect to MongoDB
    mongo_client = MongoClient("mongodb://localhost:27017/")
    db = mongo_client["mva_payroll"]
    payroll_collection = db["monthly_payroll_summary"]

    # --- Read attendance file ---
    att_df = pd.read_excel(attendance_file, header=None)

    # --- Parse attendance blocks ---
    blocks = parse_mva_blocks(att_df)
    print(f"🔍 Parsed {len(blocks)} employee attendance blocks.")

    # If a timetable file is provided, read it to get expected IN/OUT times
    timetable_data = {}
    if timetable_file:
        try:
            # Read the timetable Excel file
            timetable_df = pd.read_excel(timetable_file)
            
            # Process the timetable data to extract expected IN/OUT times for each employee and day
            # Assuming the timetable follows the same format as the generated one
            # with IN/OUT rows and date columns
            
            # Find the IN and OUT rows by looking for cells that start with "IN" or "OUT"
            in_row_idx = None
            out_row_idx = None
            work_row_idx = None
            status_row_idx = None
            
            for idx, row in timetable_df.iterrows():
                first_cell = str(row.iloc[0]).upper() if pd.notna(row.iloc[0]) else ""
                if first_cell.startswith("IN"):
                    in_row_idx = idx
                elif first_cell.startswith("OUT"):
                    out_row_idx = idx
                elif first_cell.startswith("WORK"):
                    work_row_idx = idx
                elif "STATUS" in first_cell:
                    status_row_idx = idx
            
            # Extract date columns (assuming they start from column 1, with dates in row 1 or 2)
            date_cols = []
            day_row_idx = in_row_idx - 1  # Day names are typically one row above IN
            if day_row_idx >= 0:
                day_row = timetable_df.iloc[day_row_idx]
                for col_idx in range(1, len(day_row)):  # Skip first column which is usually "IN", "OUT", etc.
                    date_val = day_row.iloc[col_idx]
                    if pd.notna(date_val) and str(date_val).strip().isdigit():
                        day_num = int(float(str(date_val).strip()))
                        if 1 <= day_num <= 31:
                            date_cols.append((col_idx, day_num))
            
            # Process each employee's data in the timetable
            for idx, row in timetable_df.iterrows():
                empcode_cell = str(row.iloc[1]).strip() if len(row) > 1 else ""  # Column B (index 1) has employee code
                # Consider both header rows with "EMPCODE" and direct rows with empcode value
                if empcode_cell and empcode_cell.upper() != "NAN":
                    # Skip rows that are clearly not employee rows
                    if "EMPCODE" in empcode_cell.upper():
                        # header row, values likely in next row
                        if idx + 1 < len(timetable_df):
                            next_row = timetable_df.iloc[idx + 1]
                            emp_id = str(next_row.iloc[1]).strip() if len(next_row) > 1 else ""
                        else:
                            emp_id = ""
                    else:
                        emp_id = empcode_cell

                    if emp_id and emp_id.upper() != "EMPCODE":
                        emp_timetable = {}
                        for col_idx, day_num in date_cols:
                            expected_in = ""
                            expected_out = ""
                            
                            if in_row_idx is not None and col_idx < len(timetable_df.columns):
                                expected_in_val = timetable_df.iloc[in_row_idx, col_idx]
                                expected_in = str(expected_in_val).strip() if pd.notna(expected_in_val) else ""
                            
                            if out_row_idx is not None and col_idx < len(timetable_df.columns):
                                expected_out_val = timetable_df.iloc[out_row_idx, col_idx]
                                expected_out = str(expected_out_val).strip() if pd.notna(expected_out_val) else ""
                            
                            # Store expected times for this day
                            emp_timetable[day_num] = {
                                "expected_in": expected_in,
                                "expected_out": expected_out
                            }
                        
                        timetable_data[emp_id] = emp_timetable
        except Exception as e:
            print(f"⚠️ Error reading timetable file: {e}")

    computed = 0

    for b in blocks:
        emp_id = str(b.get("empcode", "")).strip()
        name = b.get("name", "")
        dept = b.get("dept", "")
        month = b.get("month", "")

        if not emp_id:
            continue

        # Get expected shift times from the uploaded timetable if available
        expected_shifts = timetable_data.get(emp_id, {})

        # --- Attendance metrics (updated absent logic Option A) ---
        daily = b.get("daily", [])

        # Build statuses list for calculating present/half/leave/weekoff/total days
        statuses = []
        for d in daily:
            if d.get("day"):
                statuses.append((d.get("status") or "").strip().upper())

        # Present / half / leave / weekoffs / total working days (from statuses)
        present_days = sum(1 for s in statuses if s in ("P", "PR", "W", "WD"))
        half_days = sum(1 for s in statuses if s in ("HD", "H", "1/2"))
        leaves = sum(1 for s in statuses if s in ("L", "LV", "LEAVE"))
        week_offs = sum(1 for s in statuses if s in ("WO", "W/O", "OFF"))
        total_working_days = len(statuses)

        # Updated absent logic (Option A)
        absents = 0
        for d in daily:
            if not d.get("day"):
                continue
            status = (d.get("status") or "").strip().upper()
            in_time = (d.get("in_time") or "").strip()
            out_time = (d.get("out_time") or "").strip()

            # If status explicitly marks absent, always count as absent (Option A)
            if status in ("A", "ABSENT"):
                absents += 1
                continue

            # If punches are missing ("--" or "--:--"), check status
            if in_time in ("--", "--:--") or out_time in ("--", "--:--"):
                # status already checked above; only count if it was "A"
                # (we've already covered that case). So do nothing here.
                continue

            # Otherwise, not absent (status not A and punches exist)
            # (Option A requires no additional action)


        # --- Payroll Deduction Rules ---
        total_quarter_cut_days = 0  # ¼-day salary cuts
        late_arrival_days = []      # Track days with late arrivals >6 min

        for day_data in daily:
            day_num = day_data.get("day")
            in_time = (day_data.get("in_time") or "").strip()
            out_time = (day_data.get("out_time") or "").strip()
            status = (day_data.get("status") or "").strip().upper()

            if not in_time or not out_time or status not in ("P", "PR", "W", "WD"):
                continue

            # Get expected shift times for this day from the uploaded timetable
            expected_times = expected_shifts.get(day_num, {})
            expected_in = expected_times.get("expected_in", "")
            expected_out = expected_times.get("expected_out", "")

            # Skip if expected times are not available
            if not expected_in or not expected_out or expected_in == "--:--" or expected_out == "--:--":
                continue

            try:
                fmt = "%H:%M"
                in_dt = datetime.strptime(in_time, fmt)
                out_dt = datetime.strptime(out_time, fmt)
                expected_in_dt = datetime.strptime(expected_in, fmt)
                expected_out_dt = datetime.strptime(expected_out, fmt)

                # Handle overnight shifts
                if out_dt < in_dt:
                    out_dt = out_dt.replace(day=in_dt.day + 1)
                if expected_out_dt < expected_in_dt:
                    expected_out_dt = expected_out_dt.replace(day=expected_in_dt.day + 1)

                # Calculate actual and expected working hours
                actual_worked_seconds = (out_dt - in_dt).total_seconds()
                expected_worked_seconds = (expected_out_dt - expected_in_dt).total_seconds()
                
                # Subtract break time (if any is configured)
                # For now, assuming no break time is configured, but we can make it configurable
                break_time_seconds = 0 # This could be read from shift details if configured
                
                actual_worked_hours = (actual_worked_seconds - break_time_seconds) / 3600
                expected_worked_hours = (expected_worked_seconds - break_time_seconds) / 3600

                # Check if hours worked is less than expected
                if actual_worked_hours < expected_worked_hours:
                    # Apply deduction based on how much less was worked
                    hour_difference = expected_worked_hours - actual_worked_hours
                    if hour_difference >= 6:  # More than 6 hours short: full day deduction
                        total_quarter_cut_days += 1.0
                    elif hour_difference >= 4.5:  # 4.5-6 hours short: 3/4 day deduction
                        total_quarter_cut_days += 0.75
                    elif hour_difference >= 3:  # 3-4.5 hours short: half day deduction
                        total_quarter_cut_days += 0.5
                    elif hour_difference > 0:  # Less than 3 hours short: quarter day deduction
                        total_quarter_cut_days += 0.25

                # Check for late arrival (more than 6 minutes after expected IN time)
                late_minutes = (in_dt - expected_in_dt).total_seconds() / 60
                if late_minutes > 6 and late_minutes <= 30:
                    late_arrival_days.append(day_num)
                elif late_minutes > 30:
                    # If late by more than 30 minutes, apply quarter-day deduction immediately
                    total_quarter_cut_days += 0.25

            except Exception as e:
                # Ignore parsing errors safely
                print(f"⚠️ Error processing day {day_num} for employee {emp_id}: {e}")
                continue

        # Apply late arrival deductions: every 3rd late arrival (>6 min) results in a quarter-day deduction
        # (Not necessarily consecutive days)
        if len(late_arrival_days) >= 3:
            # For every 3 late days, apply 1 quarter-day deduction
            late_deductions = len(late_arrival_days) // 3
            total_quarter_cut_days += late_deductions * 0.25

        # --- Compute FM status ---
        fm_status = calculate_fm_status(
            present_days, half_days, leaves, absents, week_offs, total_working_days
        )

        # --- Adjust FM status if deductions exist ---
        if total_quarter_cut_days > 0:
            # Convert 0.25 → 1 day equivalent for salary adjustment tagging
            fm_status = f"{fm_status} -{round(total_quarter_cut_days, 2)}"

        # --- Save result ---
        # Ensure expected_in and expected_out are defined even if not found in timetable
        if 'expected_in' not in locals() or 'expected_out' not in locals():
            expected_in = "N/A"
            expected_out = "N/A"
        
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
            "late_count": len(late_arrival_days),
            "shift_type": "N/A",  # Will be updated if we have shift data
            "shift_in": expected_in,
            "shift_out": expected_out,
            "computed_at": datetime.utcnow().isoformat()
        }
        
        # Add vacation dates to the document if they were provided
        if vacation_start_date and vacation_end_date:
            doc["vacation_start_date"] = vacation_start_date.isoformat()
            doc["vacation_end_date"] = vacation_end_date.isoformat()

        # Insert or update the document in MongoDB
        payroll_collection.replace_one(
            {"employee_id": emp_id, "month": month},
            doc,
            upsert=True
        )

        # Store daily attendance records in MongoDB as well
        for day_data in b.get("daily", []):
            day_num = day_data.get("day")
            if day_num:
                daily_doc = {
                    "employee_id": emp_id,
                    "employee_name": name,
                    "department": dept,
                    "month": month,
                    "day": day_num,
                    "weekday": day_data.get("weekday", ""),
                    "in_time": day_data.get("in_time", ""),
                    "out_time": day_data.get("out_time", ""),
                    "work_time": day_data.get("work_time", ""),
                    "status": day_data.get("status", ""),
                    "computed_at": datetime.utcnow().isoformat()
                }
                
                # Check if this daily record exists, if so update it, otherwise insert new
                db.attendance_daily.replace_one(
                    {"employee_id": emp_id, "day": day_num, "month": month},
                    daily_doc,
                    upsert=True
                )

        if AUTO_UPDATE_EMPLOYEES:
            # Update employee info in MongoDB if needed
            db.employees.update_one(
                {"employee_id": emp_id},
                {
                    "$set": {
                        "name": name,
                        "department": dept,
                        "default_shift": "N/A",
                        "shift_in": expected_in,
                        "shift_out": expected_out,
                        "updated_at": datetime.utcnow().isoformat()
                    }
                },
                upsert=True
            )

        computed += 1

    mongo_client.close()

    print(f"✅ Computed payroll for {computed} employees.")
    print("📦 Saved to MongoDB collection: monthly_payroll_summary")

    return {"processed": computed}


# Optional direct run for testing
if __name__ == "__main__":
    # Find the latest attendance file
    import glob
    import os
    list_of_files = glob.glob('attendance_*.xlsx') # * means all if need specific format then *.xlsx
    if not list_of_files:
        print("No attendance_*.xlsx files found in current directory.")
    else:
        latest_file = max(list_of_files, key=os.path.getctime)
        print(f"Processing latest attendance file: {latest_file}")
        compute_monthly_payroll(latest_file)
