# compute_attendance_metrics.py
from datetime import datetime, timedelta
from mongo_helpers import get_db
import math

def hhmm_to_minutes(tstr):
    if not tstr:
        return None
    try:
        s = str(tstr).strip()
        if " " in s and ":" in s:
            s = s.split(" ")[-1]
        parts = s.split(":")
        h = int(parts[0]); m = int(parts[1])
        return h*60 + m
    except:
        return None

def minutes_to_hours(m):
    return round(m/60.0, 2)

def normalize_punches(raw_in, raw_out):
    """
    Accept strings like:
    - '08:45' single in/out
    - '08:45,15:00|13:00,18:00' not standard; we will assume simple one in/one out.
    For now return first pair. Later we can extend to multiple punches per day.
    """
    in_val = raw_in.strip() if raw_in else ""
    out_val = raw_out.strip() if raw_out else ""
    return in_val, out_val

def fetch_assigned_shift(db, emp_id, date_str):
    # 1) check for explicit employee assignment for that date
    doc = db.employee_shift_assignments.find_one({"employee_id": str(emp_id), "date": date_str})
    if doc:
        return db.shift_master.find_one({"shift_code": doc.get("shift_code")})
    # 2) check employee default_shift
    emp = db.employees.find_one({"_id": str(emp_id)})
    if emp and emp.get("default_shift"):
        return db.shift_master.find_one({"shift_code": emp.get("default_shift")})
    # 3) fallback: if no assignment, return None (compute step will handle)
    return None

def compute_for_month(year_month, db=None):
    """
    Process attendance_daily rows where date starts with year_month (YYYY-MM).
    Writes computed metrics back to attendance_daily.
    """
    if db is None:
        db = get_db()
    regex = f"^{year_month}"
    cursor = db.attendance_daily.find({"date": {"$regex": regex}})
    processed = 0
    for doc in cursor:
        processed += 1
        emp_id = doc.get("employee_id")
        date_str = doc.get("date")
        raw_in = doc.get("punch_in_raw", "")
        raw_out = doc.get("punch_out_raw", "")

        # normalize and get primary in/out
        in_val, out_val = normalize_punches(raw_in, raw_out)

        # fetch assigned shift for that date
        shift = fetch_assigned_shift(db, emp_id, date_str) or {}
        intervals = shift.get("intervals", [])  # list of {start,end}
        grace = int(shift.get("grace_minutes", 0)) if shift else 0
        break_minutes = int(shift.get("break_minutes", 0)) if shift else 0

        in_min = hhmm_to_minutes(in_val)
        out_min = hhmm_to_minutes(out_val)

        # If no assigned shift, try to infer: if employee.default_shift exists, use it; else leave None
        exp_in_min = None
        exp_out_min = None
        assigned_shift_code = None
        if intervals:
            assigned_shift_code = shift.get("shift_code")
            # For now, take the first interval as primary expected times for late/early comparisons.
            # This handles simple shifts. For split shifts, better logic below will consider multiple intervals.
            primary = intervals[0]
            exp_in_min = hhmm_to_minutes(primary.get("start"))
            exp_out_min = hhmm_to_minutes(primary.get("end"))

        work_minutes = 0
        ot_minutes = 0
        late_minutes = 0
        early_exit_minutes = 0
        overnight = False
        miss_punch = False
        wrong_shift = False

        # handle missing punches
        if in_min is None or out_min is None:
            miss_punch = True

        # handle overnight where out < in
        if in_min is not None and out_min is not None and out_min < in_min:
            out_min = out_min + 24*60
            overnight = True

        # If there are multiple intervals defined (split-shift), try to match actual punches to intervals.
        # Basic approach: assume single in/out -> calculate raw duration minus break_minutes.
        if in_min is not None and out_min is not None:
            raw_mins = max(0, out_min - in_min - int(break_minutes))
            work_minutes = raw_mins

            # compute OT relative to shift end if available
            if exp_out_min is not None:
                adj_exp_out = exp_out_min
                # if exp_out_min < exp_in_min, it's overnight shift -> add 24h
                if exp_in_min is not None and exp_out_min < exp_in_min:
                    adj_exp_out = exp_out_min + 24*60
                if out_min > adj_exp_out:
                    ot_minutes = max(0, out_min - adj_exp_out)

            # late: in > expected_in + grace
            if exp_in_min is not None and in_min > (exp_in_min + grace):
                late_minutes = max(0, in_min - (exp_in_min + grace))

            # early exit: out < expected_out
            if exp_out_min is not None:
                adj_exp_out = exp_out_min
                if exp_in_min is not None and exp_out_min < exp_in_min:
                    adj_exp_out = exp_out_min + 24*60
                if out_min < (adj_exp_out):
                    early_exit_minutes = max(0, adj_exp_out - out_min)

        # wrong-shift heuristic: if expected exists and in diff > 6 hours -> wrong shift
        if exp_in_min is not None and in_min is not None and abs(in_min - exp_in_min) > 360:
            wrong_shift = True

        update = {
            "work_hours": round(work_minutes/60.0, 2),
            "ot_hours": round(ot_minutes/60.0, 2),
            "late_minutes": int(late_minutes),
            "early_exit_minutes": int(early_exit_minutes),
            "overnight": overnight,
            "miss_punch": miss_punch,
            "wrong_shift": wrong_shift,
            "assigned_shift_code": assigned_shift_code,
            "updated_at": datetime.utcnow()
        }
        db.attendance_daily.update_one({"_id": doc["_id"]}, {"$set": update})
    return {"processed": processed}
