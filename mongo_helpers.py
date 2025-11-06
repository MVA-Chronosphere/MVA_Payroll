# mongo_helpers.py
from pymongo import MongoClient, UpdateOne
from datetime import datetime
from bson import ObjectId

# << EDIT THIS to your MongoDB URI >>
MONGO_URI = "mongodb://localhost:27017"
DB_NAME = "mva_payroll"

def get_db():
    client = MongoClient(MONGO_URI)
    return client[DB_NAME]

# ========== Employees ==========
def upsert_employee(db, emp_id, name, department, default_shift=None):
    if not emp_id:
        return
    db.employees.update_one(
        {"_id": str(emp_id)},
        {"$set": {
            "name": name,
            "department": department,
            "default_shift": default_shift,
            "status": "Active",
            "updated_at": datetime.utcnow()
        }},
        upsert=True
    )

# ========== Shift master ==========
def upsert_shift_master(db, shift_code, intervals, weekly_off=None, grace_minutes=10, break_minutes=60, description=""):
    """
    intervals: list of {"start":"HH:MM","end":"HH:MM"} - supports split shifts and overnight end < start
    weekly_off: list of weekday names like ["Sunday"]
    """
    doc = {
        "shift_code": shift_code,
        "intervals": intervals,
        "weekly_off": weekly_off or [],
        "grace_minutes": int(grace_minutes),
        "break_minutes": int(break_minutes),
        "description": description,
        "updated_at": datetime.utcnow()
    }
    db.shift_master.update_one({"shift_code": shift_code}, {"$set": doc}, upsert=True)
    return db.shift_master.find_one({"shift_code": shift_code})

# ========== Employee-specific daywise assignments (shift calendar) ==========
def assign_shift_to_employee_on_date(db, employee_id, shift_code, date_str):
    """
    date_str: 'YYYY-MM-DD'
    Creates/updates an assignment for that one date.
    """
    doc = {
        "employee_id": str(employee_id),
        "date": date_str,
        "shift_code": shift_code,
        "assigned_at": datetime.utcnow()
    }
    db.employee_shift_assignments.update_one({"employee_id": str(employee_id), "date": date_str},
                                            {"$set": doc}, upsert=True)
    return True

def bulk_assign_shift_for_month(db, employee_id, shift_code, year_month):
    """
    year_month: 'YYYY-MM'
    Bulk assigns given shift_code to all dates of that month.
    """
    import calendar
    y, m = [int(x) for x in year_month.split('-')]
    ops = []
    for day in range(1, calendar.monthrange(y, m)[1] + 1):
        date_str = f"{y}-{m:02d}-{day:02d}"
        doc = {
            "employee_id": str(employee_id),
            "date": date_str,
            "shift_code": shift_code,
            "assigned_at": datetime.utcnow()
        }
        ops.append(UpdateOne({"employee_id": str(employee_id), "date": date_str}, {"$set": doc}, upsert=True))
    if ops:
        res = db.employee_shift_assignments.bulk_write(ops, ordered=False)
        return res.bulk_api_result
    return None

# ========== Save uploaded file metadata ==========
def save_uploaded_file_meta(db, filename, month=None, department=None, uploader="admin"):
    doc = {
        "filename": filename,
        "uploaded_at": datetime.utcnow(),
        "month": month,
        "department": department,
        "uploader": uploader
    }
    res = db.attendance_uploaded_files.insert_one(doc)
    return res.inserted_id

# ========== Ingest parsed blocks to attendance_daily ==========
def ingest_blocks_to_db(db, blocks, source_file_id=None, year_month=None):
    """
    blocks: output of mva.parse_mva_blocks
    source_file_id: optional ObjectId or None
    year_month: 'YYYY-MM' to help construct dates (optional)
    """
    ops = []
    for b in blocks:
        emp_id = str(b.get("empcode", "")).strip()
        name = b.get("name", "").strip()
        dept = b.get("dept", "").strip()
        month_label = b.get("month", "")
        ym = year_month
        if not ym and month_label:
            try:
                parts = month_label.replace(" ", "").split("-")
                if len(parts) == 2:
                    mm_map = {'jan':'01','feb':'02','mar':'03','apr':'04','may':'05','jun':'06',
                              'jul':'07','aug':'08','sep':'09','oct':'10','nov':'11','dec':'12'}
                    mstr = parts[0][:3].lower()
                    yy = parts[1]
                    mm = mm_map.get(mstr, '01')
                    ym = f"{yy}-{mm}"
            except:
                ym = None

        # upsert employee meta
        upsert_employee(db, emp_id, name, dept)

        # ingest daily rows
        for rec in b.get("daily", []):
            dnum = rec.get("day")
            if not dnum:
                continue
            if ym:
                y, m = ym.split("-")
                daystr = f"{y}-{int(m):02d}-{int(dnum):02d}"
            else:
                today = datetime.utcnow()
                daystr = f"{today.year}-{today.month:02d}-{int(dnum):02d}"

            status = (rec.get("status") or "").strip()
            # allow multiple punches as strings or comma-separated - preserve raw to process later
            in_time = rec.get("in_time") or ""
            out_time = rec.get("out_time") or ""
            work_time_raw = rec.get("work_time", None)

            # try parse numeric work_time else 0
            try:
                work_hours = float(work_time_raw) if work_time_raw not in (None, "") else 0.0
            except:
                work_hours = 0.0

            daily_doc = {
                "employee_id": emp_id,
                "employee_name": name,
                "department": dept,
                "date": daystr,
                "punch_in_raw": str(in_time).strip(),
                "punch_out_raw": str(out_time).strip(),
                "raw_status": status,
                "status": status,
                "work_hours": work_hours,
                "ot_hours": 0.0,
                "late_minutes": 0,
                "early_exit_minutes": 0,
                "overnight": False,
                "wrong_shift": False,
                "miss_punch": False,
                "assigned_shift_code": None,
                "source_file_id": ObjectId(source_file_id) if source_file_id else None,
                "meta": {"imported_from_block": True},
                "created_at": datetime.utcnow(),
                "updated_at": datetime.utcnow()
            }
            ops.append(UpdateOne({"employee_id": emp_id, "date": daystr}, {"$set": daily_doc}, upsert=True))
    if ops:
        res = db.attendance_daily.bulk_write(ops, ordered=False)
        return res.bulk_api_result
    return None
