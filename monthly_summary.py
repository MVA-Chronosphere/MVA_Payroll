# monthly_summary.py
from pymongo import UpdateOne
from datetime import datetime
from mongo_helpers import get_db
from mva import calculate_fm_status

def calc_total_working_days(year_month):
    try:
        y, m = year_month.split("-")
        import calendar
        y = int(y); m = int(m)
        total = 0
        for day in range(1, calendar.monthrange(y, m)[1] + 1):
            wkday = datetime(y, m, day).weekday()
            if wkday < 5:
                total += 1
        return total
    except:
        return 26

def generate_monthly_summary(year_month, db=None, source_file_id=None):
    if db is None:
        db = get_db()
    pipeline = [
        {"$match": {"date": {"$regex": f"^{year_month}"}}}, 
        {"$group": {
            "_id": "$employee_id",
            "present_days": {"$sum": {"$cond": [{"$in": ["$status", ["P","P OT","POT","P OT","POT:"]]}, 1, 0]}},
            "half_days": {"$sum": {"$cond": [{"$eq": ["$status", "H"]}, 1, 0]}},
            "leaves": {"$sum": {"$cond": [{"$eq": ["$status", "L"]}, 1, 0]}},
            "week_offs": {"$sum": {"$cond": [{"$eq": ["$status", "WO"]}, 1, 0]}},
            "absent_days": {"$sum": {"$cond": [{"$eq": ["$status", "A"]}, 1, 0]}},
            "wrong_shifts": {"$sum": {"$cond": [{"$eq": ["$wrong_shift", True]}, 1, 0]}},
            "miss_punches": {"$sum": {"$cond": [{"$eq": ["$miss_punch", True]}, 1, 0]}},
            "late_occurrences": {"$sum": {"$cond": [{"$gt": ["$late_minutes", 0]}, 1, 0]}},
            "total_late_minutes": {"$sum": "$late_minutes"},
            "total_ot_hours": {"$sum": "$ot_hours"},
            "total_work_hours": {"$sum": "$work_hours"},
            "count_days": {"$sum": 1}
        }},
        {"$lookup": {
            "from": "employees",
            "localField": "_id",
            "foreignField": "_id",
            "as": "employee"
        }},
        {"$unwind": {"path": "$employee", "preserveNullAndEmptyArrays": True}},
        {"$project": {
            "employee_id": "$_id",
            "employee_name": "$employee.name",
            "department": "$employee.department",
            "present_days": 1,
            "half_days": 1,
            "leaves": 1,
            "week_offs": 1,
            "absent_days": 1,
            "wrong_shifts": 1,
            "miss_punches": 1,
            "late_occurrences": 1,
            "total_late_minutes": 1,
            "total_ot_hours": 1,
            "total_work_hours": 1
        }}
    ]
    results = list(db.attendance_daily.aggregate(pipeline))
    total_working_days = calc_total_working_days(year_month)
    ops = []
    out_docs = []
    for r in results:
        present = r.get("present_days", 0)
        half = r.get("half_days", 0)
        leaves = r.get("leaves", 0)
        absents = r.get("absent_days", 0)
        wo = r.get("week_offs", 0)
        total_ot = round(r.get("total_ot_hours", 0.0) or 0.0, 2)
        fm_status = calculate_fm_status(present, half, leaves, absents, wo, total_working_days)
        doc = {
            "employee_id": r.get("employee_id"),
            "employee_name": r.get("employee_name"),
            "department": r.get("department"),
            "year_month": year_month,
            "present_days": present,
            "half_days": half,
            "leaves": leaves,
            "week_offs": wo,
            "absent_days": absents,
            "wrong_shifts": r.get("wrong_shifts", 0),
            "miss_punches": r.get("miss_punches", 0),
            "late_occurrences": r.get("late_occurrences", 0),
            "total_late_minutes": int(r.get("total_late_minutes", 0)),
            "total_ot_hours": total_ot,
            "total_work_hours": round(r.get("total_work_hours", 0.0) or 0.0, 2),
            "fm_status": fm_status,
            "generated_at": datetime.utcnow(),
            "source_file_id": source_file_id
        }
        out_docs.append(doc)
        ops.append(UpdateOne({"employee_id": doc["employee_id"], "year_month": year_month}, {"$set": doc}, upsert=True))
    if ops:
        db.attendance_summary.bulk_write(ops)
    return out_docs
