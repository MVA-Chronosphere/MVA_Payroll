# shift_admin.py
import streamlit as st
from mongo_helpers import get_db, upsert_shift_master, assign_shift_to_employee_on_date, bulk_assign_shift_for_month, upsert_employee
import pandas as pd
from datetime import date

st.set_page_config(page_title="Shift Admin", layout="wide")
db = get_db()

st.title("Shift Admin")

with st.expander("Existing Shifts"):
    shifts = list(db.shift_master.find({}))
    if shifts:
        st.dataframe(pd.DataFrame(shifts))
    else:
        st.info("No shifts yet. Create one below.")

st.subheader("Create / Update Shift Pattern")
with st.form("shift_form"):
    shift_code = st.text_input("Shift Code (e.g. GEN, NIGHT, SPLIT1)", value="GEN")
    intervals_txt = st.text_area("Intervals (one per line, format start-end, e.g. 09:00-13:00 or 15:00-19:00)", value="09:00-18:00")
    weekly_off_txt = st.text_input("Weekly offs (comma separated, e.g. Sunday)", value="Sunday")
    grace = st.number_input("Grace minutes", min_value=0, max_value=120, value=10)
    brk = st.number_input("Break minutes", min_value=0, max_value=240, value=60)
    desc = st.text_area("Description", value="")
    submitted = st.form_submit_button("Save Shift")
    if submitted:
        intervals = []
        for line in intervals_txt.splitlines():
            s = line.strip()
            if not s: 
                continue
            if "-" in s:
                a, b = s.split("-", 1)
                intervals.append({"start": a.strip(), "end": b.strip()})
        weekly_off = [x.strip() for x in weekly_off_txt.split(",") if x.strip()]
        upsert_shift_master(db, shift_code, intervals, weekly_off=weekly_off, grace_minutes=grace, break_minutes=brk, description=desc)
        st.success(f"Shift {shift_code} saved.")

st.subheader("Assign Shift to Employee (single day)")
employees = list(db.employees.find({}, {"_id":1, "name":1}))
emp_map = {e["_id"]: e.get("name", "") for e in employees}
if not employees:
    st.info("No employees found. They will be created on ingest automatically if present in file.")
else:
    col1, col2, col3 = st.columns(3)
    with col1:
        emp_choice = st.selectbox("Employee", options=list(emp_map.keys()), format_func=lambda x: f"{x} - {emp_map.get(x)}")
    with col2:
        shift_choice = st.selectbox("Shift", options=[s["shift_code"] for s in shifts] if shifts else [])
    with col3:
        date_choice = st.date_input("Date", value=date.today())
    if st.button("Assign Shift"):
        assign_shift_to_employee_on_date(db, emp_choice, shift_choice, date_choice.strftime("%Y-%m-%d"))
        st.success("Assigned")

st.subheader("Bulk Assign (month)")
if employees and shifts:
    emp_bulk = st.selectbox("Employee (bulk)", options=list(emp_map.keys()), format_func=lambda x: f"{x} - {emp_map.get(x)}", key="bulk_emp")
    shift_bulk = st.selectbox("Shift (bulk)", options=[s["shift_code"] for s in shifts], key="bulk_shift")
    ym = st.text_input("Year-Month (YYYY-MM)", value="")
    if st.button("Bulk Assign"):
        if ym:
            res = bulk_assign_shift_for_month(db, emp_bulk, shift_bulk, ym)
            st.write("Result:", res)
            st.success("Bulk assignment complete")
        else:
            st.error("Provide Year-Month like 2025-07")
