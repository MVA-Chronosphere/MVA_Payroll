import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from pymongo import MongoClient
import re

# Mongo connection
client = MongoClient("mongodb://localhost:27017/")
db = client["mva_payroll"]
shift_collection = db["shift_master"]

def parse_time(t):
    """Convert time string to HH:MM 24hr"""
    if pd.isna(t) or t == "":
        return None
    try:
        t = str(t).strip()
        return datetime.strptime(t, "%I:%M %p").strftime("%H:%M")
    except:
        try:
            return datetime.strptime(t, "%H:%M").strftime("%H:%M")
        except:
            return None

def load_shift_file(file, year_month):
    df = pd.read_excel(file)
    df.columns = [c.strip().lower() for c in df.columns]

    for _, row in df.iterrows():
        emp_id = str(row.get("id", "")).strip()
        name = row.get("name", "").strip()

        if not emp_id:
            continue

        shift_records = []

        for day in range(1, 32):
            day_col_in = f"in{day}"
            day_col_out = f"out{day}"

            if day_col_in in df.columns:
                in_time = parse_time(row.get(day_col_in))
                out_time = parse_time(row.get(day_col_out))

                if not in_time or not out_time:
                    continue

                date_str = f"{year_month}-{day:02d}"

                # Handle night shift next-day logic
                dt_in = datetime.strptime(f"{date_str} {in_time}", "%Y-%m-%d %H:%M")
                dt_out = datetime.strptime(f"{date_str} {out_time}", "%Y-%m-%d %H:%M")
                if dt_out < dt_in:
                    dt_out += timedelta(days=1)

                shift_records.append({
                    "date": date_str,
                    "in": in_time,
                    "out": out_time,
                    "expected_hours": round((dt_out - dt_in).total_seconds() / 3600, 2)
                })

        if shift_records:
            shift_collection.update_one(
                {"employee_id": emp_id, "month": year_month},
                {"$set": {
                    "employee_id": emp_id,
                    "name": name,
                    "month": year_month,
                    "shifts": shift_records
                }},
                upsert=True
            )

    return True


# Streamlit UI
st.title("🕒 Shift Master Upload")

uploaded_file = st.file_uploader("Upload Staff Shift Excel File", type=["xlsx"])

year_month = st.text_input("Enter year-month (YYYY-MM)", "")

if st.button("Upload and Store Shifts"):
    if not uploaded_file or not year_month:
        st.error("Please upload file and enter month like 2025-07")
    else:
        try:
            load_shift_file(uploaded_file, year_month)
            st.success("Shift data imported successfully!")
        except Exception as e:
            st.error(f"Error: {e}")
