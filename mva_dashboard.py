# ======================================================
# MVA PAYROLL DASHBOARD — FINAL PRODUCTION VERSION
# ======================================================

import streamlit as st
import pandas as pd
from datetime import datetime
from pymongo import MongoClient
from compute_monthly_payroll import compute_monthly_payroll
from mongo_helpers import get_db


# ======================================================
# 1️⃣ DATABASE CONNECTION
# ======================================================
client = MongoClient("mongodb://localhost:27017/")
db = client["mva_payroll"]
employees_coll = db["employees"]
payroll_coll = db["monthly_payroll_summary"]
shift_coll = db["shift_master"]

# ======================================================
# 2️⃣ STREAMLIT CONFIG
# ======================================================
st.set_page_config(
    page_title="MVA Payroll Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ======================================================
# 3️⃣ SIDEBAR NAVIGATION
# ======================================================
st.sidebar.title("📊 MVA Payroll System")
page = st.sidebar.radio(
    "Navigate to",
    ["🏠 Dashboard", "📤 Upload Monthly Attendance", "👥 Employees"],
)

# ======================================================
# 4️⃣ 🏠 DASHBOARD PAGE
# ======================================================
if page == "🏠 Dashboard":
    st.title("🏠 Payroll Overview Dashboard")

    # Fetch most recent payroll record
    latest_doc = payroll_coll.find_one(sort=[("_id", -1)])
    if not latest_doc:
        st.info("No payroll data found yet. Please upload attendance to generate payroll.")
    else:
        latest_month = latest_doc.get("month", "Unknown")
        st.markdown(f"### 📅 Latest Computed Month: **{latest_month}**")

        # Fetch all payroll records for that month
        month_records = list(payroll_coll.find({"month": latest_month}, {"_id": 0}))
        if not month_records:
            st.warning("No records found for the latest month.")
        else:
            df = pd.DataFrame(month_records)
            total_employees = len(df)
            st.metric("👥 Employees Processed", total_employees)

            # Summary metrics
            fm_plus = df[df["full_month_status"] == "FM+1"].shape[0]
            fm = df[df["full_month_status"] == "FM"].shape[0]
            fm_minus = df[df["full_month_status"].str.startswith("FM-")].shape[0]
            half_month = df[df["half_days"] > 0].shape[0]
            total_absent = df["absent_days"].sum()
            total_leave = df["leaves"].sum()

            c1, c2, c3, c4, c5, c6 = st.columns(6)
            c1.metric("✅ FM+1", fm_plus)
            c2.metric("💼 FM", fm)
            c3.metric("⚠️ Partial (FM–x)", fm_minus)
            c4.metric("🌓 Half Month", half_month)
            c5.metric("🚫 Absents", total_absent)
            c6.metric("🏖️ Leaves", total_leave)

           # --- Visualization ---
            st.divider()
            st.subheader("🥧 Full Month Status Distribution")

            import matplotlib.pyplot as plt

            fm_counts = df["full_month_status"].value_counts().reset_index()
            fm_counts.columns = ["Status", "Count"]

            # Group smaller categories into 'Others'
            top_n = 7  # show top 5 statuses
            if len(fm_counts) > top_n:
                top = fm_counts.head(top_n)
                others_sum = fm_counts["Count"].iloc[top_n:].sum()
                top.loc[len(top)] = ["Others", others_sum]
                fm_counts = top

            if not fm_counts.empty:
                fig, ax = plt.subplots(figsize=(2, 2))
                colors = plt.cm.Set3.colors[:len(fm_counts)]  # pleasant palette

                wedges, texts, autotexts = ax.pie(
                    fm_counts["Count"],
                    labels=fm_counts["Status"],
                    autopct="%1.1f%%",
                    startangle=90,
                    counterclock=False,
                    colors=colors,
                    textprops={'fontsize': 5}
                )

                # Title and legend
                ax.set_title("Payroll Status Breakdown", fontsize=14, fontweight="bold")
                ax.axis("equal")
                st.pyplot(fig)
            else:
                st.info("No data available for chart.")

        
            # --- Data Table ---
            st.divider()
            st.subheader("📋 Payroll Data Table")

            display_cols = [
                "employee_id", "employee_name", "department", "month",
                "present_days", "half_days", "leaves", "week_offs", "absent_days",
                "total_working_days", "full_month_status", "shift_type", "shift_in", "shift_out"
            ]
            df_display = df[display_cols] if all(c in df.columns for c in display_cols) else df
            st.dataframe(df_display, use_container_width=True)

            # --- Export Excel ---
            from io import BytesIO

            st.divider()
            st.subheader("📥 Export Payroll Data")

            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
                df_display.to_excel(writer, index=False, sheet_name="Payroll")

            excel_buffer.seek(0)

            st.download_button(
                label="📥 Download Payroll (Excel)",
                data=excel_buffer,
                file_name=f"Payroll_{latest_month}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# ======================================================
# 5️⃣ 📤 UPLOAD MONTHLY ATTENDANCE PAGE
# ======================================================
elif page == "📤 Upload Monthly Attendance":
    st.title("📤 Upload Monthly Attendance")

    st.markdown("""
    ### 🧾 Instructions:
    1. Select or confirm the **Payroll Month** from the dropdown.
    2. Upload the **Monthly Attendance Excel** file (e.g., `monthperformance...xlsx`).
    3. Shift details are fetched automatically from MongoDB (`shift_master`).
    4. The computed payroll will be saved to the database and visible on the dashboard.
    """)

    # --- Fetch existing months from DB ---
    db = get_db()
    existing_months = (
        db["monthly_payroll_summary"]
        .distinct("report_month")
    )
    existing_months = sorted([m for m in existing_months if m], reverse=True)

    # --- Month selection dropdown ---
    current_year = datetime.now().year
    month_options = [
        f"{current_year}-{m:02d}" for m in range(1, 13)
    ][::-1]  # show recent first

    default_month = existing_months[0] if existing_months else month_options[0]

    selected_month = st.selectbox(
        "📅 Select or Confirm Payroll Month",
        options=month_options,
        index=month_options.index(default_month)
        if default_month in month_options else 0
    )

    st.caption("If the uploaded Excel includes 'Report Month', it will override this selection automatically.")

    # --- File uploader ---
    attendance_file = st.file_uploader(
        "📂 Upload Monthly Attendance File (.xls / .xlsx)",
        type=["xls", "xlsx"]
    )

    # --- Process Payroll Button ---
    if st.button("🧮 Process Payroll"):
        if not attendance_file:
            st.error("Please upload the attendance file first.")
        else:
            with st.spinner("⏳ Processing payroll... Please wait..."):
                try:
                    att_path = f"uploads/attendance_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    with open(att_path, "wb") as f:
                        f.write(attendance_file.read())

                    # Pass selected month hint into function
                    result = compute_monthly_payroll(att_path)

                    st.success(f"✅ Payroll computed successfully for {result['processed']} employees!")

                    if "month" in result:
                        st.info(f"📅 Computed for month: **{result['month']}**")
                    else:
                        st.warning(f"⚠️ Could not auto-detect month; used selected month: {selected_month}")

                    # Auto-refresh dashboard
                    st.info("🔄 Updating dashboard...")
                    st.rerun()

                except Exception as e:
                    st.error(f"❌ Error during computation: {e}")


# ======================================================
# 6️⃣ 👥 EMPLOYEE MANAGEMENT PAGE (with delete option)
# ======================================================
elif page == "👥 Employees":
    st.title("👥 Employee Management")

    # Upload Excel
    st.subheader("📤 Bulk Upload Employees")
    emp_file = st.file_uploader("Upload Employee Excel (.xls / .xlsx)", type=["xls", "xlsx"])

    if emp_file:
        try:
            df = pd.read_excel(emp_file)
            required_cols = ["Employee_ID", "Employee_Name", "Department"]
            if not all(col in df.columns for col in required_cols):
                st.error(f"Excel must contain columns: {required_cols}")
            else:
                inserted = 0
                for _, row in df.iterrows():
                    emp_id = str(row["Employee_ID"]).strip()
                    name = str(row["Employee_Name"]).strip()
                    dept = str(row.get("Department", "")).strip()
                    shift_type = str(row.get("Shift_Type", "")).strip()
                    shift_in = str(row.get("Shift_In", "")).strip()
                    shift_out = str(row.get("Shift_Out", "")).strip()
                    weekoff = str(row.get("Weekoff", "")).strip()

                    employees_coll.update_one(
                        {"employee_id": emp_id},
                        {"$set": {
                            "name": name,
                            "department": dept,
                            "shift_type": shift_type,
                            "shift_in": shift_in,
                            "shift_out": shift_out,
                            "weekoff": weekoff,
                            "updated_at": datetime.utcnow()
                        }},
                        upsert=True
                    )
                    inserted += 1
                st.success(f"✅ Imported/updated {inserted} employees successfully!")
        except Exception as e:
            st.error(f"Error importing employees: {e}")

    st.divider()

    # Manual Add
    st.subheader("➕ Add Employee Manually")
    with st.form("add_emp_form"):
        col1, col2, col3 = st.columns(3)
        emp_id = col1.text_input("Employee ID")
        name = col2.text_input("Employee Name")
        dept = col3.text_input("Department")

        col4, col5, col6 = st.columns(3)
        shift_type = col4.text_input("Shift Type")
        shift_in = col5.text_input("Shift In (HH:MM)")
        shift_out = col6.text_input("Shift Out (HH:MM)")

        weekoff = st.text_input("Weekly Off", value="Sunday")
        submit = st.form_submit_button("💾 Save Employee")

        if submit:
            if emp_id and name:
                employees_coll.update_one(
                    {"employee_id": emp_id},
                    {"$set": {
                        "name": name,
                        "department": dept,
                        "shift_type": shift_type,
                        "shift_in": shift_in,
                        "shift_out": shift_out,
                        "weekoff": weekoff,
                        "updated_at": datetime.utcnow()
                    }},
                    upsert=True
                )
                st.success("✅ Employee added/updated successfully!")
                st.rerun()

            else:
                st.error("Employee ID and Name are required.")

    st.divider()
    st.subheader("📋 All Employees")

    data = list(employees_coll.find({}, {"_id": 0}))
    if data:
        df_emp = pd.DataFrame(data)
        st.dataframe(df_emp, use_container_width=True)

        # --- Delete Option ---
        st.subheader("🗑️ Delete Employee")
        emp_ids = [d.get("employee_id") for d in data if d.get("employee_id")]
        if emp_ids:
            emp_to_delete = st.selectbox("Select Employee to Delete", options=emp_ids)
            confirm_delete = st.checkbox("⚠️ Confirm delete this employee permanently")

            if st.button("Delete Employee"):
                if confirm_delete:
                    employees_coll.delete_one({"employee_id": emp_to_delete})
                    st.success(f"✅ Employee {emp_to_delete} deleted successfully!")
                    st.rerun()
                else:
                    st.warning("Please confirm before deleting an employee.")
        else:
            st.info("No employees found in database.")
    else:
        st.info("No employees found yet.")

