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
# 1️⃣ DATABASE INITIALIZATION & CONNECTION
# ======================================================
# --- MongoDB Connection ---
client = MongoClient("mongodb://localhost:27017/")
db = client["mva_payroll"]
payroll_coll = db["monthly_payroll_summary"]
shift_coll = db["shift_master"]
weekly_shifts_coll = db["weekly_shifts"]
employees_coll = db["employees"]

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
    ["🏠 Dashboard", "📤 Upload Monthly Attendance", "👥 Employee Management"],
)

# ======================================================
# 4️⃣ 🏠 DASHBOARD PAGE
# ======================================================
if page == "🏠 Dashboard":
    st.title("🏠 Payroll Overview Dashboard")

    # Fetch all unique months from the database
    all_months = list(payroll_coll.distinct("month"))
    # Filter out empty month values and sort in descending order (most recent first)
    all_months = [month for month in all_months if month and month.strip() != ""]
    all_months.sort(reverse=True)
    
    if not all_months:
        st.info("No payroll data found yet. Please upload attendance to generate payroll.")
    else:
        # Month selection dropdown
        selected_month = st.selectbox(
            "📅 Select Month to View Payroll Data",
            options=all_months,
            index=0
        )
        
        st.markdown(f"### 📅 Selected Month: **{selected_month}**")

        # Fetch all payroll records for the selected month
        month_records = list(payroll_coll.find({"month": selected_month}, {"_id": 0}))
        if not month_records:
            st.warning(f"No records found for the month: {selected_month}")
        else:
            df = pd.DataFrame(month_records)
            total_employees = len(df)
            st.metric("👥 Employees Processed", total_employees)

            # Summary metrics
            fm = df[df["full_month_status"] == "FM"].shape[0]
            fm_minus = df[df["full_month_status"].str.startswith("FM-")].shape[0]
            half_month = df[df["half_days"] > 0].shape[0]
            total_absent = df["absent_days"].sum()
            total_leave = df["leaves"].sum()

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("💼 FM", fm)
            c2.metric("⚠️ Partial (FM–x)", fm_minus)
            c3.metric("🌓 Half Month", half_month)
            c4.metric("🚫 Absents", total_absent)
            c5.metric("🏖️ Leaves", total_leave)

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
                file_name=f"Payroll_{selected_month}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# ======================================================
# 5️⃣ 📤 UPLOAD MONTHLY ATTENDANCE PAGE
# ======================================================
elif page == "📤 Upload Monthly Attendance":
    st.title("📤 Upload Monthly Attendance")

    # --- Download Timetable Section ---
    st.subheader("📥 Download Timetable Template")
    st.markdown("Download a timetable template with default shift times for the selected month:")

    # Month selection for the template
    current_year = datetime.now().year
    month_options = [
        f"{current_year}-{m:02d}" for m in range(1, 13)
    ][::-1]  # show recent first
    
    template_month = st.selectbox(
        "📅 Select Month for Template",
        options=month_options,
        index=0  # default to current or previous month
    )

    if st.button("📥 Generate and Download Timetable"):
        with st.spinner("⏳ Generating timetable template..."):
            try:
                import calendar
                import math
                from io import BytesIO
                import xlsxwriter
                from datetime import datetime
                import os

                # Parse selected month
                year, month = map(int, template_month.split('-'))
                month_name = calendar.month_name[month]
                num_days = calendar.monthrange(year, month)[1]

                # Use the updated timetable.py to generate a CSV first
                from timetable import generate_timetable
                generate_timetable(month, year)
                
                # Load the generated CSV to create the Excel template
                timetable_file = f"timetable_{year}_{str(month).zfill(2)}.csv"
                if not os.path.exists(timetable_file):
                    st.warning("No timetable file generated.")
                    st.stop()
                
                df_timetable = pd.read_csv(timetable_file)
                
                # Create Excel file in memory
                excel_buffer = BytesIO()
                workbook = xlsxwriter.Workbook(excel_buffer, {'in_memory': True})
                worksheet = workbook.add_worksheet(f"{month_name}_{year}")

                # --- Define Formats ---
                header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
                cell_format = workbook.add_format({'align': 'center', 'border': 1})
                emp_info_format = workbook.add_format({'bold': True, 'align': 'left'})

                # --- Write Headers ---
                worksheet.merge_range('D1:F1', 'MACRO VISION ACADEMY', header_format)
                worksheet.write('H1', 'Report Month', header_format)
                worksheet.write('I1', f"{month_name}-{year}", header_format)

                # --- Main Loop for Each Employee (from the generated timetable) ---
                current_row = 3
                for idx, emp_row in df_timetable.iterrows():
                    # --- Employee Info Header ---
                    worksheet.write(f'A{current_row}', 'Dept. Name', emp_info_format)
                    worksheet.write(f'B{current_row}', emp_row.get("department", "Default"))
                    worksheet.write(f'A{current_row + 1}', 'Empcode', emp_info_format)
                    worksheet.write(f'B{current_row + 1}', emp_row.get("employee_id", "0000"))
                    worksheet.write(f'D{current_row + 1}', 'Name', emp_info_format)
                    worksheet.write(f'E{current_row + 1}', emp_row.get("name", "Default"))

                    # --- Date and Day Rows ---
                    date_row = current_row + 2
                    day_row = current_row + 3
                    
                    worksheet.write(f'A{date_row}', '') # Empty cell before dates
                    worksheet.write(f'A{day_row}', '') # Empty cell before days

                    for day in range(1, num_days + 1):
                        # Date Number
                        worksheet.write(date_row, day, str(day), header_format)
                        # Day Name
                        date_obj = datetime(year, month, day)
                        day_name = date_obj.strftime('%a')
                        worksheet.write(day_row, day, day_name, header_format)

                    # --- Data Rows (IN, OUT, etc.) ---
                    data_labels = ["IN", "OUT", "Status"]
                    data_start_row = current_row + 4

                    for i, label in enumerate(data_labels):
                        row_num = data_start_row + i
                        worksheet.write(row_num, 0, label, header_format) # Write label in first column
                        
                        for day_col in range(1, num_days + 1):
                            date_obj = datetime(year, month, day_col)
                            current_day_name = date_obj.strftime('%A').lower() # e.g., "tuesday"

                            content = ""
                            # Get the shift details from the timetable data
                            date_key = f"{day_col:02d}-{date_obj.strftime('%a')}"  # e.g., "01-Mon"
                            shift_info = emp_row.get(date_key, "--:-- to --:-- (Off)")
                            
                            if label == "IN":
                                # Extract IN time from the shift_info string (e.g., "09:30 to 18:30 (General)")
                                try:
                                    in_out = shift_info.split(" to ")
                                    in_time = in_out[0] if in_out else "--:--"
                                    # Further split to remove shift type if needed
                                    in_time = in_time.split(" (")[0] if " (" in in_time else in_time
                                    content = in_time
                                except:
                                    content = "--:--"
                            elif label == "OUT":
                                # Extract OUT time from the shift_info string
                                try:
                                    in_out = shift_info.split(" to ")
                                    out_with_type = in_out[1] if len(in_out) > 1 else "--:--"
                                    # Remove shift type in parentheses
                                    out_time = out_with_type.split(" (")[0] if " (" in out_with_type else out_with_type
                                    content = out_time
                                except:
                                    content = "--:--"
                            elif label == "Status":
                                content = "A"  # Default to Absent, will be filled by user

                            # Final check for empty or invalid content
                            if content is None or (isinstance(content, float) and math.isnan(content)):
                                content = "--:--"
                                
                            worksheet.write(row_num, day_col, content, cell_format)
                    
                    # Move to the next employee block
                    current_row += len(data_labels) + 4 # 4 for header rows + space

                workbook.close()
                excel_buffer.seek(0)

                st.download_button(
                    label="✅ Click to Download Timetable",
                    data=excel_buffer,
                    file_name=f"timetable_template_{template_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.success(f"✅ Timetable template for {template_month} generated successfully!")

            except Exception as e:
                st.error(f"❌ Error generating timetable template: {e}")

    st.divider()

    # --- Payroll Calculation Section ---
    st.subheader("🧮 Calculate Payroll")
    st.markdown("Upload both your monthly timetable and attendance file to calculate payroll:")

    # --- File uploaders ---
    timetable_file = st.file_uploader(
        "📅 Upload Monthly Timetable File (.xls / .xlsx)",
        type=["xls", "xlsx"],
        key="timetable_upload"
    )
    
    attendance_file = st.file_uploader(
        "📂 Upload Monthly Attendance File (.xls / .xlsx)",
        type=["xls", "xlsx"],
        key="attendance_upload"
    )

    # --- Process Payroll Button ---
    if st.button("🧮 Process Payroll"):
        if not timetable_file or not attendance_file:
            st.error("Please upload both the timetable and attendance files.")
        else:
            with st.spinner("⏳ Processing payroll... Please wait..."):
                try:
                    # Save uploaded files temporarily
                    import os
                    timetable_path = f"uploads/timetable_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    attendance_path = f"uploads/attendance_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    
                    os.makedirs("uploads", exist_ok=True)
                    
                    with open(timetable_path, "wb") as f:
                        f.write(timetable_file.read())
                    
                    with open(attendance_path, "wb") as f:
                        f.write(attendance_file.read())

                    # Process the payroll with both files
                    result = compute_monthly_payroll(attendance_path, timetable_path)

                    st.success(f"✅ Payroll computed successfully for {result['processed']} employees!")
                    st.info("📊 Check the Dashboard page for results.")

                    # Auto-refresh dashboard
                    st.rerun()

                except Exception as e:
                    st.error(f"❌ Error during computation: {e}")

# ======================================================
# 6️⃣ 👥 EMPLOYEE MANAGEMENT PAGE
# ======================================================
elif page == "👥 Employee Management":
    st.title("👥 Employee & Shift Management")
    
    # Define helper functions
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

    def load_weekly_shift_file(file):
        """
        Uploads master shift details for employees.
        Expects an Excel file with columns: Employee_ID, Employee_Name, Department, Days, Shift_Type, Shift_In, Shift_Out, Crosses_Midnight, Shift_Duration_Hours, Weekoff
        """
        df = pd.read_excel(file)
        df = df.fillna("")

        # Group by Employee_ID to handle multiple rows per employee (one per day/shift type)
        grouped = df.groupby("Employee_ID")

        for emp_id, group in grouped:
            emp_name = group["Employee_Name"].iloc[0]
            dept = group["Department"].iloc[0]

            week_data = []
            for _, row in group.iterrows():
                week_data.append({
                    "day": row["Days"],
                    "shift_type": row["Shift_Type"],
                    "shift_in": row["Shift_In"],
                    "shift_out": row["Shift_Out"],
                    "crosses_midnight": row["Crosses_Midnight"],
                    "shift_duration_hours": row["Shift_Duration_Hours"],
                    "weekoff": row["Weekoff"]
                })

            # Store the weekly shift data for this employee in MongoDB
            db.weekly_shifts.update_one(
                {"employee_id": str(emp_id)},
                {
                    "$set": {
                        "employee_id": str(emp_id),
                        "employee_name": emp_name,
                        "department": dept,
                        "week_data": week_data,
                        "uploaded_at": datetime.utcnow()
                    }
                },
                upsert=True
            )

        return True

    def add_employee(emp_id, name, department):
        """Add a new employee to MongoDB"""
        # Check if employee already exists
        existing_employee = db.employees.find_one({"employee_id": emp_id})
        if existing_employee:
            return False, "Employee ID already exists"
        
        # Insert new employee
        employee_doc = {
            "employee_id": emp_id,
            "name": name,
            "department": department,
            "updated_at": datetime.utcnow().isoformat()
        }
        
        result = db.employees.insert_one(employee_doc)
        return True, "Employee added successfully"

    def delete_employee(emp_id):
        """Delete an employee from MongoDB"""
        result = db.employees.delete_one({"employee_id": emp_id})
        
        if result.deleted_count > 0:
            # Also delete from MongoDB if needed
            db.weekly_shifts.delete_one({"employee_id": str(emp_id)})
            return True, "Employee deleted successfully"
        else:
            return False, "Employee not found"

    def get_all_employees():
        """Get all employees from MongoDB"""
        employees = list(db.employees.find({}))
        return employees

    # Create tabs for different functionalities
    tab1, tab2, tab3 = st.tabs(["Add Employee", "Delete Employee", "Upload Shifts"])

    with tab1:
        st.header("Add New Employee")
        new_emp_id = st.text_input("Employee ID", key="new_emp_id_2")
        new_name = st.text_input("Employee Name", key="new_name_2")
        new_dept = st.text_input("Department", key="new_dept_2")
        
        if st.button("Add Employee", key="add_employee_2"):
            if new_emp_id and new_name and new_dept:
                success, message = add_employee(new_emp_id, new_name, new_dept)
                if success:
                    st.success(message)
                else:
                    st.error(message)
            else:
                st.error("Please fill in all fields")

    with tab2:
        st.header("Delete Employee")
        all_employees = get_all_employees()
        emp_options = {f"{emp['employee_id']} - {emp['name']}": emp['employee_id'] for emp in all_employees}
        
        selected_emp = st.selectbox("Select Employee to Delete", options=[""] + list(emp_options.keys()), key="delete_emp_select_2")
        
        if selected_emp:
            emp_to_delete = emp_options[selected_emp]
            confirm_delete = st.checkbox("⚠️ Confirm delete this employee permanently", key="confirm_delete_2")
            
            if st.button("Delete Employee", key="delete_employee_btn_2"):
                if confirm_delete:
                    success, message = delete_employee(emp_to_delete)
                    if success:
                        st.success(message)
                        st.rerun()
                    else:
                        st.error(message)
                else:
                    st.warning("Please confirm before deleting an employee.")

    with tab3:
        st.header("Upload Master Shift Details")
        uploaded_file = st.file_uploader("Upload Staff Shift Excel File", type=["xlsx"])

        if st.button("Upload and Store Weekly Shifts", key="upload_shifts_2"):
            if not uploaded_file:
                st.error("Please upload the Excel file")
            else:
                try:
                    load_weekly_shift_file(uploaded_file)
                    st.success("✅ Weekly shift data stored successfully in MongoDB!")
                except Exception as e:
                    st.error(f"Error: {e}")

    # Display all employees
    st.divider()
    st.header("📋 All Employees")
    employees_df = pd.DataFrame(get_all_employees())
    if not employees_df.empty:
        st.dataframe(employees_df, use_container_width=True)
    else:
        st.info("No employees found in database.")
