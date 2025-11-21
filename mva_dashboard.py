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

        
            # --- Filter Options ---
            st.divider()
            st.subheader("🔍 Filter Options")
            
            # Create columns for filter options
            filter_col1, filter_col2 = st.columns(2)
            
            with filter_col1:
                filter_type = st.selectbox(
                    "Select Filter Type",
                    ["All Records", "ABV/AAV (Absent Before/After Vacation)", "PIV (Present In Vacation)"],
                    key="filter_type"
                )
            
            # Check if there are any vacation dates in the database records for this month
            vacation_info = payroll_coll.find_one({"month": selected_month, "vacation_start_date": {"$exists": True}}, {"vacation_start_date": 1, "vacation_end_date": 1})
            if vacation_info and "vacation_start_date" in vacation_info and "vacation_end_date" in vacation_info:
                # Parse the ISO format date strings back to date objects
                import json
                from datetime import date
                try:
                    vacation_start_date = datetime.fromisoformat(vacation_info["vacation_start_date"].replace("Z", "+00:00")).date()
                    vacation_end_date = datetime.fromisoformat(vacation_info["vacation_end_date"].replace("Z", "+00:00")).date()
                except:
                    # Fallback: try different date formats
                    vacation_start_date = vacation_info["vacation_start_date"]
                    vacation_end_date = vacation_info["vacation_end_date"]
            else:
                # Try to get vacation dates from any record in the month
                sample_record = payroll_coll.find_one({"month": selected_month, "vacation_start_date": {"$exists": True}})
                if sample_record:
                    # Check if vacation dates exist in the record
                    vacation_start_str = sample_record.get("vacation_start_date")
                    vacation_end_str = sample_record.get("vacation_end_date")
                    if vacation_start_str and vacation_end_str:
                        try:
                            vacation_start_date = datetime.fromisoformat(vacation_start_str.replace("Z", "+00:00")).date()
                            vacation_end_date = datetime.fromisoformat(vacation_end_str.replace("Z", "+00:00")).date()
                        except:
                            vacation_start_date = vacation_start_str
                            vacation_end_date = vacation_end_str
                else:
                    vacation_start_date = None
                    vacation_end_date = None
            
            with filter_col2:
                if filter_type != "All Records":
                    if vacation_start_date and vacation_end_date:
                        st.info(f"Using vacation period: {vacation_start_date} to {vacation_end_date}")
                    else:
                        st.warning("No vacation period specified. Please upload attendance with vacation dates to use these filters.")
            
            # --- Data Table ---
            st.divider()
            st.subheader("📋 Payroll Data Table")

            # Apply filters based on selection
            df_display = df.copy()
            
            if filter_type == "ABV/AAV (Absent Before/After Vacation)":
                if vacation_start_date and vacation_end_date:
                    # Find employees who were absent before or after the vacation period
                    try:
                        # Convert to datetime for comparison
                        vacation_start = pd.to_datetime(vacation_start_date)
                        vacation_end = pd.to_datetime(vacation_end_date)
                        
                        # Get daily attendance records from the database
                        daily_attendance = list(db.attendance_daily.find({
                            "month": selected_month
                        }))
                        
                        daily_df = pd.DataFrame(daily_attendance)
                        
                        if not daily_df.empty:
                            # Convert day to datetime for comparison
                            daily_df['date'] = daily_df.apply(
                                lambda row: pd.to_datetime(f"{vacation_start.year}-{vacation_start.month}-{row['day']}"), axis=1
                            )
                            
                            # Find employees absent before vacation (day before vacation start)
                            day_before_vacation = vacation_start - pd.Timedelta(days=1)
                            day_after_vacation = vacation_end + pd.Timedelta(days=1)
                            
                            # Filter for days before and after vacation
                            before_vacation_records = daily_df[
                                (daily_df['date'] == day_before_vacation) & 
                                (daily_df['status'].isin(['A', 'ABSENT']))
                            ]
                            
                            after_vacation_records = daily_df[
                                (daily_df['date'] == day_after_vacation) & 
                                (daily_df['status'].isin(['A', 'ABSENT']))
                            ]
                            
                            # Get employee IDs who were absent before or after vacation
                            absent_employees = set()
                            absent_employees.update(before_vacation_records['employee_id'].tolist())
                            absent_employees.update(after_vacation_records['employee_id'].tolist())
                            
                            # Filter the main dataframe to show only these employees
                            df_display = df[df['employee_id'].isin(list(absent_employees))]
                            
                    except Exception as e:
                        st.error(f"Error applying ABV/AAV filter: {e}")
                        
            elif filter_type == "PIV (Present In Vacation)":
                if vacation_start_date and vacation_end_date:
                    # Find employees who were present during the vacation period
                    try:
                        # Convert to datetime for comparison
                        vacation_start = pd.to_datetime(vacation_start_date)
                        vacation_end = pd.to_datetime(vacation_end_date)
                        
                        # Get daily attendance records from the database
                        daily_attendance = list(db.attendance_daily.find({
                            "month": selected_month
                        }))
                        
                        daily_df = pd.DataFrame(daily_attendance)
                        
                        if not daily_df.empty:
                            # Parse selected_month to extract year and month number
                            # Handle various formats: 2025-10, 2025-October, October-2025, October 2025, October, 10
                            import calendar
                            from datetime import datetime

                            sm = str(selected_month).strip()
                            selected_year = None
                            selected_month_num = None

                            try:
                                if '-' in sm:
                                    left, right = sm.split('-', 1)
                                    left = left.strip(); right = right.strip()
                                    if left.isdigit():
                                        selected_year = int(left); month_part = right
                                    elif right.isdigit():
                                        selected_year = int(right); month_part = left
                                    else:
                                        month_part = sm.replace('-', ' ')
                                        selected_year = vacation_start.year if vacation_start else datetime.now().year
                                else:
                                    if ' ' in sm:
                                        a, b = sm.split(' ', 1)
                                        a = a.strip(); b = b.strip()
                                        if a.isdigit():
                                            selected_year = int(a); month_part = b
                                        elif b.isdigit():
                                            selected_year = int(b); month_part = a
                                        else:
                                            month_part = sm
                                            selected_year = vacation_start.year if vacation_start else datetime.now().year
                                    else:
                                        month_part = sm
                                        selected_year = vacation_start.year if vacation_start else datetime.now().year

                                mp = month_part.strip()
                                if mp.isdigit():
                                    selected_month_num = int(mp)
                                else:
                                    try:
                                        selected_month_num = list(calendar.month_name).index(mp.capitalize())
                                    except:
                                        try:
                                            selected_month_num = list(calendar.month_abbr).index(mp.capitalize()[:3])
                                        except:
                                            selected_month_num = vacation_start.month if vacation_start else datetime.now().month

                                if not (1 <= int(selected_month_num) <= 12):
                                    selected_month_num = vacation_start.month if vacation_start else datetime.now().month

                                date_fmt = f"{int(selected_year)}-{int(selected_month_num):02d}-%d"

                                def build_date(row):
                                    try:
                                        return pd.to_datetime(date_fmt.replace("%d", str(int(row["day"]))))
                                    except:
                                        return pd.NaT

                                daily_df["date"] = daily_df.apply(build_date, axis=1)

                            except Exception as e:
                                st.error(f"Error parsing selected month '{selected_month}': {e}")
                                now = datetime.now()
                                fallback_fmt = f"{now.year}-{now.month:02d}-%d"
                                daily_df["date"] = daily_df["day"].apply(lambda d: pd.to_datetime(fallback_fmt.replace("%d", str(int(d)))) if pd.notna(d) else pd.NaT)
                            
                            # Filter for days within the vacation period where status is P (Present) only
                            # Explicitly exclude WO (Week Off), W/O, and OFF from the PIV filter
                            vacation_period_present = daily_df[
                                (daily_df['date'] >= vacation_start) & 
                                (daily_df['date'] <= vacation_end) & 
                                (daily_df['status'] == 'P')
                            ]
                            
                            # Get employee IDs who were present during vacation
                            present_employees = set(vacation_period_present['employee_id'].tolist())
                            
                            # Filter the main dataframe to show only these employees
                            df_display = df[df['employee_id'].isin(list(present_employees))]
                            
                    except Exception as e:
                        st.error(f"Error applying PIV filter: {e}")

            display_cols = [
                "employee_id", "employee_name", "department", "month",
                "present_days", "half_days", "leaves", "week_offs", "absent_days",
                "total_working_days", "full_month_status", "shift_type", "shift_in", "shift_out"
            ]
            df_display = df_display[display_cols] if all(c in df_display.columns for c in display_cols) else df_display
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

    # Add vacation checkbox functionality
    vacation_checkbox = st.checkbox("🏖️ Include Vacation Days")
    vacation_start_date = None
    vacation_end_date = None
    
    if vacation_checkbox:
        col1, col2 = st.columns(2)
        with col1:
            vacation_start_date = st.date_input("Start Date", value=None, key="vacation_start")
        with col2:
            vacation_end_date = st.date_input("End Date", value=None, key="vacation_end")
    
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
                    result = compute_monthly_payroll(attendance_path, timetable_path, vacation_start_date, vacation_end_date)

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
        df = pd.read_excel(file, engine='openpyxl' if file.name.endswith('.xlsx') else 'xlrd')
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

            # Store the weekly shift data for this employee in MongoDB - REPLACE existing data
            db.weekly_shifts.replace_one(
                {"employee_id": str(emp_id)},
                {
                    "employee_id": str(emp_id),
                    "employee_name": emp_name,
                    "department": dept,
                    "week_data": week_data,
                    "uploaded_at": datetime.utcnow()
                },
                upsert=True
            )

        return True

    def add_or_update_employee_weekly_shifts(emp_id, name, department, week_data):
        """Add or update an employee's weekly shifts in MongoDB weekly_shifts collection"""
        # Prepare the document to insert/update
        employee_doc = {
            "employee_id": str(emp_id),
            "employee_name": name,
            "department": department,
            "week_data": week_data,
            "updated_at": datetime.utcnow().isoformat()
        }
        
        # Use upsert=True to insert if doesn't exist or update if exists
        result = db.weekly_shifts.update_one(
            {"employee_id": str(emp_id)},
            {"$set": employee_doc},
            upsert=True
        )
        
        if result.upserted_id or result.modified_count > 0:
            return True, "Employee weekly shifts added/updated successfully"
        else:
            return False, "No changes made to employee data"

    def delete_employee_weekly_shifts(emp_id):
        """Delete an employee from weekly_shifts collection in MongoDB"""
        result = db.weekly_shifts.delete_one({"employee_id": str(emp_id)})
        
        if result.deleted_count > 0:
            # Also delete from employees collection if needed
            db.employees.delete_one({"employee_id": str(emp_id)})
            return True, "Employee deleted successfully from weekly_shifts collection"
        else:
            return False, "Employee not found in weekly_shifts collection"

    def add_employee_weekly_shifts(emp_id, name, department, week_data=None):
        """Add a new employee with weekly shifts to MongoDB weekly_shifts collection"""
        # Prepare the document to insert/update
        employee_doc = {
            "employee_id": str(emp_id),
            "employee_name": name,
            "department": department,
            "week_data": week_data or [],
            "updated_at": datetime.utcnow().isoformat()
        }
        
        # Use upsert=True to insert if doesn't exist or update if exists
        result = db.weekly_shifts.update_one(
            {"employee_id": str(emp_id)},
            {"$set": employee_doc},
            upsert=True
        )
        
        if result.upserted_id or result.modified_count > 0:
            return True, "Employee added/updated successfully in weekly_shifts collection"
        else:
            return False, "No changes made to employee data"

    def delete_employee_weekly_shifts_only(emp_id):
        """Delete an employee from weekly_shifts collection in MongoDB"""
        result = db.weekly_shifts.delete_one({"employee_id": str(emp_id)})
        
        if result.deleted_count > 0:
            # Also delete from employees collection if needed
            db.employees.delete_one({"employee_id": str(emp_id)})
            return True, "Employee deleted successfully from weekly_shifts collection"
        else:
            return False, "Employee not found in weekly_shifts collection"

    def get_all_employees():
        """Get all employees from weekly_shifts collection"""
        employees = list(db.weekly_shifts.find({}))
        return employees

    # Create tabs for different functionalities
    tab1, tab2, tab3 = st.tabs(["Add Employee", "Delete Employee", "Upload Shifts"])

    with tab1:
        st.header("Add New Employee")
        new_emp_id = st.text_input("Employee ID", key="new_emp_id_2")
        new_name = st.text_input("Employee Name", key="new_name_2")
        new_dept = st.text_input("Department", key="new_dept_2")
        
        # Add fields for weekly shift data
        st.subheader("Weekly Shift Details")
        day = st.selectbox("Day", ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"], key="day_2")
        shift_type = st.text_input("Shift Type", key="shift_type_2")
        shift_in = st.text_input("Shift In (HH:MM)", key="shift_in_2")
        shift_out = st.text_input("Shift Out (HH:MM)", key="shift_out_2")
        crosses_midnight = st.checkbox("Crosses Midnight", key="crosses_midnight_2")
        shift_duration = st.number_input("Shift Duration (Hours)", key="shift_duration_2", min_value=0.0, step=0.5)
        weekoff = st.checkbox("Week Off", key="weekoff_2")
        
        if st.button("Add Employee", key="add_employee_2"):
            if new_emp_id and new_name and new_dept:
                # Create week_data from the input fields
                week_data = [{
                    "day": day,
                    "shift_type": shift_type,
                    "shift_in": shift_in,
                    "shift_out": shift_out,
                    "crosses_midnight": crosses_midnight,
                    "shift_duration_hours": shift_duration,
                    "weekoff": weekoff
                }]
                
                success, message = add_employee_weekly_shifts(new_emp_id, new_name, new_dept, week_data)
                if success:
                    st.success(message)
                else:
                    st.error(message)
            else:
                st.error("Please fill in all required fields (Employee ID, Name, Department)")

    with tab2:
        st.header("Delete Employee")
        all_employees = get_all_employees()
        if all_employees:
            emp_options = {f"{emp.get('employee_name', emp.get('Employee_Name', 'Unknown'))} (ID: {emp.get('employee_id', emp.get('Employee_ID', 'Unknown'))})": emp.get('employee_id', emp.get('Employee_ID', '')) for emp in all_employees}
            
            selected_emp = st.selectbox("Select Employee to Delete", options=[""] + list(emp_options.keys()), key="delete_emp_select_2")
            
            if selected_emp:
                emp_to_delete = emp_options[selected_emp]
                confirm_delete = st.checkbox("⚠️ Confirm delete this employee permanently", key="confirm_delete_2")
                
                if st.button("Delete Employee", key="delete_employee_btn_2"):
                    if confirm_delete:
                        success, message = delete_employee_weekly_shifts_only(emp_to_delete)
                        if success:
                            st.success(message)
                            st.rerun()
                        else:
                            st.error(message)
                    else:
                        st.warning("Please confirm before deleting an employee.")
        else:
            st.info("No employees available to delete.")

    with tab3:
        st.header("Upload Master Shift Details")
        uploaded_file = st.file_uploader("Upload Staff Shift Excel File", type=["xlsx","xls"])

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
    employees = get_all_employees()
    if employees:
        # Create a DataFrame with only the required columns
        df_data = []
        for emp in employees:
            # Create a row with only required fields
            row = {
                'Employee_ID': emp.get('employee_id', 'N/A'),
                'Employee_Name': emp.get('employee_name', 'N/A'),
                'Department': emp.get('department', 'N/A'),
                'Expected_hours': emp.get('expected_hours', 'N/A')  # This might not exist in current data
            }
            
            # Initialize day-specific shift in/out columns with default values
            days_of_week = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            for day in days_of_week:
                day_lower = day.lower()
                row[f"{day_lower}_in"] = "N/A"
                row[f"{day_lower}_out"] = "N/A"
            
            # Add week_data as separate shift in/out columns for each day
            week_data = emp.get('week_data', [])
            for day_data in week_data:
                if isinstance(day_data, dict):
                    day_name = day_data.get('day', 'Unknown')
                    day_lower = day_name.lower() if day_name != 'Unknown' else 'unknown'
                    if day_lower.capitalize() in days_of_week:
                        row[f"{day_lower}_in"] = day_data.get('shift_in', 'N/A')
                        row[f"{day_lower}_out"] = day_data.get('shift_out', 'N/A')
            
            df_data.append(row)
        
        employees_df = pd.DataFrame(df_data)
        
        # Display the DataFrame with only required columns
        st.dataframe(employees_df, use_container_width=True)
        
        # Also provide an option to see the raw data in JSON format
        if st.checkbox("Show raw employee data in JSON format"):
            for emp in employees:
                with st.expander(f"Employee: {emp.get('employee_name', 'N/A')} (ID: {emp.get('employee_id', 'N/A')})"):
                    st.json(emp)
    else:
        st.info("No employees found in database.")
