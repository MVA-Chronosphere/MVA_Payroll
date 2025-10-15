# app.py
import streamlit as st
import pandas as pd
import sqlite3
import io
from datetime import datetime
import mva

# Initialize databases
def init_databases():
    conn = sqlite3.connect('attendance_system.db')
    c = conn.cursor()
    
    # Files table
    c.execute('''CREATE TABLE IF NOT EXISTS uploaded_files
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  filename TEXT,
                  upload_date TEXT,
                  file_data BLOB)''')
    
    # Drop and recreate the table to remove total_days column
    c.execute("DROP TABLE IF EXISTS attendance_results")
    c.execute('''CREATE TABLE attendance_results
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  employee_id TEXT,
                  employee_name TEXT,
                  department TEXT,
                  month TEXT,
                  present_days REAL,
                  half_days INTEGER,
                  leaves INTEGER,
                  week_offs INTEGER,
                  absent_days INTEGER,
                  work_hours REAL,
                  fm_status TEXT,
                  calculation_date TEXT)''')
    
    conn.commit()
    conn.close()

# Database functions
def save_uploaded_file(filename, file_data):
    conn = sqlite3.connect('attendance_system.db')
    c = conn.cursor()
    c.execute("INSERT INTO uploaded_files (filename, upload_date, file_data) VALUES (?, ?, ?)",
              (filename, datetime.now().isoformat(), file_data))
    conn.commit()
    file_id = c.lastrowid
    conn.close()
    return file_id

def update_database_schema():
    """Update the database schema to include total_ot_hours if it doesn't exist"""
    conn = sqlite3.connect('attendance_system.db')
    c = conn.cursor()
    
    # Check if total_ot_hours column exists
    c.execute("PRAGMA table_info(attendance_results)")
    columns = [column[1] for column in c.fetchall()]
    
    if 'total_ot_hours' not in columns:
        try:
            # Add the new column
            c.execute('''ALTER TABLE attendance_results 
                        ADD COLUMN total_ot_hours REAL DEFAULT 0.0''')
            conn.commit()
        except sqlite3.OperationalError:
            # Column might already exist in some form
            pass
    
    conn.close()

def save_results(results_df):
    # First ensure the database schema is up to date
    update_database_schema()
    
    conn = sqlite3.connect('attendance_system.db')
    c = conn.cursor()
    
    # Ensure 'Total OT Hours' column exists
    if 'Total OT Hours' not in results_df.columns:
        results_df['Total OT Hours'] = 0.0
    
    # Convert any datetime objects to strings to avoid Arrow serialization issues
    results_df = results_df.copy()
    for col in results_df.columns:
        if results_df[col].dtype == 'object':
            results_df[col] = results_df[col].astype(str)
    
    for _, row in results_df.iterrows():
        c.execute('''INSERT INTO attendance_results 
                     (employee_id, employee_name, department, month,
                      present_days, half_days, leaves, week_offs,
                      absent_days, total_ot_hours, fm_status, calculation_date)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                  (row['Employee ID'], row['Employee Name'], row['Department'],
                   row['Month'], row['Present Days'], row['Half Days'], 
                   row['Leaves'], row['Week Offs'], row['Absent Days'], 
                   float(row['Total OT Hours']), row['FM Status'], datetime.now().isoformat()))
    
    conn.commit()
    conn.close()

def get_recent_results():
    conn = sqlite3.connect('attendance_system.db')
    df = pd.read_sql_query("SELECT * FROM attendance_results ORDER BY calculation_date DESC LIMIT 20", conn)
    conn.close()
    return df

# Streamlit UI
def main():
    st.set_page_config(page_title="MVA Attendance Management System", layout="wide")
    init_databases()
    
    st.title("📊 Macro Vision Academy Attendance Management System")
    
    # File upload section
    st.subheader("Upload Attendance File")
    uploaded_file = st.file_uploader("Choose Excel file (.xlsx or .xls)", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            # Verify file extension
            file_extension = uploaded_file.name.split('.')[-1].lower()
            if file_extension not in ['xlsx', 'xls']:
                st.error("Invalid file type. Please upload an Excel file (.xlsx or .xls)")
                return
                
            # Read file content
            file_bytes = uploaded_file.read()
            
            # Try to read the Excel file with appropriate engine
            try:
                engine = 'openpyxl' if file_extension == 'xlsx' else 'xlrd'
                df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine=engine)
                
                # If we get here, file is valid - save to database
                file_id = save_uploaded_file(uploaded_file.name, file_bytes)
                if file_id:
                    st.success(f"File uploaded successfully! (ID: {file_id})")
                else:
                    st.error("Failed to save file to database")
                    return
                    
            except Exception as e:
                st.error(f"Error reading Excel file: {str(e)}\n\n"
                        f"Please ensure the file is a valid Excel file and not corrupted.")
                if file_extension == 'xls':
                    st.info("For .xls files, please ensure they are in Excel 97-2003 format.")
                return
                
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            return
        
        # Display raw data
        with st.expander("View Raw Data"):
            st.dataframe(df)
        
        # Process data
        if st.button("Calculate Attendance"):
            with st.spinner("Processing attendance data..."):
                try:
                    # Parse MVA format
                    blocks = mva.parse_mva_blocks(df)
                    
                    if not blocks:
                        st.warning("No employee data found in the file. Please check the file format.")
                        return
                    
                    # Process attendance data
                    result_df = mva.process_attendance_data(blocks)
                    
                    # Display results
                    st.subheader("Attendance Results")
                    st.dataframe(result_df)
                    
                    # Save results to database
                    save_results(result_df)
                    st.success("Results saved to database!")
                    
                    # Download button
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        result_df.to_excel(writer, index=False, sheet_name='Attendance Results')
                    output.seek(0)
                    
                    st.download_button(
                        label="Download Results as Excel",
                        data=output,
                        file_name=f"mva_attendance_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"Error processing: {e}")
    
    # Recent results section
    st.subheader("Recent Calculations")
    recent_results = get_recent_results()
    if not recent_results.empty:
        st.dataframe(recent_results)
    else:
        st.info("No previous calculations found")

if __name__ == "__main__":
    main()