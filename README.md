# Macro Vision Academy Attendance Management System

This is a Streamlit-based web application designed to automate the processing and management of attendance data for Macro Vision Academy (MVA). It allows users to upload Excel attendance files, calculates various attendance metrics, and stores the results in a local SQLite database.

## Features

*   **Excel File Upload:** Supports `.xlsx` and `.xls` formats for attendance data.
*   **Attendance Calculation:** Processes MVA-specific Excel formats to calculate:
    *   Present Days
    *   Half Days
    *   Leaves
    *   Week Offs
    *   Absent Days
    *   Total Overtime (OT) Hours
    *   FM Status (a custom attendance status based on specific rules)
*   **Database Storage:** Stores uploaded files and calculated attendance results in an `attendance_system.db` SQLite database.
*   **Recent Calculations Display:** Shows a list of the most recent attendance calculations directly in the application.
*   **Results Download:** Allows users to download the processed attendance results as an Excel file.

## Project Structure

*   `app.py`: The main Streamlit application file, handling UI, file uploads, database interactions, and calling the attendance processing logic.
*   `mva.py`: Contains the core logic for parsing the MVA Excel attendance sheet format and calculating attendance metrics.
*   `extract_info.py`: A utility script to extract department and month information from Excel files.
*   `attendance_system.db`: SQLite database file (generated on first run) to store uploaded files and attendance results.
*   `requirements.txt`: Lists all Python dependencies required to run the application.
*   `uploads/`: Directory for example or temporary uploaded files.

## Setup

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/Sri174/PayrollMva.git
    cd PayrollMva
    ```

2.  **Install dependencies:**
    Ensure you have Python installed (version 3.8+ recommended). Then, install the required packages:
    ```bash
    pip install -r requirements.txt
    pip install "xlrd>=2.0.1"
    ```
    *Note: `xlrd` is specifically required for reading older `.xls` Excel file formats.*

## Running the Application

To start the Streamlit application, navigate to the project's root directory in your terminal and execute:

```bash
streamlit run app.py
```

This command will launch the application, and you can access it in your web browser. Streamlit typically provides a local URL (e.g., `http://localhost:8501` or `http://localhost:8502`) where the app is running.

## Usage

1.  **Upload File:** On the main page, use the "Choose Excel file (.xlsx or .xls)" button to upload an attendance sheet.
2.  **View Raw Data:** (Optional) Expand the "View Raw Data" section to see the raw content of your uploaded Excel file.
3.  **Calculate Attendance:** Click the "Calculate Attendance" button to process the uploaded data.
4.  **Review Results:** The "Attendance Results" table will display the calculated metrics.
5.  **Download Results:** Use the "Download Results as Excel" button to save the processed data.
6.  **Recent Calculations:** The "Recent Calculations" section at the bottom shows previously processed attendance data.

## Database Schema

The `attendance_system.db` contains two main tables:

### `uploaded_files`
*   `id`: INTEGER PRIMARY KEY AUTOINCREMENT
*   `filename`: TEXT
*   `upload_date`: TEXT (ISO format)
*   `file_data`: BLOB (stores the raw binary data of the uploaded Excel file)

### `attendance_results`
*   `id`: INTEGER PRIMARY KEY AUTOINCREMENT
*   `employee_id`: TEXT
*   `employee_name`: TEXT
*   `department`: TEXT
*   `month`: TEXT
*   `present_days`: REAL
*   `half_days`: INTEGER
*   `leaves`: INTEGER
*   `week_offs`: INTEGER
*   `absent_days`: INTEGER
*   `total_ot_hours`: REAL (stores total overtime hours)
*   `fm_status`: TEXT (calculated FM status)
*   `calculation_date`: TEXT (ISO format)
