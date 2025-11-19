import sqlite3

DB_FILE = "mva_payroll.db"

def get_db_connection():
    """Creates a connection to the SQLite database."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row  # This allows accessing columns by name
    return conn

def initialize_db():
    """
    Initializes the database. Creates the employees table if it doesn't exist,
    and adds columns for day-specific shifts if they are missing.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Create the table with basic fields if it doesn't exist
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS employees (
        employee_id TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        department TEXT,
        shift_type TEXT,
        shift_in TEXT,      -- General shift_in as a fallback
        shift_out TEXT,     -- General shift_out as a fallback
        weekoff TEXT,
        default_shift TEXT,
        updated_at TEXT
    )
    """)

    # --- Add day-specific columns if they don't exist ---
    # Get the list of existing columns
    cursor.execute("PRAGMA table_info(employees)")
    existing_columns = [row['name'] for row in cursor.fetchall()]

    # Define the columns to add for each day
    days_of_week = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
    for day in days_of_week:
        for shift_time in ["in", "out"]:
            column_name = f"{day}_{shift_time}"
            if column_name not in existing_columns:
                # Use ALTER TABLE to add the new column, with a default value
                cursor.execute(f"ALTER TABLE employees ADD COLUMN {column_name} TEXT DEFAULT '--:--'")
    
    # --- Create shift_master table ---
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS shift_master (
        shift_code TEXT PRIMARY KEY,
        description TEXT,
        intervals TEXT
    )
    """)

    # --- Create employee_shift_assignments table ---
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS employee_shift_assignments (
        employee_id TEXT PRIMARY KEY,
        shift_code TEXT
    )
    """)
    
    # --- Create a default shift if not exists ---
    cursor.execute("SELECT 1 FROM shift_master WHERE shift_code = 'GEN'")
    if cursor.fetchone() is None:
        cursor.execute("""
            INSERT INTO shift_master (shift_code, description, intervals)
            VALUES ('GEN', 'General Shift', '[{"start": "09:30", "end": "18:30"}]')
        """)

    conn.commit()
    conn.close()

# Initialize the database when this module is first imported
initialize_db()
