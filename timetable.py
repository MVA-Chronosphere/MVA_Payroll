import pandas as pd
from datetime import datetime
from pymongo import MongoClient

def get_mongo_connection():
    """Creates a connection to the MongoDB database."""
    client = MongoClient("mongodb://localhost:27017/")
    db = client["mva_payroll"]
    return db

def generate_timetable(month, year):
    """
    Generates a timetable for the given month and year based on master shift details in MongoDB
    and saves it to a CSV file.
    """
    print(f"📅 Generating timetable for {month}/{year}...")
    db = get_mongo_connection()
    
    try:
        # Get the number of days in the month
        num_days = pd.Period(f'{year}-{month}-01').days_in_month
        dates = pd.date_range(start=f'{year}-{month}-01', periods=num_days)

        # Fetch employee weekly shift data
        weekly_shifts_cursor = db.weekly_shifts.find({})
        employees_shifts = list(weekly_shifts_cursor)

        if not employees_shifts:
            print("No employee shift data found to generate timetable.")
            return

        timetable_records = []
        for emp_shift in employees_shifts:
            record = {
                "employee_id": emp_shift["employee_id"],
                "name": emp_shift["employee_name"],
                "department": emp_shift["department"]
            }
            
            # Create a mapping of day names to shift details for this employee
            day_shift_map = {}
            for day_info in emp_shift["week_data"]:
                day_name = day_info["day"].lower()  # e.g., "monday"
                day_shift_map[day_name] = {
                    "shift_in": day_info["shift_in"],
                    "shift_out": day_info["shift_out"],
                    "shift_type": day_info["shift_type"]
                }

            for date in dates:
                day_name = date.strftime('%A').lower()  # e.g., "monday"
                
                date_key = date.strftime('%d-%a')  # e.g., "01-Mon"
                
                # Get the shift details for this day from the employee's weekly pattern
                shift_info = day_shift_map.get(day_name, {"shift_in": "--:--", "shift_out": "--:--", "shift_type": "Off"})
                shift_in = shift_info["shift_in"]
                shift_out = shift_info["shift_out"]
                shift_type = shift_info["shift_type"]
                
                record[date_key] = f"{shift_in} to {shift_out} ({shift_type})"

            timetable_records.append(record)

        df = pd.DataFrame(timetable_records)
        output_filename = f"timetable_{year}_{str(month).zfill(2)}.csv"
        df.to_csv(output_filename, index=False)
        print(f"✅ Timetable saved to {output_filename}")

    except Exception as e:
        print(f"❌ Error generating timetable: {e}")

if __name__ == "__main__":
    try:
        month = int(input("Enter the month (1-12): "))
        year = int(input("Enter the year (e.g., 2025): "))
        if not (1 <= month <= 12 and 1900 <= year <= 2100):
            raise ValueError("Invalid month or year.")
        generate_timetable(month, year)
    except ValueError as e:
        print(f"Invalid input: {e}")
