# mva.py
import pandas as pd
import re

# Utility functions
def to_clean_str(x) -> str:
    """Convert to clean string, handling various data types"""
    if pd.isna(x):
        return ""
    return str(x).strip()

def find_label_value(row: list, keywords: list, lookahead=6):
    """Search row cells for any keyword (case-insensitive substring)"""
    row = [to_clean_str(c) for c in row]
    row_lower = [c.lower() for c in row]
    for j, cell in enumerate(row_lower):
        for kw in keywords:
            if kw.lower() in cell:
                # find next non-empty after j
                for k in range(j+1, min(j+1+lookahead, len(row))):
                    val = row[k]
                    if val != "":
                        return val
                return ""
    return ""

def is_int_str(s: str) -> bool:
    """Check if string represents an integer"""
    try:
        int(float(s))
        return True
    except:
        return False

# MVA-specific parser
def parse_mva_blocks(df: pd.DataFrame) -> list:
    """Parse Macro Vision Academy format"""
    blocks = []
    nrows, ncols = df.shape
    df = df.fillna("").astype(str)
    
    i = 0
    while i < nrows:
        row = df.iloc[i].tolist()
        row_join = " ".join([to_clean_str(x) for x in row])
        
        if "empcode" in row_join.lower():
            header_row = row
            # Extract employee information
            empcode = find_label_value(header_row, ["empcode"])
            name = find_label_value(header_row, ["name"])
            
            # Fetch department as 'DeptName'
            dept = find_label_value(header_row, ["deptname", "dept name", "department"])
            
            # Fetch month as 'ReportMonth'
            month = find_label_value(header_row, ["reportmonth", "report month"])
            
            # If month not found in header row, try the next row
            if not month and i+1 < nrows:
                srow = df.iloc[i+1].tolist()
                month = find_label_value(srow, ["reportmonth", "report month"])
            
            # Find summary row (usually i+1)
            srow = df.iloc[i+1].tolist() if i+1 < nrows else [""]*ncols
            present = to_clean_str(find_label_value(srow, ["present"]))
            wo = to_clean_str(find_label_value(srow, ["wo", "weekly off"]))
            hl = to_clean_str(find_label_value(srow, ["hl", "half", "half day"]))
            lv = to_clean_str(find_label_value(srow, ["lv", "leave"]))
            absent = to_clean_str(find_label_value(srow, ["absent"]))
            
            # Convert numerical values safely
            def safe_float_str(x):
                try:
                    return float(x)
                except:
                    return 0.0
                    
            present_v = safe_float_str(present)
            wo_v = safe_float_str(wo)
            hl_v = safe_float_str(hl)
            lv_v = safe_float_str(lv)
            absent_v = safe_float_str(absent)
            
            # Find day row (1..31) by scanning ahead
            day_row_idx = None
            for j in range(i+1, min(i+15, nrows)):
                rowj = df.iloc[j].tolist()
                ints = [int(float(x)) for x in rowj if is_int_str(x)]
                if ints and min(ints) == 1 and max(ints) <= 31 and len(ints) >= 10:
                    day_row_idx = j
                    break
                    
            if day_row_idx is None:
                i += 1
                continue
            
            weekdays_row_idx = day_row_idx + 1 if day_row_idx + 1 < nrows else None
            
            # Find the label rows for IN/OUT/WORK/Status below weekdays row
            label_to_idx = {}
            start_scan = (weekdays_row_idx + 1) if weekdays_row_idx else day_row_idx + 1
            for j in range(start_scan, min(start_scan + 12, nrows)):
                first_cell = to_clean_str(df.iloc[j, 0])
                up = first_cell.upper()
                if up.startswith("IN"):
                    label_to_idx["IN"] = j
                elif up.startswith("OUT"):
                    label_to_idx["OUT"] = j
                elif up.startswith("WORK"):
                    label_to_idx["WORK"] = j
                elif "STATUS" in up:
                    label_to_idx["Status"] = j
            
            # Build list of day columns (col index -> day number)
            day_cols = []
            for col in range(ncols):
                val = to_clean_str(df.iat[day_row_idx, col])
                if is_int_str(val):
                    try:
                        dnum = int(float(val))
                        if 1 <= dnum <= 31:
                            day_cols.append((col, dnum))
                    except:
                        pass
            
            daily_records = []
            for col_idx, dnum in day_cols:
                weekday = to_clean_str(df.iat[weekdays_row_idx, col_idx]) if weekdays_row_idx else ""
                
                def cell_at(ridx, cidx):
                    if ridx is None or ridx < 0 or ridx >= nrows or cidx < 0 or cidx >= ncols:
                        return ""
                    return to_clean_str(df.iat[ridx, cidx])
                    
                in_time = cell_at(label_to_idx.get("IN", None), col_idx)
                out_time = cell_at(label_to_idx.get("OUT", None), col_idx)
                work_time = cell_at(label_to_idx.get("WORK", None), col_idx)
                status = cell_at(label_to_idx.get("Status", None), col_idx)
                
                daily_records.append({
                    "day": dnum,
                    "weekday": weekday,
                    "in_time": in_time,
                    "out_time": out_time,
                    "work_time": work_time,
                    "status": status
                })
            
            blocks.append({
                "empcode": empcode,
                "name": name,
                "dept": dept,
                "month": month,
                "present": present_v,
                "wo": wo_v,
                "hl": hl_v,
                "lv": lv_v,
                "absent": absent_v,
                "daily": daily_records,
                "header_row_index": i,
                "day_row_index": day_row_idx
            })
            
            # Move to next employee block
            i = (label_to_idx.get("Status", day_row_idx) or day_row_idx) + 2
            continue
        i += 1
    
    return blocks

def calculate_work_hours(in_time, out_time):
    """Calculate work hours from in and out times"""
    if pd.isna(in_time) or pd.isna(out_time):
        return 0.0
    
    # Convert to datetime if they're strings
    if isinstance(in_time, str):
        try:
            in_time = pd.to_datetime(in_time)
        except:
            return 0.0
    if isinstance(out_time, str):
        try:
            out_time = pd.to_datetime(out_time)
        except:
            return 0.0
    
    # Calculate difference in hours
    if hasattr(in_time, 'timestamp') and hasattr(out_time, 'timestamp'):
        hours = (out_time - in_time).total_seconds() / 3600
        return max(0, min(12, hours))  # Cap at 12 hours, minimum 0
    return 0.0

def calculate_fm_status(present_days, half_days, leaves, absent_days, week_offs, total_working_days):
    """
    Calculate FM status based on attendance rules:
    1. Week offs (WO) are paid and don't affect FM status (max 4 days)
    2. 1 sick leave is allowed without penalty (counts as FM)
    3. FM+1: Present all working days with no leaves/absences (including 27+4 case)
    4. FM: Present all working days with 1 sick leave
    5. FM-N: For N days absent beyond the 1 allowed sick leave
    """
    # Special case: If present for all working days + week offs (27+4=31) with no absences/leaves/half-days
    if (present_days >= 27 and week_offs == 4 and half_days == 0 and leaves == 0 and absent_days == 0) or \
       (present_days == 31 and half_days == 0 and leaves == 0 and absent_days == 0):
        return "FM+1"
        
    # Calculate total working days (excluding week offs)
    working_days = total_working_days - week_offs
    
    # Calculate total attendance (present + half days + leaves)
    total_attendance = present_days + (half_days * 0.5) + leaves
    
    # Case 1: Perfect attendance - all working days present, no leaves/absences
    if present_days == working_days and half_days == 0 and leaves == 0 and absent_days == 0:
        return "FM+1"
    
    # Case 2: 1 day absent (covered by sick leave) - counts as FM
    if absent_days == 1 and leaves == 0:
        return "FM"
        
    # Case 3: 1 sick leave taken (counts as FM)
    if absent_days == 0 and leaves == 1:
        return "FM"
    
    # Case 4: Multiple absences beyond sick leave
    if absent_days > 1:
        return f"FM-{absent_days - 1}"
    
    # Case 5: Multiple leaves beyond sick leave
    if leaves > 1:
        return f"FM-{leaves - 1}"
    
    # Half days are already counted as 0.5 present + 0.5 absent
    # No need for special handling here as absences are already counted
    
    # Default to FM if none of the above conditions are met
    return "FM"

def process_attendance_data(blocks: list) -> pd.DataFrame:
    """Process parsed blocks and calculate attendance metrics"""
    results = []
    
    for block in blocks:
        # Extract basic information
        emp_id = block.get('empcode', '')
        emp_name = block.get('name', '')
        department = block.get('dept', '')
        month = block.get('month', '')
        
        # Calculate total working days in the month (excluding weekends)
        if month:
            try:
                if '-' in month:
                    month_part, year_part = month.split('-')
                    month_days = {
                        'jan': 31, 'feb': 29 if int(year_part) % 4 == 0 else 28, 
                        'mar': 31, 'apr': 30, 'may': 31, 'jun': 30,
                        'jul': 31, 'aug': 31, 'sep': 30, 'oct': 31, 'nov': 30, 'dec': 31
                    }
                    total_days = month_days.get(month_part.lower()[:3], 30)
                    # Calculate working days (excluding weekends)
                    total_working_days = sum(1 for day in range(1, total_days + 1) 
                                          if pd.Timestamp(f'{year_part}-{month_part}-{day}').dayofweek < 5)
            except:
                # Default to 26 working days if month parsing fails
                total_working_days = 26
        else:
            # Default to 26 working days if no month info
            total_working_days = 26
        
        # Initialize counters
        present_days = 0
        half_days = 0
        leaves = 0
        week_offs = 0  # Will be counted from the sheet
        absent_days = 0
        total_ot = 0.0  # Will store total OT hours
        total_ot_from_sheet = 0.0  # Will store TotalOT from the sheet
        daily_records = []
        
        # Process daily records if available
        for record in block.get('daily', []):
            status = str(record.get('status', '')).strip().upper()
            in_time = record.get('in_time')
            out_time = record.get('out_time')
            
            # Update counters based on status
            if 'P' in status:  # Present
                present_days += 1
            elif 'H' in status:  # Half day
                half_days += 1
                present_days += 0.5  # Count as 0.5 present
                absent_days += 0.5   # And 0.5 absent
            elif 'L' in status:  # Leave
                leaves += 1
            elif 'WO' in status:  # Week off from sheet
                week_offs += 1
            elif status == 'A':  # Absent
                absent_days += 1
            # Ignore empty or invalid status
            
            # Check for OT in the status
            if 'OT' in status:
                try:
                    # Extract OT hours (e.g., 'P OT:2' or 'OT:4')
                    ot_part = status.split('OT:')[-1].split()[0]
                    total_ot += float(ot_part)
                except (ValueError, IndexError):
                    pass  # If OT format is not as expected, skip it
                    
            # Check for TotalOT in the status (from the Excel sheet)
            if 'TOTALOT' in status.upper():
                try:
                    # Extract the numeric value after 'TOTALOT:'
                    ot_value = status.upper().split('TOTALOT:')[-1].strip()
                    if ot_value.replace('.', '').isdigit():  # Check if it's a valid number
                        total_ot_from_sheet = float(ot_value)
                except (ValueError, IndexError, AttributeError):
                    pass
            
            # Store daily record (without work hours)
            daily_records.append({
                'date': record.get('date', ''),
                'status': status,
                'in_time': in_time if pd.notna(in_time) else '',
                'out_time': out_time if pd.notna(out_time) else ''
            })
        
        # Calculate FM status
        fm_status = calculate_fm_status(
            present_days, 
            half_days, 
            leaves, 
            absent_days, 
            week_offs,
            total_working_days
        )
        
        # Add to results (excluding 'Total Working Days' and 'Daily Records')
        results.append({
            'Employee ID': emp_id,
            'Employee Name': emp_name,
            'Department': department,
            'Month': month,
            'Present Days': present_days,
            'Half Days': half_days,
            'Leaves': leaves,
            'Week Offs': week_offs,
            'Absent Days': absent_days,
            'Total OT Hours': round(total_ot_from_sheet if total_ot_from_sheet > 0 else total_ot, 2),
            'FM Status': fm_status
        })
    
    return pd.DataFrame(results)