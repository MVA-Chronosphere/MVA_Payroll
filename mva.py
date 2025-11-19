# mva.py
import pandas as pd
import re

def to_clean_str(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()

def find_label_value(row: list, keywords: list, lookahead=6):
    row = [to_clean_str(c) for c in row]
    row_lower = [c.lower() for c in row]
    for j, cell in enumerate(row_lower):
        for kw in keywords:
            if kw.lower() in cell:
                for k in range(j+1, min(j+1+lookahead, len(row))):
                    val = row[k]
                    if val != "":
                        return val
                return ""
    return ""

def is_int_str(s: str) -> bool:
    try:
        int(float(s))
        return True
    except:
        return False

def parse_mva_blocks(df: pd.DataFrame) -> list:
    blocks = []
    nrows, ncols = df.shape
    df = df.fillna("").astype(str)
    i = 0
    while i < nrows:
        row = df.iloc[i].tolist()
        row_join = " ".join([to_clean_str(x) for x in row])
        if "empcode" in row_join.lower():
            header_row = row
            empcode = find_label_value(header_row, ["empcode"])
            name = find_label_value(header_row, ["name"])
            dept = find_label_value(header_row, ["deptname", "dept name", "department"])
            
            # Try multiple approaches to find the month
            month = find_label_value(header_row, ["reportmonth", "report month"])
            if not month and i+1 < nrows:
                srow = df.iloc[i+1].tolist()
                month = find_label_value(srow, ["reportmonth", "report month"])
            
            # If month still not found, try looking for other common month patterns in nearby rows
            if not month:
                # Look in a few rows above and below for month information
                for check_row_idx in [i-2, i-1, i+1, i+2]:
                    if 0 <= check_row_idx < nrows:
                        check_row = df.iloc[check_row_idx].tolist()
                        check_row_str = " ".join([to_clean_str(x) for x in check_row])
                        
                        # Look for month names in the row
                        import re
                        month_pattern = r'(january|february|march|april|may|june|july|august|september|october|november|december)'
                        month_match = re.search(month_pattern, check_row_str.lower())
                        if month_match:
                            found_month = month_match.group(1).capitalize()
                            # Try to find the year as well
                            year_pattern = r'\b(20\d{2})\b'
                            year_match = re.search(year_pattern, check_row_str)
                            if year_match:
                                month = f"{found_month}-{year_match.group(1)}"
                            else:
                                month = found_month
                            break
                
                # If still no month, try looking for date patterns like "Month Year" in the entire header area
                if not month:
                    for check_row_idx in range(max(0, i-3), min(nrows, i+5)):
                        check_row = df.iloc[check_row_idx].tolist()
                        for cell in check_row:
                            cell_str = to_clean_str(cell)
                            # Look for patterns like "November 2024", "Nov 2024", etc.
                            date_pattern = r'(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\s+(20\d{2})'
                            date_match = re.search(date_pattern, cell_str.lower())
                            if date_match:
                                month_name = date_match.group(1).capitalize()
                                year = date_match.group(2)
                                month = f"{month_name}-{year}"
                                break
                        if month:
                            break
            
            srow = df.iloc[i+1].tolist() if i+1 < nrows else [""]*ncols
            present = to_clean_str(find_label_value(srow, ["present"]))
            wo = to_clean_str(find_label_value(srow, ["wo", "weekly off"]))
            hl = to_clean_str(find_label_value(srow, ["hl", "half", "half day"]))
            lv = to_clean_str(find_label_value(srow, ["lv", "leave"]))
            absent = to_clean_str(find_label_value(srow, ["absent"]))
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
            i = (label_to_idx.get("Status", day_row_idx) or day_row_idx) + 2
            continue
        i += 1
    return blocks

def calculate_fm_status(present_days, half_days, leaves, absent_days, week_offs, total_working_days):
    working_days = total_working_days - week_offs
    total_attendance = present_days + (half_days * 0.5) + leaves
    if present_days == working_days and half_days == 0 and leaves == 0 and absent_days == 0:
        return "FM"
    if absent_days == 1 and leaves == 0:
        return "FM"
    if absent_days == 0 and leaves == 1:
        return "FM"
    if absent_days > 1:
        return f"FM-{absent_days - 1}"
    if leaves > 1:
        return f"FM-{leaves - 1}"
    return "FM"
