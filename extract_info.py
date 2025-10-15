import pandas as pd
import mva

def extract_dept_and_month(file_path):
    """Extract department name and month from the Excel file."""
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, header=None)
        
        # Find the first non-empty row that might contain the headers
        for i in range(min(10, len(df))):  # Check first 10 rows
            row = df.iloc[i].fillna('').astype(str).tolist()
            
            # Look for department info
            dept = mva.find_label_value(row, ["deptname", "dept name", "department"])
            month = mva.find_label_value(row, ["reportmonth", "report month"])
            
            # If we found both, return them
            if dept and month:
                return {"Department": dept, "Month": month}
            
            # If we found one but not the other, continue searching
            if dept or month:
                # Check next row for the missing value
                if i + 1 < len(df):
                    next_row = df.iloc[i+1].fillna('').astype(str).tolist()
                    if not dept:
                        dept = mva.find_label_value(next_row, ["deptname", "dept name", "department"])
                    if not month:
                        month = mva.find_label_value(next_row, ["reportmonth", "report month"])
                    
                    if dept and month:
                        return {"Department": dept, "Month": month}
        
        # If we get here, we didn't find both values
        return {"error": "Could not find both department and month in the file"}
        
    except Exception as e:
        return {"error": f"Error processing file: {str(e)}"}

if __name__ == "__main__":
    file_path = r"uploads/MVA.xls"
    result = extract_dept_and_month(file_path)
    print("Extracted Information:")
    for key, value in result.items():
        print(f"{key}: {value}")
