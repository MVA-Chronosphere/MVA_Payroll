import pandas as pd

def inspect_excel(file_path):
    # Read the Excel file
    try:
        # Try reading with header=None to see raw data
        df = pd.read_excel(file_path, header=None)
        print("First 10 rows of the Excel file:")
        print(df.head(10).to_string())
        
        # Try to find department and month by searching all cells
        print("\nSearching for department and month in the first 20 rows...")
        for i in range(min(20, len(df))):
            for j in range(len(df.columns)):
                cell_value = str(df.iat[i, j]).strip().lower()
                if any(term in cell_value for term in ["dept", "department"]):
                    print(f"\nPossible department header at row {i+1}, column {j+1}: {df.iat[i, j]}")
                    # Look at the next cell for the value
                    if j + 1 < len(df.columns):
                        print(f"Possible department value: {df.iat[i, j+1]}")
                
                if any(term in cell_value for term in ["month", "report"]):
                    print(f"\nPossible month header at row {i+1}, column {j+1}: {df.iat[i, j]}")
                    # Look at the next cell for the value
                    if j + 1 < len(df.columns):
                        print(f"Possible month value: {df.iat[i, j+1]}")
                    
    except Exception as e:
        print(f"Error reading Excel file: {e}")

if __name__ == "__main__":
    file_path = r"uploads/MVA.xls"
    inspect_excel(file_path)
