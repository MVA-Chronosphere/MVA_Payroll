import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Employee Shift Extractor")

st.write("Upload your Excel file with columns: `Employee`, `Shift Type`, `Start Time`, `End Time`, `Date`")

file = st.file_uploader("Upload Excel File", type=["xlsx"])

def to_excel(df_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet, df in df_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet)
    output.seek(0)
    return output

if file is not None:
    try:
        df = pd.read_excel(file)

        st.subheader("Uploaded Data Preview")
        st.dataframe(df.head())

        required_cols = ["Employee", "Shift Type", "Start Time", "End Time", "Date"]
        
        if not all(col in df.columns for col in required_cols):
            st.error(f"Excel must contain these columns: {required_cols}")
        else:
            # Convert times to datetime
            df["Start Time"] = pd.to_datetime(df["Start Time"], errors="coerce")
            df["End Time"] = pd.to_datetime(df["End Time"], errors="coerce")

            # Identify shift types
            day_shift = df[df["Shift Type"].str.contains("day", case=False, na=False)]
            
            # Night shift: end next day 00:00-01:00 or explicit "night"
            night_shift = df[
                (df["Shift Type"].str.contains("night", case=False, na=False))
                | (df["End Time"].dt.hour.isin([0, 1]))
            ]
            
            off_duty = df[df["Shift Type"].str.contains("off", case=False, na=False)]

            result = {
                "Day Shift": day_shift,
                "Night Shift": night_shift,
                "Off Duty": off_duty
            }

            st.success("Shifts separated successfully")

            st.subheader("Day Shift")
            st.dataframe(day_shift)

            st.subheader("Night Shift")
            st.dataframe(night_shift)

            st.subheader("Off Duty")
            st.dataframe(off_duty)

            excel_file = to_excel(result)
            st.download_button(
                label="Download Result Excel",
                data=excel_file,
                file_name="shift_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error reading file: {e}")
