# MVA Attendance — Full Engine (Mongo)

## Setup
1. Edit `mongo_helpers.py` and set `MONGO_URI`.
2. Create and activate a Python virtualenv.
3. `pip install -r requirements.txt`

## Quick start
1. Start MongoDB (or ensure Atlas reachable).
2. Run shift admin to define shifts & assign:
   `streamlit run shift_admin.py`
   - Create shift patterns (Morning, Night, Split)
   - Assign shifts to employees for the month (bulk assign or per-day)
3. Run main app:
   `streamlit run app.py`
   - Upload monthly attendance (MVA-style Excel)
   - Click "Process & Save (Full Engine)"
   - Review daily metrics and monthly summary; download Excel

## Notes
- Shift importer is a starting point. If your shift sheet format is complex I will tailor `shift_importer.py` to the exact structure you provide.
- Split shifts and overnight shifts are supported.
- You can correct a day's punches by editing `attendance_daily` using any Mongo GUI (e.g., MongoDB Compass) and re-run `compute_attendance_metrics.compute_for_month(year_month)` to re-calc.
