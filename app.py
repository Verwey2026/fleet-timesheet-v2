import streamlit as st
import pandas as pd
import re
import io

# ===== PASSWORD GATE =====
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("Fleet Timesheet Processor")
    password = st.text_input("Enter password", type="password")
    if password:
        if password == st.secrets["app_password"]:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password")
    st.stop()

# ===== MAIN APP =====
st.set_page_config(page_title="Fleet Timesheet Processor V2.4", layout="wide")
st.title("Fleet Timesheet Processor VERSION 2.4 - Yard Hours + Fleet Number Extraction")

st.markdown("Merges on **Fleet Number + Date**. Reads driver name from sheet name. Extracts yard hours from Column K.")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload Tracking Report")
    tracking_file = st.file_uploader("Must have: Fleet Number, Start Time, End Time", type=["xlsx", "xls", "csv"], key="tracking")

with col2:
    st.subheader("2. Upload Driver Allocation")
    allocation_file = st.file_uploader("Sheet name = Driver. Column K = Activity", type=["xlsx", "xls"], key="allocation")

def find_header_row(df_raw):
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(x) for x in row.values]).lower()
        if any(word in row_str for word in ['date', 'fleet', 'start', 'reg', 'registration']):
            return idx
    return 0

def extract_yard_hours(text):
    if pd.isna(text): return 0.0
    text = str(text).lower()
    match = re.search(r'(\d+\.?\d*)\s*h(?:our)?s?\s*yard|yard[:\s]*(\d+\.?\d*)', text)
    if match: return float(match.group(1) or match.group(2))
    return 0.0

def clean_col_name(col):
    return str(col).strip().lower().replace(':', '').replace('.', '').strip()

def standardize_columns(df):
    rename_map = {
        'date': 'Date', 'trip date': 'Date', 'day': 'Date',
        'start': 'Start Time', 'start time': 'Start Time', 'departure time': 'Start Time', 'first movement': 'Start Time',
        'end': 'End Time', 'end time': 'End Time', 'arrival time': 'End Time', 'last movement': 'End Time',
        'fleet': 'Fleet Number', 'vehicle': 'Fleet Number', 'truck': 'Fleet Number', 'fleet no': 'Fleet Number', 
        'reg': 'Fleet Number', 'registration': 'Fleet Number', 'registration nr': 'Fleet Number', 'reg nr': 'Fleet Number', 'fleet number': 'Fleet Number'
    }
    df.columns = [clean_col_name(col) for col in df.columns]
    for old, new in rename_map.items():
        df.columns = [new if old == col else col for col in df.columns]
    return df

if tracking_file and allocation_file:
    try:
        # Read tracking file
        df_track = pd.read_excel(tracking_file, header=find_header_row(pd.read_excel(tracking_file, header=None)))
        df_track = standardize_columns(df_track)

        # Read ALL sheets from allocation file and combine them
        xls = pd.ExcelFile(allocation_file)
        all_alloc_dfs = []
        
        for sheet_name in xls.sheet_names:
            df_sheet_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            header_row = find_header_row(df_sheet_raw)
            df_sheet = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)
            df_sheet = standardize_columns(df_sheet)
            
            # Sheet name = Employee Name
            df_sheet['Employee Name'] = sheet_name
            
            # Column K is index 10 - rename it to Activity Description
            if len(df_sheet.columns) > 10:
                df_sheet = df_sheet.rename(columns={df_sheet.columns[10]: 'Activity Description'})
            
            all_alloc_dfs.append(df_sheet)
        
        df_alloc = pd.concat(all_alloc_dfs, ignore_index=True)

        # Extract Date from Start Time if Date column missing in tracking
        if 'Date' not in df_track.columns and 'Start Time' in df_track.columns:
            df_track['Date'] = pd.to_datetime(df_track['Start Time'], errors='coerce').dt.date
        
        # Convert dates to date type for matching
        if 'Date' in df_track.columns:
            df_track['Date'] = pd.to_datetime(df_track['Date'], errors='coerce').dt.date
        if 'Date' in df_alloc.columns:
            df_alloc['Date'] = pd.to_datetime(df_alloc['Date'], errors='coerce').dt.date

        # Check required columns
        required_track = ['Fleet Number', 'Date', 'Start Time', 'End Time']
        required_alloc = ['Fleet Number', 'Date', 'Employee Name', 'Activity Description']
        
        missing_track = [col for col in required_track if col not in df_track.columns]
        missing_alloc = [col for col in required_alloc if col not in df_alloc.columns]
        
        if missing_track:
            st.error(f"Tracking file missing: {missing_track}. Found: {list(df_track.columns)}")
            st.stop()
        if missing_alloc:
            st.error(f"Allocation file missing: {missing_alloc}. Found: {list(df_alloc.columns)}")
            st.stop()

        # Merge on Fleet Number + Date
        df_merged = pd.merge(df_track, df_alloc, on=['Fleet Number', 'Date'], how='inner')

        if df_merged.empty:
            st.error("No matching rows found between files.")
            st.write("**Tracking sample:**")
            st.dataframe(df_track[['Fleet Number', 'Date']].head())
            st.write("**Allocation sample:**")
            st.dataframe(df_alloc[['Fleet Number', 'Date', 'Employee Name']].head())
            st.stop()

        # Extract yard hours from Column K
        df_merged['yard_hours'] = df_merged['Activity Description'].apply(extract_yard_hours)
        
        # Calculate hours
        df_merged['Start Time'] = pd.to_datetime(df_merged['Start Time'], errors='coerce')
        df_merged['End Time'] = pd.to_datetime(df_merged['End Time'], errors='coerce')
        df_merged['total_hours'] = (df_merged['End Time'] - df_merged['Start Time']).dt.total_seconds() / 3600
        df_merged['driving_hours'] = df_merged['total_hours'] - df_merged['yard_hours']

        display_cols = ['Date', 'Employee Name', 'Fleet Number', 'Start Time', 'End Time', 
                       'total_hours', 'yard_hours', 'driving_hours', 'Activity Description']
        display_cols = [col for col in display_cols if col in df_merged.columns]
        
        st.success(f"Merged {len(df_merged)} rows successfully!")
        st.dataframe(df_merged[display_cols])

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_merged[display_cols].to_excel(writer, index=False, sheet_name='Processed')
        output.seek(0)
        
        st.download_button("📥 Download Processed Excel", output, "fleet_timesheet_processed.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error: {str(e)}")
        if 'df_track' in locals(): st.write("Tracking columns:", list(df_track.columns))
        if 'df_alloc' in locals(): st.write("Allocation columns:", list(df_alloc.columns))
else:
    st.info("👆 Upload both files to start processing")
