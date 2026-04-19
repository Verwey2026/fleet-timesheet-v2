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
st.set_page_config(page_title="Fleet Timesheet Processor V4.0", layout="wide")
st.title("Fleet Timesheet Processor VERSION 4.0 - Verwey Vervoer")

st.markdown("Allocates drivers to Fleet Numbers and calculates **Normal Hours, Overtime @1.5, Yard Hours**.")

NORMAL_HOURS_PER_DAY = 9.0
st.sidebar.subheader("Settings")
NORMAL_HOURS_PER_DAY = st.sidebar.number_input("Normal Hours Per Day", value=9.0, step=0.5)

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload Tracking Report")
    tracking_file = st.file_uploader("First and Last Movement report", type=["xlsx", "xls", "csv"], key="tracking")

with col2:
    st.subheader("2. Upload Driver Allocation")
    allocation_file = st.file_uploader("Sheet = Driver, Headers: DAY | DATE | FLEET", type=["xlsx", "xls"], key="allocation")

def find_tracking_header(df_raw):
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(x).upper() for x in row.values])
        if 'REGISTRATION' in row_str and 'DEPARTURE' in row_str and 'ARRIVAL' in row_str:
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
        'date': 'Date', 'trip date': 'Date', 'day': 'Date', 'tripdate': 'Date',
        'driver': 'Employee Name', 'employee': 'Employee Name', 'employee name': 'Employee Name', 'name': 'Employee Name', 'phh': 'Employee Name',
        'notes': 'Activity Description', 'description': 'Activity Description', 'activity': 'Activity Description', 'activity description': 'Activity Description', 'unnamed 11': 'Activity Description',
        'start': 'Start Time', 'start time': 'Start Time', 'departure time': 'Start Time', 'first movement': 'Start Time', 'departure': 'Start Time', 'starttime': 'Start Time',
        'end': 'End Time', 'end time': 'End Time', 'arrival time': 'End Time', 'last movement': 'End Time', 'arrival': 'End Time', 'endtime': 'End Time',
        'fleet': 'Fleet Number', 'vehicle': 'Fleet Number', 'truck': 'Fleet Number', 'fleet no': 'Fleet Number', 
        'reg': 'Fleet Number', 'registration': 'Fleet Number', 'registration nr': 'Fleet Number', 'reg nr': 'Fleet Number', 'fleet number': 'Fleet Number', 'reg no': 'Fleet Number'
    }
    df.columns = [clean_col_name(col) for col in df.columns]
    for old, new in rename_map.items():
        df.columns = [new if old == col else col for col in df.columns]
    
    df = df.loc[:, ~df.columns.duplicated()]
    return df

if tracking_file and allocation_file:
    try:
        # Read tracking file - auto detect header
        df_track_raw = pd.read_excel(tracking_file, header=None)
        track_header = find_tracking_header(df_track_raw)
        df_track = pd.read_excel(tracking_file, header=track_header)
        df_track = standardize_columns(df_track)

        # Read allocation file - clean format, header row 0
        xls = pd.ExcelFile(allocation_file)
        all_alloc_dfs = []
        
        for sheet_name in xls.sheet_names:
            df_sheet = pd.read_excel(xls, sheet_name=sheet_name)  # Header = row 1
            df_sheet = standardize_columns(df_sheet)
            df_sheet['Employee Name'] = sheet_name
            all_alloc_dfs.append(df_sheet)
        
        if not all_alloc_dfs:
            st.error("No valid sheets found in allocation file")
            st.stop()
            
        df_alloc = pd.concat(all_alloc_dfs, ignore_index=True)

        # Create Date from Start Time if tracking file has no Date column
        if 'Date' not in df_track.columns:
            if 'Start Time' in df_track.columns:
                df_track['Date'] = pd.to_datetime(df_track['Start Time'], errors='coerce')
            else:
                st.error(f"Tracking file missing both 'Date' and 'Start Time' columns. Found: {list(df_track.columns)}")
                st.stop()
        
        # Parse dates - handles '2026 02 21' format
        df_track['Date'] = pd.to_datetime(df_track['Date'], errors='coerce', dayfirst=True).dt.strftime('%Y-%m-%d')
        df_alloc['Date'] = pd.to_datetime(df_alloc['Date'], errors='coerce', dayfirst=True).dt.strftime('%Y-%m-%d')
        
        # Standardize Fleet Number
        df_track['Fleet Number'] = df_track['Fleet Number'].astype(str).str.strip().str.upper()
        df_alloc['Fleet Number'] = df_alloc['Fleet Number'].astype(str).str.strip().str.upper()

        st.write(f"Tracking rows: {len(df_track)}")
        st.write(f"Allocation rows: {len(df_alloc)}")
        
        df_track = df_track.dropna(subset=['Fleet Number', 'Date', 'Start Time', 'End Time'])
        df_alloc = df_alloc.dropna(subset=['Fleet Number', 'Date'])
        
        st.write(f"After cleaning - Tracking: {len(df_track)}, Allocation: {len(df_alloc)}")

        if df_alloc.empty:
            st.error("Allocation data is empty after cleaning. Make sure FLEET column has values for worked days.")
            st.stop()

        df_merged = pd.merge(df_track, df_alloc, on=['Fleet Number', 'Date'], how='inner')

        if df_merged.empty:
            st.error("No matching rows.")
            with st.expander("Debug: Keys"):
                st.write("**Tracking keys:**")
                st.dataframe(df_track[['Fleet Number', 'Date']].drop_duplicates().head(10))
                st.write("**Allocation keys:**")
                st.dataframe(df_alloc[['Fleet Number', 'Date']].drop_duplicates().head(10))
            st.stop()

        if 'Activity Description' in df_merged.columns:
            df_merged['yard_hours'] = df_merged['Activity Description'].apply(extract_yard_hours)
        else:
            df_merged['yard_hours'] = 0.0
        
        df_merged['Start Time'] = pd.to_datetime(df_merged['Start Time'], errors='coerce')
        df_merged['End Time'] = pd.to_datetime(df_merged['End Time'], errors='coerce')
        df_merged['total_hours'] = (df_merged['End Time'] - df_merged['Start Time']).dt.total_seconds() / 3600
        df_merged['driving_hours'] = (df_merged['total_hours'] - df_merged['yard_hours']).clip(lower=0)
        df_merged['normal_hours'] = df_merged['total_hours'].clip(upper=NORMAL_HOURS_PER_DAY)
        df_merged['overtime_hours'] = (df_merged['total_hours'] - NORMAL_HOURS_PER_DAY).clip(lower=0)

        display_cols = ['Date', 'Employee Name', 'Fleet Number', 'Start Time', 'End Time', 
                       'total_hours', 'normal_hours', 'overtime_hours', 'yard_hours', 'driving_hours']
        display_cols = [col for col in display_cols if col in df_merged.columns]
        
        st.success(f"Allocated {len(df_merged)} trips to drivers!")
        
        summary = df_merged.groupby('Employee Name')[['total_hours', 'normal_hours', 'overtime_hours', 'yard_hours']].sum().round(2)
        summary.columns = ['TOTAL HOURS', 'NORMAL HOURS', 'OVERTIME @1.5', 'YARD']
        
        st.subheader("Summary by Driver - NBCRFLI Format")
        st.dataframe(summary)
        
        st.subheader("Detailed Trips")
        st.dataframe(df_merged[display_cols].round(2))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_merged[display_cols].to_excel(writer, index=False, sheet_name='Detailed')
            summary.to_excel(writer, sheet_name='Summary')
        output.seek(0)
        
        st.download_button("📥 Download Excel with Summary", output, "fleet_timesheet_processed.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error: {str(e)}")
else:
    st.info("👆 Upload both files to start processing")
