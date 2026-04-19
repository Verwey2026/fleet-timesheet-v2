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
st.set_page_config(page_title="Fleet Timesheet Processor V2.6", layout="wide")
st.title("Fleet Timesheet Processor VERSION 2.6 - Verwey Vervoer Format")

st.markdown("Allocates drivers to Fleet Numbers and calculates **Normal Hours, Overtime @1.5, Yard Hours**.")

# Settings
NORMAL_HOURS_PER_DAY = 9.0
st.sidebar.subheader("Settings")
NORMAL_HOURS_PER_DAY = st.sidebar.number_input("Normal Hours Per Day", value=9.0, step=0.5)

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload Tracking Report")
    tracking_file = st.file_uploader("Must have: Fleet Number, Start Time, End Time", type=["xlsx", "xls", "csv"], key="tracking")

with col2:
    st.subheader("2. Upload Driver Allocation")
    allocation_file = st.file_uploader("Verwey format: Sheet = Driver, Has DAY|DATE|FLEET columns", type=["xlsx", "xls"], key="allocation")

def find_header_row(df_raw):
    # Look for the row containing 'DAY' and 'DATE' and 'FLEET'
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(x).upper() for x in row.values])
        if 'DAY' in row_str and 'DATE' in row_str and 'FLEET' in row_str:
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
        'driver': 'Employee Name', 'employee': 'Employee Name', 'employee name': 'Employee Name', 'name': 'Employee Name', 'phh': 'Employee Name',
        'notes': 'Activity Description', 'description': 'Activity Description', 'activity': 'Activity Description', 'activity description': 'Activity Description', 'unnamed 11': 'Activity Description',
        'start': 'Start Time', 'start time': 'Start Time', 'departure time': 'Start Time', 'first movement': 'Start Time',
        'end': 'End Time', 'end time': 'End Time', 'arrival time': 'End Time', 'last movement': 'End Time',
        'fleet': 'Fleet Number', 'vehicle': 'Fleet Number', 'truck': 'Fleet Number', 'fleet no': 'Fleet Number', 
        'reg': 'Fleet Number', 'registration': 'Fleet Number', 'registration nr': 'Fleet Number', 'reg nr': 'Fleet Number', 'fleet number': 'Fleet Number'
    }
    df.columns = [clean_col_name(col) for col in df.columns]
    for old, new in rename_map.items():
        df.columns = [new if old == col else col for col in df.columns]
    
    df = df.loc[:, ~df.columns.duplicated()]
    return df

if tracking_file and allocation_file:
    try:
        # Read tracking file
        df_track_raw = pd.read_excel(tracking_file, header=None)
        track_header = find_header_row(df_track_raw)
        df_track = pd.read_excel(tracking_file, header=track_header)
        df_track = standardize_columns(df_track)

        # Read allocation file - each sheet = one driver
        xls = pd.ExcelFile(allocation_file)
        all_alloc_dfs = []
        
        for sheet_name in xls.sheet_names:
            df_sheet_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            header_row = find_header_row(df_sheet_raw)
            
            if header_row == 0 and 'DAY' not in ' '.join([str(x) for x in df_sheet_raw.iloc[0].values]):
                st.warning(f"Skipping sheet '{sheet_name}' - couldn't find DAY|DATE|FLEET header")
                continue
                
            df_sheet = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)
            df_sheet = standardize_columns(df_sheet)
            
            # Driver name = sheet name
            df_sheet['Employee Name'] = sheet_name
            
            all_alloc_dfs.append(df_sheet)
        
        if not all_alloc_dfs:
            st.error("No valid sheets found with DAY|DATE|FLEET headers")
            st.stop()
            
        df_alloc = pd.concat(all_alloc_dfs, ignore_index=True)

        # Create Date from Start Time if tracking file has no Date column
        if 'Date' not in df_track.columns:
            if 'Start Time' in df_track.columns:
                df_track['Date'] = pd.to_datetime(df_track['Start Time'], errors='coerce')
            else:
                st.error("Tracking file missing both 'Date' and 'Start Time' columns")
                st.stop()
        
        # Parse allocation dates with dayfirst=True for "21 02 2026" format
        df_track['Date'] = pd.to_datetime(df_track['Date'], errors='coerce', dayfirst=True).dt.strftime('%Y-%m-%d')
        df_alloc['Date'] = pd.to_datetime(df_alloc['Date'], errors='coerce', dayfirst=True).dt.strftime('%Y-%m-%d')
        
        # Force Fleet Number to string, strip whitespace, uppercase for matching
        df_track['Fleet Number'] = df_track['Fleet Number'].astype(str).str.strip().str.upper()
        df_alloc['Fleet Number'] = df_alloc['Fleet Number'].astype(str).str.strip().str.upper()

        # Remove empty rows
        df_track = df_track.dropna(subset=['Fleet Number', 'Date'])
        df_alloc = df_alloc.dropna(subset=['Fleet Number', 'Date'])

        # Check required columns
        required_track = ['Fleet Number', 'Date', 'Start Time', 'End Time']
        required_alloc = ['Fleet Number', 'Date', 'Employee Name']
        
        missing_track = [col for col in required_track if col not in df_track.columns]
        missing_alloc = [col for col in required_alloc if col not in df_alloc.columns]
        
        if missing_track:
            st.error(f"Tracking file missing: {missing_track}. Found: {list(df_track.columns)}")
            st.stop()
        if missing_alloc:
            st.error(f"Allocation file missing: {missing_alloc}. Found: {list(df_alloc.columns)}")
            st.stop()

        # Show what we're trying to merge
        with st.expander("Debug: Data being merged"):
            st.write("**Tracking:**")
            st.dataframe(df_track[['Fleet Number', 'Date']].drop_duplicates().head(10))
            st.write("**Allocation:**")
            st.dataframe(df_alloc[['Fleet Number', 'Date', 'Employee Name']].drop_duplicates().head(10))

        # Merge = allocate driver to fleet number
        df_merged = pd.merge(df_track, df_alloc, on=['Fleet Number', 'Date'], how='inner')

        if df_merged.empty:
            st.error("No matching rows found. Check that Fleet Number and Date exist in BOTH files and match exactly.")
            st.stop()

        # Extract yard hours if Activity Description exists
        if 'Activity Description' in df_merged.columns:
            df_merged['yard_hours'] = df_merged['Activity Description'].apply(extract_yard_hours)
        else:
            df_merged['yard_hours'] = 0.0
        
        # Calculate total hours from Start/End Time
        df_merged['Start Time'] = pd.to_datetime(df_merged['Start Time'], errors='coerce')
        df_merged['End Time'] = pd.to_datetime(df_merged['End Time'], errors='coerce')
        df_merged['total_hours'] = (df_merged['End Time'] - df_merged['Start Time']).dt.total_seconds() / 3600
        
        # Driving hours = total - yard
        df_merged['driving_hours'] = (df_merged['total_hours'] - df_merged['yard_hours']).clip(lower=0)
        
        # Normal vs Overtime calculation
        df_merged['normal_hours'] = df_merged['total_hours'].clip(upper=NORMAL_HOURS_PER_DAY)
        df_merged['overtime_hours'] = (df_merged['total_hours'] - NORMAL_HOURS_PER_DAY).clip(lower=0)

        display_cols = ['Date', 'Employee Name', 'Fleet Number', 'Start Time', 'End Time', 
                       'total_hours', 'normal_hours', 'overtime_hours', 'yard_hours', 'driving_hours']
        display_cols = [col for col in display_cols if col in df_merged.columns]
        
        st.success(f"Allocated {len(df_merged)} trips to drivers successfully!")
        
        # Summary matching your NBCRFLI format
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
        if 'df_track' in locals(): st.write("Tracking columns:", list(df_track.columns))
        if 'df_alloc' in locals(): st.write("Allocation columns:", list(df_alloc.columns))
else:
    st.info("👆 Upload both files to start processing")
