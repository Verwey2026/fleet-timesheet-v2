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

st.markdown("Upload your **Tracking Report** and **Driver Allocation** files. The app will merge them, extract yard hours, and pull fleet numbers.")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload Tracking Report")
    tracking_file = st.file_uploader(
        "Tracking file with Start/End times", 
        type=["xlsx", "xls", "csv"], 
        key="tracking"
    )

with col2:
    st.subheader("2. Upload Driver Allocation")
    allocation_file = st.file_uploader(
        "Allocation file with Employee Names + Descriptions", 
        type=["xlsx", "xls", "csv"], 
        key="allocation"
    )

def extract_yard_hours(text):
    if pd.isna(text):
        return 0.0
    text = str(text).lower()
    # Matches: "2h yard", "yard: 1.5", "1.5hr yard", "yard work 2"
    match = re.search(r'(\d+\.?\d*)\s*h(?:our)?s?\s*yard|yard[:\s]*(\d+\.?\d*)', text)
    if match:
        return float(match.group(1) or match.group(2))
    return 0.0

def extract_fleet_number(text):
    if pd.isna(text):
        return ""
    text = str(text).upper()
    # Matches: TDH123, Fleet #456, FLEET: ABC789
    match = re.search(r'(?:FLEET|TDH|ABC|TRUCK)[\s#:]*([A-Z0-9]+)', text)
    if match:
        return match.group(1)
    return ""

def standardize_columns(df):
    # Map common column variations to standard names
    rename_map = {
        'date': 'Date', 'trip date': 'Date',
        'driver': 'Employee Name', 'employee': 'Employee Name', 'employee name': 'Employee Name', 'name': 'Employee Name',
        'notes': 'Activity Description', 'description': 'Activity Description', 'activity': 'Activity Description', 'comments': 'Activity Description',
        'start': 'Start Time', 'start time': 'Start Time', 'clock in': 'Start Time', 'time start': 'Start Time',
        'end': 'End Time', 'end time': 'End Time', 'clock out': 'End Time', 'time end': 'End Time',
        'fleet': 'Fleet Number', 'vehicle': 'Fleet Number', 'truck': 'Fleet Number', 'fleet no': 'Fleet Number'
    }
    df.columns = [col.strip() for col in df.columns]
    df = df.rename(columns={k.title(): v for k, v in rename_map.items() if k.title() in df.columns})
    df = df.rename(columns={k.lower(): v for k, v in rename_map.items() if k.lower() in df.columns})
    df = df.rename(columns={k.upper(): v for k, v in rename_map.items() if k.upper() in df.columns})
    return df

if tracking_file and allocation_file:
    try:
        # Read files
        if tracking_file.name.endswith('.csv'):
            df_track = pd.read_csv(tracking_file)
        else:
            df_track = pd.read_excel(tracking_file)
            
        if allocation_file.name.endswith('.csv'):
            df_alloc = pd.read_csv(allocation_file)
        else:
            df_alloc = pd.read_excel(allocation_file)

        # Standardize column names
        df_track = standardize_columns(df_track)
        df_alloc = standardize_columns(df_alloc)

        # Check required columns
        required_track = ['Date', 'Employee Name', 'Start Time', 'End Time']
        required_alloc = ['Date', 'Employee Name', 'Activity Description']
        
        missing_track = [col for col in required_track if col not in df_track.columns]
        missing_alloc = [col for col in required_alloc if col not in df_alloc.columns]
        
        if missing_track:
            st.error(f"Tracking file missing columns: {missing_track}. Found: {list(df_track.columns)}")
            st.stop()
        if missing_alloc:
            st.error(f"Allocation file missing columns: {missing_alloc}. Found: {list(df_alloc.columns)}")
            st.stop()

        # Merge on Date + Employee Name
        df_merged = pd.merge(
            df_track, 
            df_alloc, 
            on=['Date', 'Employee Name'], 
            how='inner',
            suffixes=('_track', '_alloc')
        )

        if df_merged.empty:
            st.error("No matching rows found between files. Check that Date and Employee Name match exactly in both files.")
            st.stop()

        # Extract yard hours and fleet numbers from Activity Description
        df_merged['yard_hours'] = df_merged['Activity Description'].apply(extract_yard_hours)
        df_merged['fleet_number_extracted'] = df_merged['Activity Description'].apply(extract_fleet_number)
        
        # Use fleet number from tracking if available, else use extracted
        if 'Fleet Number' in df_merged.columns:
            df_merged['fleet_number'] = df_merged['Fleet Number'].fillna(df_merged['fleet_number_extracted'])
        else:
            df_merged['fleet_number'] = df_merged['fleet_number_extracted']

        # Calculate total hours
        df_merged['Start Time'] = pd.to_datetime(df_merged['Start Time'], errors='coerce')
        df_merged['End Time'] = pd.to_datetime(df_merged['End Time'], errors='coerce')
        df_merged['total_hours'] = (df_merged['End Time'] - df_merged['Start Time']).dt.total_seconds() / 3600
        df_merged['driving_hours'] = df_merged['total_hours'] - df_merged['yard_hours']

        # Clean up columns for display
        display_cols = ['Date', 'Employee Name', 'fleet_number', 'Start Time', 'End Time', 
                       'total_hours', 'yard_hours', 'driving_hours', 'Activity Description']
        display_cols = [col for col in display_cols if col in df_merged.columns]
        
        st.success(f"Merged {len(df_merged)} rows successfully!")
        st.dataframe(df_merged[display_cols])

        # Download button
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_merged[display_cols].to_excel(writer, index=False, sheet_name='Processed')
        output.seek(0)
        
        st.download_button(
            label="📥 Download Processed Excel",
            data=output,
            file_name="fleet_timesheet_processed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error processing files: {str(e)}")
        st.write("**Debug info:**")
        if 'df_track' in locals(): st.write("Tracking columns:", list(df_track.columns))
        if 'df_alloc' in locals(): st.write("Allocation columns:", list(df_alloc.columns))

else:
    st.info("👆 Upload both files to start processing")
