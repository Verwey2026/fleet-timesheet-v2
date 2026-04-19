import streamlit as st
import pandas as pd
import re
import io
from datetime import time

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
st.set_page_config(page_title="Fleet Timesheet Processor V4.8", layout="wide")
st.title("Fleet Timesheet Processor VERSION 4.8 - Verwey Vervoer")

st.markdown("**Rules:** 07:00-17:00 shift. 1h unpaid lunch auto-deducted if worked through midday. Overtime @1.5 after 195.03 paid hours.")

NORMAL_HOURS_THRESHOLD = 195.03
st.sidebar.subheader("NBCRFLI Settings")
NORMAL_HOURS_THRESHOLD = st.sidebar.number_input("Normal Hours Threshold", value=195.03, step=0.01)

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload Tracking Report")
    tracking_file = st.file_uploader("First and Last Movement report", type=["xlsx", "xls", "csv"], key="tracking")

with col2:
    st.subheader("2. Upload Driver Allocation")
    allocation_file = st.file_uploader("Sheet = Driver, Headers: DAY | DATE | FLEET | MEAL HOUR | SLEEP OUT", type=["xlsx", "xls"], key="allocation")

def find_tracking_header(df_raw):
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(x).upper() for x in row.values])
        if 'REGISTRATION' in row_str and 'DEPARTURE' in row_str and 'ARRIVAL' in row_str:
            return idx
    return 0

def extract_yard_hours_from_text(text):
    if pd.isna(text): return 0.0
    text = str(text).lower()
    match = re.search(r'(\d+\.?\d*)\s*h(?:our)?s?\s*yard|yard[:\s]*(\d+\.?\d*)', text)
    if match: return float(match.group(1) or match.group(2))
    return 0.0

def clean_col_name(col):
    return str(col).strip().lower().replace(':', '').replace('.', '').strip()

def standardize_columns(df):
    rename_map = {
        'date': 'Date', 'trip date': 'Date', 'tripdate': 'Date',
        'driver': 'Employee Name', 'employee': 'Employee Name', 'employee name': 'Employee Name', 'name': 'Employee Name', 'phh': 'Employee Name',
        'notes': 'Activity Description', 'description': 'Activity Description', 'activity': 'Activity Description', 'activity description': 'Activity Description', 'unnamed 11': 'Activity Description',
        'start': 'Start Time', 'start time': 'Start Time', 'departure time': 'Start Time', 'first movement': 'Start Time', 'departure': 'Start Time', 'starttime': 'Start Time',
        'end': 'End Time', 'end time': 'End Time', 'arrival time': 'End Time', 'last movement': 'End Time', 'arrival': 'End Time', 'endtime': 'End Time',
        'fleet': 'Fleet Number', 'vehicle': 'Fleet Number', 'truck': 'Fleet Number', 'fleet no': 'Fleet Number',
        'reg': 'Fleet Number', 'registration': 'Fleet Number', 'registration nr': 'Fleet Number', 'reg nr': 'Fleet Number', 'fleet number': 'Fleet Number', 'reg no': 'Fleet Number',
        'meal hour': 'Meal Hour', 'meal': 'Meal Hour', 'meal hours': 'Meal Hour',
        'sleep out': 'Sleep Out', 'sleep': 'Sleep Out', 'sleepout': 'Sleep Out'
    }
    df.columns = [clean_col_name(col) for col in df.columns]
    for old, new in rename_map.items():
        df.columns = [new if old == col else col for col in df.columns]

    df = df.loc[:, ~df.columns.duplicated()]
    return df

def auto_lunch_deduction(row):
    """If Meal Hour blank/0, deduct 1h if trip covers 12:00-13:00 and >5h total"""
    meal = pd.to_numeric(row.get('Meal Hour', 0), errors='coerce')
    if pd.notna(meal) and meal > 0:
        return meal

    start = row['Start Time']
    end = row['End Time']
    gross = (end - start).total_seconds() / 3600

    # If trip >5h and spans midday 12:00-13:00, auto deduct 1h
    if gross > 5 and start.time() <= time(12, 0) and end.time() >= time(13, 0):
        return 1.0
    return 0.0

if tracking_file and allocation_file:
    try:
        # Read files
        df_track_raw = pd.read_excel(tracking_file, header=None)
        track_header = find_tracking_header(df_track_raw)
        df_track = pd.read_excel(tracking_file, header=track_header)
        df_track = standardize_columns(df_track)

        xls = pd.ExcelFile(allocation_file)
        all_alloc_dfs = []

        for sheet_name in xls.sheet_names:
            df_sheet = pd.read_excel(xls, sheet_name=sheet_name)
            df_sheet = standardize_columns(df_sheet)
            df_sheet['Employee Name'] = sheet_name
            all_alloc_dfs.append(df_sheet)

        if not all_alloc_dfs:
            st.error("No valid sheets found in allocation file")
            st.stop()

        df_alloc = pd.concat(all_alloc_dfs, ignore_index=True)

        # Date handling
        if 'Date' not in df_track.columns:
            if 'Start Time' in df_track.columns:
                df_track['Date'] = pd.to_datetime(df_track['Start Time'], errors='coerce')
            else:
                st.error(f"Tracking file missing 'Date' and 'Start Time'. Found: {list(df_track.columns)}")
                st.stop()

        df_track['Date'] = pd.to_datetime(df_track['Date'], errors='coerce', dayfirst=True)
        df_alloc['Date'] = pd.to_datetime(df_alloc['Date'], format='%Y %m %d', errors='coerce')

        df_track['Date'] = df_track['Date'].dt.strftime('%Y-%m-%d')
        df_alloc['Date'] = df_alloc['Date'].dt.strftime('%Y-%m-%d')

        # Standardize keys
        df_track['Fleet Number'] = df_track['Fleet Number'].astype(str).str.strip().str.upper()
        df_alloc['Fleet Number'] = df_alloc['Fleet Number'].astype(str).str.strip().str.upper()

        df_track = df_track.dropna(subset=['Fleet Number', 'Date', 'Start Time', 'End Time'])
        df_alloc = df_alloc.dropna(subset=['Fleet Number', 'Date'])

        if df_alloc.empty:
            st.error("Allocation data is empty after cleaning.")
            st.stop()

        df_merged = pd.merge(df_track, df_alloc, on=['Fleet Number', 'Date'], how='inner')

        if df_merged.empty:
            st.error("No matching rows.")
            st.stop()

        # ===== HOURS CALCULATION =====
        df_merged['Start Time'] = pd.to_datetime(df_merged['Start Time'], errors='coerce')
        df_merged['End Time'] = pd.to_datetime(df_merged['End Time'], errors='coerce')
        df_merged['gross_hours'] = (df_merged['End Time'] - df_merged['Start Time']).dt.total_seconds() / 3600

        # 1. Meal Hour: use column or auto-deduct 1h if worked through lunch
        df_merged['meal_hour'] = df_merged.apply(auto_lunch_deduction, axis=1)

        # 2. Paid Yard: Sleep Out + text mentions
        df_merged['sleep_out'] = pd.to_numeric(df_merged.get('Sleep Out', 0), errors='coerce').fillna(0)
        if 'Activity Description' in df_merged.columns:
            df_merged['yard_from_text'] = df_merged['Activity Description'].apply(extract_yard_hours_from_text)
        else:
            df_merged['yard_from_text'] = 0.0
        df_merged['yard_hours'] = df_merged['sleep_out'] + df_merged['yard_from_text']

        # 3. Paid hours = gross - unpaid meal
        df_merged['total_hours'] = (df_merged['gross_hours'] - df_merged['meal_hour']).clip(lower=0)
        df_merged['driving_hours'] = (df_merged['total_hours'] - df_merged['yard_hours']).clip(lower=0)

        # Sort
        df_merged = df_merged.sort_values(['Employee Name', 'Date', 'Start Time'])

        # ===== SUMMARY WITH 195.03 RULE =====
        driver_totals = df_merged.groupby('Employee Name').agg({
            'total_hours': 'sum', # already excludes meal
            'yard_hours': 'sum',
            'driving_hours': 'sum',
            'meal_hour': 'sum' # show unpaid total for reference
        }).reset_index()

        driver_totals['NORMAL HOURS'] = driver_totals['total_hours'].clip(upper=NORMAL_HOURS_THRESHOLD)
        driver_totals['OVERTIME @1.5'] = (driver_totals['total_hours'] - NORMAL_HOURS_THRESHOLD).clip(lower=0)
        driver_totals = driver_totals.rename(columns={
            'total_hours': 'PAID HOURS',
            'yard_hours': 'YARD',
            'driving_hours': 'DRIVING',
            'meal_hour': 'UNPAID MEAL'
        })

        st.success(f"Allocated {len(df_merged)} trips to {df_merged['Employee Name'].nunique()} drivers!")

        st.subheader(f"Summary - 07:00-17:00 shift, 1h lunch deducted, Overtime after {NORMAL_HOURS_THRESHOLD}h")
        st.dataframe(driver_totals[['Employee Name', 'PAID HOURS', 'NORMAL HOURS', 'OVERTIME @1.5', 'YARD', 'DRIVING', 'UNPAID MEAL']].round(2))

        st.subheader("All Trips")
        display_cols = ['Date', 'Employee Name', 'Fleet Number', 'Start Time', 'End Time',
                       'gross_hours', 'meal_hour', 'total_hours', 'yard_hours', 'driving_hours', 'Sleep Out']
        display_cols = [col for col in display_cols if col in df_merged.columns]
        st.dataframe(df_merged[display_cols].round(2))

        # ===== EXCEL EXPORT =====
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            driver_totals[['Employee Name', 'PAID HOURS', 'NORMAL HOURS', 'OVERTIME @1.5', 'YARD', 'DRIVING', 'UNPAID MEAL']].to_excel(
                writer, index=False, sheet_name='SUMMARY'
            )

            df_merged[display_cols].to_excel(writer, index=False, sheet_name='ALL TRIPS')

            # One sheet per driver
            for driver in sorted(df_merged['Employee Name'].unique()):
                df_driver = df_merged[df_merged['Employee Name'] == driver][display_cols].copy()

                subtotal = pd.DataFrame({
                    'Date': ['TOTAL'],
                    'Employee Name': [driver],
                    'Fleet Number': [''],
                    'Start Time': [pd.NaT],
                    'End Time': [pd.NaT],
                    'gross_hours': [df_driver['gross_hours'].sum()],
                    'meal_hour': [df_driver['meal_hour'].sum()],
                    'total_hours': [df_driver['total_hours'].sum()],
                    'yard_hours': [df_driver['yard_hours'].sum()],
                    'driving_hours': [df_driver['driving_hours'].sum()],
                    'Sleep Out': [df_driver.get('Sleep Out', pd.Series([0])).sum()]
                })

                subtotal['NORMAL HOURS'] = subtotal['total_hours'].clip(upper=NORMAL_HOURS_THRESHOLD)
                subtotal['OVERTIME @1.5'] = (subtotal['total_hours'] - NORMAL_HOURS_THRESHOLD).clip(lower=0)

                df_driver_out = pd.concat([df_driver, subtotal], ignore_index=True)
                sheet_name = re.sub(r'[\\/*?:\[\]]', '', driver)[:31]
                df_driver_out.to_excel(writer, index=False, sheet_name=sheet_name)

        output.seek(0)

        st.download_button(
            "📥 Download Excel - 1 Sheet Per Driver + Summary",
            output,
            "fleet_timesheet_processed.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {str(e)}")
else:
    st.info("👆 Upload both files to start processing")
