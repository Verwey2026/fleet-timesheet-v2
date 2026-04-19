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
st.set_page_config(page_title="Fleet Timesheet Processor V4.12", layout="wide")
st.title("Fleet Timesheet Processor VERSION 4.12 - Verwey Vervoer")

st.markdown("**Rules:** 1h unpaid lunch. First 195.03h = Normal. Sat @1.5, Sun @2.0. Counts Local/Xborder sleep outs + Abnormal truck days.")

NORMAL_HOURS_THRESHOLD = 195.03
GEO_FENCE_TOWNS = ['STEVE TSHWETE', 'MIDDELBURG', 'WITBANK', 'EMALAHLENI'] # TODO: Update your geo fence
CROSSBORDER_COUNTRIES = ['ZIMBABWE', 'BOTSWANA', 'NAMIBIA', 'MOZAMBIQUE', 'ZAMBIA', 'DRC', 'LESOTHO', 'ESWATINI', 'MALAWI', 'TANZANIA']
ABNORMAL_FLEETS = ['FL221', 'FL222', 'FL223', 'FL225', 'FL229', 'FL230', 'FL238'] # TODO: Add full list

st.sidebar.subheader("NBCRFLI Settings")
NORMAL_HOURS_THRESHOLD = st.sidebar.number_input("Normal Hours Threshold", value=195.03, step=0.01)
st.sidebar.markdown("**Abnormal Fleets:** " + ", ".join(ABNORMAL_FLEETS))

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload Tracking Report")
    tracking_file = st.file_uploader("Must have DEPARTURE + ARRIVAL columns", type=["xlsx", "xls", "csv"], key="tracking")

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
        'sleep out': 'Sleep Out', 'sleep': 'Sleep Out', 'sleepout': 'Sleep Out',
        'arrival location': 'Arrival', 'arrival town': 'Arrival', 'destination': 'Arrival',
        'departure location': 'Departure', 'departure town': 'Departure', 'origin': 'Departure'
    }
    df.columns = [clean_col_name(col) for col in df.columns]
    for old, new in rename_map.items():
        df.columns = [new if old == col else col for col in df.columns]

    df = df.loc[:, ~df.columns.duplicated()]
    return df

def auto_lunch_deduction(row):
    meal = pd.to_numeric(row.get('Meal Hour', 0), errors='coerce')
    if pd.notna(meal) and meal > 0:
        return meal
    start = row['Start Time']
    end = row['End Time']
    gross = (end - start).total_seconds() / 3600
    if gross > 5 and start.time() <= time(12, 0) and end.time() >= time(13, 0):
        return 1.0
    return 0.0

def classify_sleep_out(row):
    nights = pd.to_numeric(row.get('Sleep Out', 0), errors='coerce')
    if pd.isna(nights) or nights == 0:
        return 'none', 0

    arrival = str(row.get('Arrival', '')).upper().strip()

    if any(country in arrival for country in CROSSBORDER_COUNTRIES):
        return 'crossborder', nights
    if any(town in arrival for town in GEO_FENCE_TOWNS):
        return 'geofence', 0
    return 'local', nights

def is_abnormal(fleet_no):
    return any(fleet_no.startswith(prefix) for prefix in ABNORMAL_FLEETS)

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

        df_track['Date_dt'] = df_track['Date']
        df_alloc['Date_dt'] = df_alloc['Date']
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

        df_merged['meal_hour'] = df_merged.apply(auto_lunch_deduction, axis=1)

        # Sleep out classification
        sleep_data = df_merged.apply(classify_sleep_out, axis=1, result_type='expand')
        df_merged['sleep_out_type'] = sleep_data[0]
        df_merged['sleep_out_nights'] = sleep_data[1]
        df_merged['sleep_out_local'] = df_merged.apply(lambda x: x['sleep_out_nights'] if x['sleep_out_type'] == 'local' else 0, axis=1)
        df_merged['sleep_out_crossborder'] = df_merged.apply(lambda x: x['sleep_out_nights'] if x['sleep_out_type'] == 'crossborder' else 0, axis=1)

        # Yard hours from text only
        if 'Activity Description' in df_merged.columns:
            df_merged['yard_hours'] = df_merged['Activity Description'].apply(extract_yard_hours_from_text)
        else:
            df_merged['yard_hours'] = 0.0

        df_merged['total_hours'] = (df_merged['gross_hours'] - df_merged['meal_hour']).clip(lower=0)
        df_merged['driving_hours'] = (df_merged['total_hours'] - df_merged['yard_hours']).clip(lower=0)

        df_merged['weekday'] = pd.to_datetime(df_merged['Date']).dt.dayofweek
        df_merged['is_abnormal'] = df_merged['Fleet Number'].apply(is_abnormal)

        df_merged = df_merged.sort_values(['Employee Name', 'Date', 'Start Time'])

        # ===== RUNNING 195.03 SPLIT WITH DAY RATES =====
        df_merged['normal_weekday'] = 0.0
        df_merged['normal_sat'] = 0.0
        df_merged['normal_sun'] = 0.0
        df_merged['ot_weekday'] = 0.0
        df_merged['ot_sat'] = 0.0
        df_merged['ot_sun'] = 0.0

        for driver in df_merged['Employee Name'].unique():
            mask = df_merged['Employee Name'] == driver
            driver_idx = df_merged.loc[mask].index
            cumulative_normal = 0.0

            for idx in driver_idx:
                hours_today = df_merged.at[idx, 'total_hours']
                weekday = df_merged.at[idx, 'weekday']

                remaining_normal = max(0, NORMAL_HOURS_THRESHOLD - cumulative_normal)
                normal_today = min(hours_today, remaining_normal)
                ot_today = hours_today - normal_today

                if weekday == 6: # Sunday
                    df_merged.at[idx, 'normal_sun'] = normal_today
                    df_merged.at[idx, 'ot_sun'] = ot_today
                elif weekday == 5: # Saturday
                    df_merged.at[idx, 'normal_sat'] = normal_today
                    df_merged.at[idx, 'ot_sat'] = ot_today
                else: # Weekday
                    df_merged.at[idx, 'normal_weekday'] = normal_today
                    df_merged.at[idx, 'ot_weekday'] = ot_today

                cumulative_normal += normal_today

        # ===== ABNORMAL TRUCK DAY COUNT =====
        abnormal_days = df_merged[df_merged['is_abnormal'] == True].groupby('Employee Name')['Date'].nunique().reset_index()
        abnormal_days = abnormal_days.rename(columns={'Date': 'ABNORMAL DAYS'})

        # ===== SUMMARY =====
        driver_totals = df_merged.groupby('Employee Name').agg({
            'total_hours': 'sum',
            'normal_weekday': 'sum',
            'normal_sat': 'sum',
            'normal_sun': 'sum',
            'ot_weekday': 'sum',
            'ot_sat': 'sum',
            'ot_sun': 'sum',
            'yard_hours': 'sum',
            'driving_hours': 'sum',
            'meal_hour': 'sum',
            'sleep_out_local': 'sum',
            'sleep_out_crossborder': 'sum'
        }).reset_index()

        driver_totals = pd.merge(driver_totals, abnormal_days, on='Employee Name', how='left')
        driver_totals['ABNORMAL DAYS'] = driver_totals['ABNORMAL DAYS'].fillna(0).astype(int)

        driver_totals = driver_totals.rename(columns={
            'total_hours': 'PAID HOURS',
            'normal_weekday': 'NORMAL WD',
            'normal_sat': 'NORMAL SAT@1.5',
            'normal_sun': 'NORMAL SUN@2.0',
            'ot_weekday': 'OT WD@1.5',
            'ot_sat': 'OT SAT@1.5',
            'ot_sun': 'OT SUN@2.0',
            'yard_hours': 'YARD',
            'driving_hours': 'DRIVING',
            'meal_hour': 'UNPAID MEAL',
            'sleep_out_local': 'SLEEP OUT LOCAL',
            'sleep_out_crossborder': 'SLEEP OUT XBORDER'
        })

        st.success(f"Allocated {len(df_merged)} trips to {df_merged['Employee Name'].nunique()} drivers!")

        st.subheader(f"Summary - Includes Abnormal Days + Sleep Out Counts")
        st.dataframe(driver_totals.round(2))

        st.subheader("All Trips with Abnormal Flag")
        display_cols = ['Date', 'Employee Name', 'Fleet Number', 'is_abnormal', 'Arrival', 'weekday', 'Start Time', 'End Time',
                       'total_hours', 'normal_weekday', 'normal_sat', 'normal_sun',
                       'ot_weekday', 'ot_sat', 'ot_sun', 'yard_hours', 'meal_hour',
                       'sleep_out_local', 'sleep_out_crossborder']
        display_cols = [col for col in display_cols if col in df_merged.columns]
        st.dataframe(df_merged[display_cols].round(2))

        # ===== EXCEL EXPORT =====
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            driver_totals.to_excel(writer, index=False, sheet_name='SUMMARY')
            df_merged[display_cols].to_excel(writer, index=False, sheet_name='ALL TRIPS')

            for driver in sorted(df_merged['Employee Name'].unique()):
                df_driver = df_merged[df_merged['Employee Name'] == driver][display_cols].copy()

                subtotal_data = {col: [''] for col in display_cols}
                subtotal_data['Date'] = ['TOTAL']
                subtotal_data['Employee Name'] = [driver]
                for col in ['total_hours', 'normal_weekday', 'normal_sat', 'normal_sun',
                           'ot_weekday', 'ot_sat', 'ot_sun', 'yard_hours', 'meal_hour',
                           'sleep_out_local', 'sleep_out_crossborder']:
                    if col in df_driver.columns:
                        subtotal_data[col] = [df_driver[col].sum()]

                subtotal = pd.DataFrame(subtotal_data)
                df_driver_out = pd.concat([df_driver, subtotal], ignore_index=True)
                sheet_name = re.sub(r'[\\/*?:\[\]]', '', driver)[:31]
                df_driver_out.to_excel(writer, index=False, sheet_name=sheet_name)

        output.seek(0)

        st.download_button(
            "📥 Download Excel - With Abnormal Days",
            output,
            "fleet_timesheet_processed.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {str(e)}")
        st.exception(e)
else:
    st.info("👆 Upload both files to start processing")
