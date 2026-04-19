import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta

st.set_page_config(page_title="Fleet Timesheet Processor V2.4", layout="wide")

# --- PASSWORD GATE ---
def check_password():
    """Returns True if user entered correct password."""
    
    def password_entered():
        if st.session_state["password"] == st.secrets["app_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # don't store password
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password.
        st.text_input("Password", type="password", on_change=password_entered, key="password")
        st.caption("This app is private.")
        return False
    elif not st.session_state["password_correct"]:
        # Password not correct, show input + error.
        st.text_input("Password", type="password", on_change=password_entered, key="password")
        st.error("😕 Password incorrect")
        return False
    else:
        # Password correct.
        return True

if not check_password():
    st.stop()  # Don't run rest of app unless password correct

# --- MAIN APP STARTS HERE ---
st.title("Fleet Timesheet Processor")
st.caption("**VERSION 2.4 - Yard Hours + Fleet Number Extraction**")

uploaded_file = st.file_uploader("Upload Timesheet Excel File", type=["xlsx", "xls"])

def extract_fleet_number(text):
    """Extract fleet number like 31343, 31478 from text like '3.1 Vehicle Travel 31343 to...'"""
    if pd.isna(text):
        return ""
    text = str(text)
    import re
    match = re.search(r'\b(\d{5})\b', text)
    if match:
        return match.group(1)
    match = re.search(r'\b(\d{4,6})\b', text)
    if match:
        return match.group(1)
    return ""

def process_timesheet(df):
    df = df.copy()
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
    
    rename_map = {
        'date': 'date', 'employee': 'employee_name', 'employee_name': 'employee_name', 
        'name': 'employee_name', 'activity': 'activity_description', 'description': 'activity_description',
        'activity_description': 'activity_description', 'start': 'start_time', 'start_time': 'start_time',
        'end': 'end_time', 'end_time': 'end_time'
    }
    
    for old, new in rename_map.items():
        if old in df.columns:
            df = df.rename(columns={old: new})
    
    required_cols = ['date', 'employee_name', 'activity_description', 'start_time', 'end_time']
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Missing required columns: {missing}. Found columns: {list(df.columns)}")
        st.stop()
    
    df['start_time'] = pd.to_datetime(df['start_time'], errors='coerce').dt.time
    df['end_time'] = pd.to_datetime(df['end_time'], errors='coerce').dt.time
    df['date'] = pd.to_datetime(df['date'], errors='coerce').dt.date
    df = df.dropna(subset=['start_time', 'end_time', 'date'])
    
    def calc_hours(row):
        start_dt = datetime.combine(datetime.today(), row['start_time'])
        end_dt = datetime.combine(datetime.today(), row['end_time'])
        if end_dt < start_dt:
            end_dt += timedelta(days=1)
        delta = end_dt - start_dt
        return round(delta.total_seconds() / 3600, 2)
    
    df['hours'] = df.apply(calc_hours, axis=1)
    df['fleet_number'] = df['activity_description'].apply(extract_fleet_number)
    
    yard_keywords = ['yard', 'workshop', 'depot', 'base']
    df['is_yard'] = df['activity_description'].str.lower().str.contains('|'.join(yard_keywords), na=False)
    df['yard_hours'] = df.apply(lambda r: r['hours'] if r['is_yard'] else 0, axis=1)
    df['field_hours'] = df.apply(lambda r: 0 if r['is_yard'] else r['hours'], axis=1)
    
    return df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Processed')
    output.seek(0)
    return output

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file)
        st.subheader("Raw Data Preview")
        st.dataframe(df_raw.head(10))
        
        df_processed = process_timesheet(df_raw)
        
        st.subheader("Processed Data Preview")
        st.dataframe(df_processed.head(20))
        
        summary = df_processed.groupby('employee_name').agg(
            total_hours=('hours', 'sum'),
            yard_hours=('yard_hours', 'sum'),
            field_hours=('field_hours', 'sum'),
            days_worked=('date', 'nunique')
        ).reset_index()
        
        st.subheader("Summary by Employee")
        st.dataframe(summary)
        
        excel_data = to_excel(df_processed)
        st.download_button(
            label="📥 Download Processed Excel",
            data=excel_data,
            file_name=f"processed_timesheet_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error processing file: {e}")
        st.exception(e)
else:
    st.info("Upload an Excel timesheet to start. Required columns: Date, Employee Name, Activity Description, Start Time, End Time")
