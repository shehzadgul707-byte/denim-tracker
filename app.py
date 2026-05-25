import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, date
import os

# Set layout to wide for professional R&D dashboard look
st.set_page_config(page_title="Denim R&D Advanced Tracker", layout="wide")

EXCEL_FILE = "Denim_Master_Database.xlsx"

# Standard fixed dropdown lists
SOURCE_OPTIONS = ["Existing", "Scratch", "Production beam"]
STATUS_OPTIONS = ["Yarn Demand", "Dyeing", "Weaving", "Finishing", "Inspection", "Submitted to marketing"]

# Standard precise structure requested by the user
COLUMNS_STRUCTURE = [
    "S.no", "Brand", "Ref", "Finish", "TR Code", "Source", "Pending Days in Process",
    "Request Date", "Current Date", "Delivery date", "Ready date", "Status", "Alert",
    "Warp", "Warp Slub", "Weft", "Shade", "Weave", "Reed", "Picks", "Greige Meters", "Remarks"
]

# --- HELPER FUNCTIONS FOR EXCEL HANDLING ---
def load_all_sheets_clean():
    """Forces the system to always initialize to ZERO samples for a fresh user upload experience."""
    df_master = pd.DataFrame(columns=COLUMNS_STRUCTURE)
    df_trials = pd.DataFrame(columns=["Sample ID", "Trial Number", "Parameters", "Status_OK", "Updated Date"])
    return df_master, df_trials

def save_all_sheets(df_master, df_trials):
    """Saves both dataframes into separate sheets in the same Excel file."""
    if len(df_master) > 0:
        df_master["S.no"] = df_master["S.no"].astype(str).apply(lambda x: str(x).split('.')[0].strip())
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master_Data", index=False)
        df_trials.to_excel(writer, sheet_name="Trial_Data", index=False)

# Force load an absolute clean state with ZERO rows 
df_master, df_trials = load_all_sheets_clean()

# If an Excel file is already saved and has rows, we overwrite it with fresh uploads later
if os.path.exists(EXCEL_FILE) and not st.session_state.get("fresh_uploaded", False):
    try:
        temp_master = pd.read_excel(EXCEL_FILE, sheet_name="Master_Data")
        if len(temp_master) > 0:
            temp_master["S.no"] = temp_master["S.no"].astype(str).apply(lambda x: str(x).split('.')[0].strip())
            if "Ppi" in temp_master.columns and "Picks" not in temp_master.columns:
                temp_master.rename(columns={"Ppi": "Picks"}, inplace=True)
            for col in COLUMNS_STRUCTURE:
                if col not in temp_master.columns: temp_master[col] = ""
            df_master = temp_master[COLUMNS_STRUCTURE].copy()
    except:
        pass

# --- AUTOMATIC FORMULA CALCULATIONS ---
def calculate_metrics_and_alerts(df):
    if len(df) == 0:
        return df
    today = date.today()
    df['Current Date'] = today.strftime("%Y-%m-%d")
    df["S.no"] = df["S.no"].astype(str).apply(lambda x: str(x).split('.')[0].strip())
    
    for idx, row in df.iterrows():
        try:
            req_date = pd.to_datetime(row['Request Date']).date()
            df.at[idx, 'Pending Days in Process'] = (today - req_date).days
        except:
            df.at[idx, 'Pending Days in Process'] = 0
            
        if row['Status'] != "Submitted to marketing":
            try:
                del_date = pd.to_datetime(row['Delivery date']).date()
                days_left = (del_date - today).days
                if days_left <= 3 and days_left >= 0:
                    df.at[idx, 'Alert'] = "⚠️ Critical"
                elif days_left < 0:
                    df.at[idx, 'Alert'] = "🚨 Overdue"
                else:
                    df.at[idx, 'Alert'] = "Normal"
            except:
                df.at[idx, 'Alert'] = "No Delivery Date"
        else:
            df.at[idx, 'Alert'] = "Completed"
            
    return df

df_master = calculate_metrics_and_alerts(df_master)

# Filter active running samples vs archived ones
df_running = df_master[df_master["Status"] != "Submitted to marketing"].copy()

# Categorize into Development vs Repeat based on source logic
df_dev = df_running[df_running["Source"].str.lower() == "scratch"].copy()
df_repeat = df_running[df_running["Source"].str.lower().isin(["existing", "production beam"])].copy()

# --- TOP MAIN HEADINGS & METRIC BOXES ---
st.title("👖 Denim Fabric R&D Production Pipeline Tracker")
st.markdown("---")

st.markdown("### 📊 Live Floor Workload Counters")
kpi_box1, kpi_box2, kpi_box3 = st.columns(3)

# Short clean HTML injections to block layout corruption
html_floor = f"<div style='background-color:#1E3A8A; padding:20px; border-radius:10px; text-align:center; color:white;'><h2>Total Samples On Floor</h2><p style='font-size:35px; font-weight:bold; margin:0;'>{len(df_running)}</p></div>"
html_dev = f"<div style='background-color:#0D9488; padding:20px; border-radius:10px; text-align:center; color:white;'><h2>Total Development Section</h2><p style='font-size:35px; font-weight:bold; margin:0;'>{len(df_dev)}</p></div>"
html_repeat = f"<div style='background-color:#B45309; padding:20px; border-radius:10px; text-align:center; color:white;'><h2>Total Repeat Section</h2><p style='font-size:35px; font-weight:bold; margin:0;'>{len(df_repeat)}</p></div>"

with kpi_box1:
    st.markdown(html_floor, unsafe_allow_html=True)

with kpi_box2:
    st.markdown(html_dev, unsafe_allow_html=True)

with kpi_box3:
    st.markdown(html_repeat, unsafe_allow_html=True)

st.markdown("---")

# --- CLEAR ALL & REFRESH CONTROL PANEL ---
st.markdown("## ⚙️ Control Panel: Reset and Upload Data")
cc1, cc2 = st.columns([2, 4])
with cc1:
    if st.button("🚨 Wipe Out Everything (Set Dashboard to 0)", use_container_width=True):
        df_master, df_trials = load_all_sheets_clean()
        save_all_sheets(df_master, df_trials)
        st.success("Saara data zero kar diya gaya hai!")
        st.rerun()

st.markdown("---")

# --- FRESH FILE UPLOAD BLOCK CENTER (REPLACES OLD DATA) ---
st.markdown("## 📥 Fresh Excel Upload Center")
with st.container():
    st.info("💡 Yahan apni Excel sheet select karein aur niche button daba kar data load karein.")
    uploaded_file = st.file_uploader("Excel (.xlsx) sheet browse karein:", type=["xlsx"], key="fresh_excel_uploader")
    
    if uploaded_file is not None:
        if st.button("🚀 Process & Replace Pipeline Data", use_container_width=True):
            try:
                imported_df = pd.read_excel(uploaded_file)
                required_cols = ["S.no", "Brand", "Ref", "Status"]
                missing_cols = [c for c in required_cols if c not in imported_df.columns]
                
                if missing_cols:
                    st.error(f"Excel File mein yeh basic headers hona lazmi hain: {missing_cols}")
                else:
                    imported_df["S.no"] = imported_df["S.no"].astype(str).apply(lambda x: str(x).split('.')[0].strip())
                    imported_df["Ref"] = imported_df["Ref"].astype(str).str.strip()
                    if "Ppi" in imported_df.columns and "Picks" not in imported_df.columns: 
                        imported_df.rename(columns={"Ppi": "Picks"}, inplace=True)
                    
                    # Row data filters (valid rows only)
                    valid_mask = imported_df["Ref"].notna() & (imported_df["Ref"] != "")
                    filtered_imp = imported_df[valid_mask].copy()
                    
                    # Force clean layout arrays
                    blank_master
