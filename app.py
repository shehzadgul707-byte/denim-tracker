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
        # Checking file history
        temp_master = pd.read_excel(EXCEL_FILE, sheet_name="Master_Data")
        if len(temp_master) > 0:
            # We standardize rows if file has data
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

# 1st Custom Container: Core Production Indicators
st.markdown("### 📊 Live Floor Workload Counters")
kpi_box1, kpi_box2, kpi_box3 = st.columns(3)

with kpi_box1:
    html_1 = f"<div style='background-color:#1E3A8A; padding:20px; border-radius:10px; text-align:center; color:white;'><h2>Total Samples On Floor</h2><p style='font-size:35px; font-weight:bold; margin:0;'>{len(df_running)}</p></div>"
    st.markdown(html_1, unsafe_allow_html=True)

with kpi_box2:
    html_2 = f"<div style='background-color:#0D9488; padding:20px; border-radius:10px; text-align:center; color:white;'><h2>Total Development Section</h2><p style='font-size:35px; font-weight:bold; margin:0;'>{len(df_dev)}</p></div>"
    st.markdown(html_2, unsafe_allow_html=True)

with kpi_
