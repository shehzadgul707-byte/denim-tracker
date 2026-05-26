import streamlit as st
import pandas as pd
from datetime import date
import os

st.set_page_config(page_title="Denim R&D Advanced Tracker", layout="wide")

EXCEL_FILE = "Denim_Master_Database.xlsx"

SOURCE_OPTIONS = ["Existing", "Scratch", "Production beam"]
STATUS_OPTIONS = ["Yarn Demand", "Dyeing", "Weaving", "Finishing", "Inspection", "Submitted to marketing"]

COLUMNS_STRUCTURE = [
    "S.no", "Brand", "Ref", "Finish", "TR Code", "Source", "Pending Days in Process",
    "Request Date", "Current Date", "Delivery date", "Ready date", "Status", "Alert",
    "Warp", "Warp Slub", "Weft", "Shade", "Weave", "Reed", "Picks", "Greige Meters", "Remarks"
]

if "fresh_uploaded" not in st.session_state:
    st.session_state.fresh_uploaded = False

# ====================== HELPER FUNCTIONS ======================
def load_all_sheets_clean():
    df_master = pd.DataFrame(columns=COLUMNS_STRUCTURE)
    df_trials = pd.DataFrame(columns=["Sample ID", "Trial Number", "Parameters", "Status_OK", "Updated Date"])
    return df_master, df_trials

def save_all_sheets(df_master, df_trials):
    df_master = df_master.copy()
    df_master["S.no"] = df_master["S.no"].astype(str).str.strip()
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master_Data", index=False)
        df_trials.to_excel(writer, sheet_name="Trial_Data", index=False)

def calculate_metrics_and_alerts(df):
    if len(df) == 0:
        return df.copy()
    
    df = df.copy()
    df["S.no"] = df["S.no"].astype(str).str.strip()
    
    today = date.today()
    df['Current Date'] = today.strftime("%Y-%m-%d")
    
    for idx, row in df.iterrows():
        try:
            req_date = pd.to_datetime(row['Request Date']).date()
            df.at[idx, 'Pending Days in Process'] = (today - req_date).days
        except:
            df.at[idx, 'Pending Days in Process'] = 0
        
        if row.get('Status') != "Submitted to marketing":
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

# ====================== LOAD DATA ======================
if os.path.exists(EXCEL_FILE) and not st.session_state.fresh_uploaded:
    try:
        df_master = pd.read_excel(EXCEL_FILE, sheet_name="Master_Data")
        df_trials = pd.read_excel(EXCEL_FILE, sheet_name="Trial_Data")
    except:
        df_master, df_trials = load_all_sheets_clean()
else:
    df_master, df_trials = load_all_sheets_clean()

for col in COLUMNS_STRUCTURE:
    if col not in df_master.columns:
        df_master[col] = ""

df_master = df_master[COLUMNS_STRUCTURE].copy()
df_master = calculate_metrics_and_alerts(df_master)

df_running = df_master[df_master["Status"] != "Submitted to marketing"].copy()
df_dev = df_running[df_running["Source"].str.lower() == "scratch"].copy()
df_repeat = df_running[df_running["Source"].str.lower().isin(["existing", "production beam"])].copy()

# ====================== UI ======================
st.title("👖 Denim Fabric R&D Production Pipeline Tracker")
st.markdown("---")

# KPI
c1, c2, c3 = st.columns(3)
with c1: st.markdown(f"<div style='background:#1E3A8A;padding:20px;border-radius:10px;text-align:center;color:white;'><h2>Total On Floor</h2><p style='font-size:38px;'>{len(df_running)}</p></div>", unsafe_allow_html=True)
with c2: st.markdown(f"<div style='background:#0D9488;padding:20px;border-radius:10px;text-align:center;color:white;'><h2>Development</h2><p style='font-size:38px;'>{len(df_dev)}</p></div>", unsafe_allow_html=True)
with c3: st.markdown(f"<div style='background:#B45309;padding:20px;border-radius:10px;text-align:center;color:white;'><h2>Repeat</h2><p style='font-size:38px;'>{len(df_repeat)}</p></div>", unsafe_allow_html=True)

st.markdown("---")

if st.button("🚨 Wipe Out Everything", use_container_width=True):
    df_master, df_trials = load_all_sheets_clean()
    save_all_sheets(df_master, df_trials)
    st.session_state.fresh_uploaded = False
    st.success("Reset successful!")
    st.rerun()

# ====================== UPLOAD - STRONG FIX ======================
st.markdown("## 📥 Fresh Excel Upload")
uploaded_file = st.file_uploader("Excel file select karein", type=["xlsx"])

if uploaded_file is not None:
    if st.button("🚀 Process & Replace All Data", type="primary", use_container_width=True):
        try:
            imported_df = pd.read_excel(uploaded_file)
            
            if "S.no" not in imported_df.columns or "Ref" not in imported_df.columns:
                st.error("S.no aur Ref columns zaroori hain")
            else:
                imported_df = imported_df.copy()
                
                # 🔥 STRONG FIX FOR S.NO
                imported_df["S.no"] = (
                    imported_df["S.no"]
                    .fillna("")                    # NaN ko empty string
                    .astype(str)                   # string banao
                    .str.replace(r'\.0$', '', regex=True)   # .0 remove
                    .str.strip()
                )
                
                imported_df["Ref"] = imported_df["Ref"].astype(str).str.strip()
                
                if "Ppi" in imported_df.columns and "Picks" not in imported_df.columns:
                    imported_df.rename(columns={"Ppi": "Picks"}, inplace=True)
                
                # Valid rows
                valid_df = imported_df[
                    (imported_df["Ref"].notna()) & 
                    (imported_df["Ref"] != "") & 
                    (imported_df["Ref"].str.lower() != "nan")
                ].copy()
                
                # Add missing columns
                for col in COLUMNS_STRUCTURE:
                    if col not in valid_df.columns:
                        valid_df[col] = ""
                
                final_df = valid_df[COLUMNS_STRUCTURE].copy()
                
                # Final safety
                final_df["S.no"] = final_df["S.no"].astype(str).str.strip()
                final_df = calculate_metrics_and_alerts(final_df)
                
                save_all_sheets(final_df, pd.DataFrame(columns=df_trials.columns))
                st.session_state.fresh_uploaded = True
                
                st.success(f"🎉 {len(final_df)} samples loaded successfully!")
                st.rerun()
                
        except Exception as e:
            st.error(f"Error: {str(e)}")

st.markdown("---")

# Rest of the code (Tabs, Manual Entry, Archive) same as previous version
# ... (aap purane code se yeh part copy kar sakte hain)

st.info("Agar phir bhi error aaye to mujhe pura error message aur Excel file ka pehla 5 rows ka data batao.")
