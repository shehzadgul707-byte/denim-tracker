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

# ====================== STRONG FIX FUNCTIONS ======================
def force_all_string(df):
    """Har column ko safely string banao"""
    df = df.copy()
    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()
    return df

def save_all_sheets(df_master, df_trials):
    df_master = force_all_string(df_master)
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master_Data", index=False)
        df_trials.to_excel(writer, sheet_name="Trial_Data", index=False)

def calculate_metrics_and_alerts(df):
    if len(df) == 0:
        return df.copy()
    
    df = force_all_string(df)   # Strong fix
    
    today = date.today()
    df['Current Date'] = today.strftime("%Y-%m-%d")
    
    for idx in df.index:
        try:
            req_date = pd.to_datetime(df.at[idx, 'Request Date']).date()
            df.at[idx, 'Pending Days in Process'] = (today - req_date).days
        except:
            df.at[idx, 'Pending Days in Process'] = 0
        
        status = str(df.at[idx, 'Status']).strip()
        if status != "Submitted to marketing":
            try:
                del_date = pd.to_datetime(df.at[idx, 'Delivery date']).date()
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
        df_master, df_trials = pd.DataFrame(columns=COLUMNS_STRUCTURE), pd.DataFrame()
else:
    df_master, df_trials = pd.DataFrame(columns=COLUMNS_STRUCTURE), pd.DataFrame()

df_master = force_all_string(df_master)
df_master = df_master[COLUMNS_STRUCTURE].copy()
df_master = calculate_metrics_and_alerts(df_master)

df_running = df_master[df_master["Status"] != "Submitted to marketing"].copy()
df_dev = df_running[df_running["Source"].str.contains("scratch", case=False, na=False)].copy()
df_repeat = df_running[df_running["Source"].str.contains("existing|production", case=False, na=False)].copy()

# ====================== UI ======================
st.title("👖 Denim Fabric R&D Production Pipeline Tracker")
st.markdown("---")

c1, c2, c3 = st.columns(3)
with c1: st.markdown(f"<div style='background:#1E3A8A;padding:25px;border-radius:12px;text-align:center;color:white;'><h3>Total On Floor</h3><h1>{len(df_running)}</h1></div>", unsafe_allow_html=True)
with c2: st.markdown(f"<div style='background:#0D9488;padding:25px;border-radius:12px;text-align:center;color:white;'><h3>Development</h3><h1>{len(df_dev)}</h1></div>", unsafe_allow_html=True)
with c3: st.markdown(f"<div style='background:#B45309;padding:25px;border-radius:12px;text-align:center;color:white;'><h3>Repeat</h3><h1>{len(df_repeat)}</h1></div>", unsafe_allow_html=True)

st.markdown("---")

if st.button("🚨 Wipe Out Everything", use_container_width=True):
    df_master, df_trials = pd.DataFrame(columns=COLUMNS_STRUCTURE), pd.DataFrame()
    save_all_sheets(df_master, df_trials)
    st.session_state.fresh_uploaded = False
    st.success("Sab clear ho gaya!")
    st.rerun()

# ====================== UPLOAD (FINAL STRONG FIX) ======================
st.markdown("## 📥 Fresh Excel Upload")
uploaded_file = st.file_uploader("Apni Excel file yahan upload karein", type=["xlsx"])

if uploaded_file is not None:
    if st.button("🚀 Process & Replace All Data", type="primary", use_container_width=True):
        try:
            imported_df = pd.read_excel(uploaded_file)
            
            # Strong cleaning from the beginning
            imported_df = force_all_string(imported_df)
            
            if "S.no" not in imported_df.columns or "Ref" not in imported_df.columns:
                st.error("S.no aur Ref columns hone chahiye")
            else:
                # Rename Ppi to Picks if exists
                if "Ppi" in imported_df.columns:
                    imported_df.rename(columns={"Ppi": "Picks"}, inplace=True)
                
                # Valid rows
                valid_df = imported_df[
                    (imported_df["Ref"].str.strip() != "") & 
                    (imported_df["Ref"].str.lower() != "nan")
                ].copy()
                
                # Add all required columns
                for col in COLUMNS_STRUCTURE:
                    if col not in valid_df.columns:
                        valid_df[col] = ""
                
                final_df = valid_df[COLUMNS_STRUCTURE].copy()
                final_df = force_all_string(final_df)        # ← Yeh line sabse zaroori hai
                final_df = calculate_metrics_and_alerts(final_df)
                
                save_all_sheets(final_df, pd.DataFrame())
                st.session_state.fresh_uploaded = True
                
                st.success(f"🎉 {len(final_df)} samples successfully loaded!")
                st.rerun()
                
        except Exception as e:
            st.error(f"Upload Error: {str(e)}")

st.markdown("---")

# Baaki parts (Tabs, Manual Entry, Archive) — agar chahiye to batao main poora add kar dun.
st.info("Ab upload karke dekho. Yeh version mein maine har column ko string force kiya hai.")
