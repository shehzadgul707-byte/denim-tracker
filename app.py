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

# ====================== STRONG STRING FIX ======================
def read_excel_as_string(file):
    """Excel ko pure string mode mein read karta hai"""
    df = pd.read_excel(file, dtype=str)  # Sab columns string
    return df

def clean_dataframe(df):
    df = df.copy()
    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()
    return df

def save_all_sheets(df_master, df_trials):
    df_master = clean_dataframe(df_master)
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master_Data", index=False)
        df_trials.to_excel(writer, sheet_name="Trial_Data", index=False)

def calculate_metrics_and_alerts(df):
    if len(df) == 0:
        return df.copy()
    
    df = clean_dataframe(df)
    
    today = date.today()
    df['Current Date'] = today.strftime("%Y-%m-%d")
    
    for idx in df.index:
        try:
            req_date = pd.to_datetime(df.at[idx, 'Request Date']).date()
            df.at[idx, 'Pending Days in Process'] = str((today - req_date).days)
        except:
            df.at[idx, 'Pending Days in Process'] = "0"
        
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

# ====================== LOAD SAVED DATA ======================
if os.path.exists(EXCEL_FILE) and not st.session_state.fresh_uploaded:
    try:
        df_master = pd.read_excel(EXCEL_FILE, sheet_name="Master_Data", dtype=str)
        df_trials = pd.read_excel(EXCEL_FILE, sheet_name="Trial_Data", dtype=str)
    except:
        df_master, df_trials = pd.DataFrame(columns=COLUMNS_STRUCTURE), pd.DataFrame()
else:
    df_master, df_trials = pd.DataFrame(columns=COLUMNS_STRUCTURE), pd.DataFrame()

df_master = clean_dataframe(df_master)
df_master = df_master.reindex(columns=COLUMNS_STRUCTURE, fill_value="")

df_master = calculate_metrics_and_alerts(df_master)

df_running = df_master[df_master["Status"] != "Submitted to marketing"].copy()
df_dev = df_running[df_running["Source"].str.contains("scratch", case=False, na=False)].copy()
df_repeat = df_running[df_running["Source"].str.contains("existing|production", case=False, na=False)].copy()

# ====================== MAIN UI ======================
st.title("👖 Denim Fabric R&D Production Pipeline Tracker")
st.markdown("---")

c1, c2, c3 = st.columns(3)
with c1: st.markdown(f"<div style='background:#1E3A8A;padding:20px;border-radius:10px;text-align:center;color:white;'><h2>Total On Floor</h2><h1>{len(df_running)}</h1></div>", unsafe_allow_html=True)
with c2: st.markdown(f"<div style='background:#0D9488;padding:20px;border-radius:10px;text-align:center;color:white;'><h2>Development</h2><h1>{len(df_dev)}</h1></div>", unsafe_allow_html=True)
with c3: st.markdown(f"<div style='background:#B45309;padding:20px;border-radius:10px;text-align:center;color:white;'><h2>Repeat</h2><h1>{len(df_repeat)}</h1></div>", unsafe_allow_html=True)

st.markdown("---")

if st.button("🚨 Wipe Out Everything", use_container_width=True):
    df_master, df_trials = pd.DataFrame(columns=COLUMNS_STRUCTURE), pd.DataFrame()
    save_all_sheets(df_master, df_trials)
    st.session_state.fresh_uploaded = False
    st.success("Dashboard Reset Ho Gaya!")
    st.rerun()

# ====================== UPLOAD ======================
st.markdown("## 📥 Fresh Excel Upload")
uploaded_file = st.file_uploader("Excel File Upload Karein", type=["xlsx"])

if uploaded_file is not None:
    if st.button("🚀 Process & Replace All Data", type="primary", use_container_width=True):
        try:
            imported_df = read_excel_as_string(uploaded_file)   # ← Yeh line sabse important hai
            imported_df = clean_dataframe(imported_df)
            
            if "S.no" not in imported_df.columns or "Ref" not in imported_df.columns:
                st.error("S.no aur Ref columns hone chahiye")
            else:
                if "Ppi" in imported_df.columns and "Picks" not in imported_df.columns:
                    imported_df.rename(columns={"Ppi": "Picks"}, inplace=True)
                
                # Valid rows
                valid_df = imported_df[imported_df["Ref"].str.strip() != ""].copy()
                
                for col in COLUMNS_STRUCTURE:
                    if col not in valid_df.columns:
                        valid_df[col] = ""
                
                final_df = valid_df[COLUMNS_STRUCTURE].copy()
                final_df = calculate_metrics_and_alerts(final_df)
                
                save_all_sheets(final_df, pd.DataFrame())
                st.session_state.fresh_uploaded = True
                
                st.success(f"🎉 {len(final_df)} samples successfully load ho gaye!")
                st.rerun()
                
        except Exception as e:
            st.error(f"Upload Error: {str(e)}")

st.markdown("---")
st.info("**Note:** Agar phir bhi error aaye to mujhe batao. Ab yeh code bahut strong hai.")

# Baaki features (data editor, manual entry etc.) baad mein add kar sakte hain agar yeh upload kaam kar jaye.
