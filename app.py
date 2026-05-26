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
def read_excel_as_string(file_path_or_buffer):
    return pd.read_excel(file_path_or_buffer, dtype=str)

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

# ====================== LOAD DATA ======================
@st.cache_data(ttl=5)  # Fresh data load
def load_data():
    if os.path.exists(EXCEL_FILE):
        try:
            df_master = pd.read_excel(EXCEL_FILE, sheet_name="Master_Data", dtype=str)
        except:
            df_master = pd.DataFrame(columns=COLUMNS_STRUCTURE)
    else:
        df_master = pd.DataFrame(columns=COLUMNS_STRUCTURE)
    
    df_master = clean_dataframe(df_master)
    for col in COLUMNS_STRUCTURE:
        if col not in df_master.columns:
            df_master[col] = ""
    
    df_master = df_master[COLUMNS_STRUCTURE].copy()
    df_master = calculate_metrics_and_alerts(df_master)
    return df_master

df_master = load_data()

df_running = df_master[df_master["Status"] != "Submitted to marketing"].copy()
df_dev = df_running[df_running["Source"].str.contains("scratch", case=False, na=False)].copy()
df_repeat = df_running[df_running["Source"].str.contains("existing|production", case=False, na=False)].copy()

# ====================== UI ======================
st.title("👖 Denim Fabric R&D Production Pipeline Tracker")
st.markdown("---")

c1, c2, c3 = st.columns(3)
with c1: st.markdown(f"<div style='background:#1E3A8A;padding:20px;border-radius:10px;text-align:center;color:white;'><h2>Total On Floor</h2><h1>{len(df_running)}</h1></div>", unsafe_allow_html=True)
with c2: st.markdown(f"<div style='background:#0D9488;padding:20px;border-radius:10px;text-align:center;color:white;'><h2>Development</h2><h1>{len(df_dev)}</h1></div>", unsafe_allow_html=True)
with c3: st.markdown(f"<div style='background:#B45309;padding:20px;border-radius:10px;text-align:center;color:white;'><h2>Repeat</h2><h1>{len(df_repeat)}</h1></div>", unsafe_allow_html=True)

st.markdown("---")

# Upload Section
st.markdown("## 📥 Fresh Excel Upload")
uploaded_file = st.file_uploader("Excel File Upload Karein", type=["xlsx"], key="uploader")

if uploaded_file is not None:
    if st.button("🚀 Process & Replace All Data", type="primary", use_container_width=True):
        try:
            imported_df = read_excel_as_string(uploaded_file)
            imported_df = clean_dataframe(imported_df)
            
            if "Ref" not in imported_df.columns:
                st.error("Ref column nahi mila")
            else:
                if "Ppi" in imported_df.columns and "Picks" not in imported_df.columns:
                    imported_df.rename(columns={"Ppi": "Picks"}, inplace=True)
                
                valid_df = imported_df[imported_df["Ref"].str.strip() != ""].copy()
                
                for col in COLUMNS_STRUCTURE:
                    if col not in valid_df.columns:
                        valid_df[col] = ""
                
                final_df = valid_df[COLUMNS_STRUCTURE].copy()
                final_df = calculate_metrics_and_alerts(final_df)
                
                save_all_sheets(final_df, pd.DataFrame())
                st.session_state.fresh_uploaded = True
                
                st.success(f"🎉 {len(final_df)} samples loaded successfully!")
                st.rerun()
                
        except Exception as e:
            st.error(f"Error: {str(e)}")

# ====================== TABS WITH DATA ======================
st.markdown("## 🗂️ Active Samples")

tab1, tab2, tab3 = st.tabs(["📋 Overall", "🧪 Development", "🔄 Repeat"])

with tab1:
    st.subheader("Overall Active Samples")
    if len(df_running) > 0:
        st.data_editor(df_running, use_container_width=True, num_rows="dynamic", key="overall_tab")
    else:
        st.info("Koi active sample nahi hai")

with tab2:
    st.subheader("Development Section (Scratch)")
    if len(df_dev) > 0:
        st.data_editor(df_dev, use_container_width=True, num_rows="dynamic", key="dev_tab")
    else:
        st.info("Development section khali hai")

with tab3:
    st.subheader("Repeat Section")
    if len(df_repeat) > 0:
        st.data_editor(df_repeat, use_container_width=True, num_rows="dynamic", key="repeat_tab")
    else:
        st.info("Repeat section khali hai")

# Archive
st.subheader("📁 Submitted to Marketing (Archive)")
st.dataframe(df_master[df_master["Status"] == "Submitted to marketing"], use_container_width=True)

st.success("✅ Dashboard Updated! Agar data ab bhi nahi dikh raha to page refresh karo (F5 dabao).")
