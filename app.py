import streamlit as st
import pandas as pd
from datetime import date
import os

# Page Config
st.set_page_config(page_title="Denim R&D Advanced Tracker", layout="wide")

EXCEL_FILE = "Denim_Master_Database.xlsx"

# Standard Lists
SOURCE_OPTIONS = ["Existing", "Scratch", "Production beam"]
STATUS_OPTIONS = ["Yarn Demand", "Dyeing", "Weaving", "Finishing", "Inspection", "Submitted to marketing"]

COLUMNS_STRUCTURE = [
    "S.no", "Brand", "Ref", "Finish", "TR Code", "Source", "Pending Days in Process",
    "Request Date", "Current Date", "Delivery date", "Ready date", "Status", "Alert",
    "Warp", "Warp Slub", "Weft", "Shade", "Weave", "Reed", "Picks", "Greige Meters", "Remarks"
]

# ====================== HELPER FUNCTIONS ======================

def init_session_state():
    if "fresh_uploaded" not in st.session_state:
        st.session_state.fresh_uploaded = False

def load_all_sheets_clean():
    """Return empty dataframes with correct structure"""
    df_master = pd.DataFrame(columns=COLUMNS_STRUCTURE)
    df_trials = pd.DataFrame(columns=["Sample ID", "Trial Number", "Parameters", "Status_OK", "Updated Date"])
    return df_master, df_trials

def save_all_sheets(df_master, df_trials):
    """Save to Excel"""
    if len(df_master) > 0:
        df_master = df_master.copy()
        df_master["S.no"] = df_master["S.no"].astype(str).str.split('.').str[0].str.strip()
    
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master_Data", index=False)
        df_trials.to_excel(writer, sheet_name="Trial_Data", index=False)

def calculate_metrics_and_alerts(df):
    if len(df) == 0:
        return df.copy()
    
    df = df.copy()
    today = date.today()
    df['Current Date'] = today.strftime("%Y-%m-%d")
    
    for idx, row in df.iterrows():
        # Pending Days
        try:
            req_date = pd.to_datetime(row['Request Date']).date()
            df.at[idx, 'Pending Days in Process'] = (today - req_date).days
        except:
            df.at[idx, 'Pending Days in Process'] = 0
        
        # Alert Logic
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

# ====================== INITIALIZATION ======================

init_session_state()

# Load data
if os.path.exists(EXCEL_FILE) and not st.session_state.fresh_uploaded:
    try:
        df_master = pd.read_excel(EXCEL_FILE, sheet_name="Master_Data")
        df_trials = pd.read_excel(EXCEL_FILE, sheet_name="Trial_Data", nrows=0)  # just structure
    except Exception:
        df_master, df_trials = load_all_sheets_clean()
else:
    df_master, df_trials = load_all_sheets_clean()

# Ensure all columns exist
for col in COLUMNS_STRUCTURE:
    if col not in df_master.columns:
        df_master[col] = ""

df_master = df_master[COLUMNS_STRUCTURE].copy()
df_master = calculate_metrics_and_alerts(df_master)

# Filters
df_running = df_master[df_master["Status"] != "Submitted to marketing"].copy()
df_dev = df_running[df_running["Source"].str.lower() == "scratch"].copy()
df_repeat = df_running[df_running["Source"].str.lower().isin(["existing", "production beam"])].copy()

# ====================== UI ======================

st.title("👖 Denim Fabric R&D Production Pipeline Tracker")
st.markdown("---")

# KPI Boxes
st.markdown("### 📊 Live Floor Workload Counters")
kpi1, kpi2, kpi3 = st.columns(3)

with kpi1:
    st.markdown(f"""
    <div style='background-color:#1E3A8A; padding:20px; border-radius:10px; text-align:center; color:white;'>
        <h2>Total Samples On Floor</h2>
        <p style='font-size:35px; font-weight:bold; margin:0;'>{len(df_running)}</p>
    </div>
    """, unsafe_allow_html=True)

with kpi2:
    st.markdown(f"""
    <div style='background-color:#0D9488; padding:20px; border-radius:10px; text-align:center; color:white;'>
        <h2>Development Section</h2>
        <p style='font-size:35px; font-weight:bold; margin:0;'>{len(df_dev)}</p>
    </div>
    """, unsafe_allow_html=True)

with kpi3:
    st.markdown(f"""
    <div style='background-color:#B45309; padding:20px; border-radius:10px; text-align:center; color:white;'>
        <h2>Repeat Section</h2>
        <p style='font-size:35px; font-weight:bold; margin:0;'>{len(df_repeat)}</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# Control Panel
st.markdown("## ⚙️ Control Panel")
col_c1, col_c2 = st.columns([1, 4])
with col_c1:
    if st.button("🚨 Wipe Out Everything", use_container_width=True):
        df_master, df_trials = load_all_sheets_clean()
        save_all_sheets(df_master, df_trials)
        st.session_state.fresh_uploaded = False
        st.success("Saara data clear kar diya gaya hai!")
        st.rerun()

# Fresh Upload
st.markdown("## 📥 Fresh Excel Upload (Full Replace)")
uploaded_file = st.file_uploader("Excel file select karein (.xlsx)", type=["xlsx"], key="uploader")

if uploaded_file is not None:
    if st.button("🚀 Process & Replace All Data", use_container_width=True, type="primary"):
        try:
            imported_df = pd.read_excel(uploaded_file)
            
            # Basic validation
            required = ["S.no", "Brand", "Ref", "Status"]
            missing = [col for col in required if col not in imported_df.columns]
            if missing:
                st.error(f"Missing columns: {missing}")
            else:
                # Clean data
                imported_df["S.no"] = imported_df["S.no"].astype(str).str.split('.').str[0].str.strip()
                imported_df["Ref"] = imported_df["Ref"].astype(str).str.strip()
                
                if "Ppi" in imported_df.columns and "Picks" not in imported_df.columns:
                    imported_df.rename(columns={"Ppi": "Picks"}, inplace=True)
                
                # Filter valid rows
                valid_df = imported_df[imported_df["Ref"].notna() & (imported_df["Ref"] != "")].copy()
                
                # Add missing columns
                for col in COLUMNS_STRUCTURE:
                    if col not in valid_df.columns:
                        valid_df[col] = ""
                
                final_df = valid_df[COLUMNS_STRUCTURE].copy()
                final_df = calculate_metrics_and_alerts(final_df)
                
                # Save
                save_all_sheets(final_df, pd.DataFrame(columns=df_trials.columns))
                st.session_state.fresh_uploaded = True
                
                st.success(f"✅ {len(final_df)} samples successfully loaded!")
                st.rerun()
                
        except Exception as e:
            st.error(f"Error: {str(e)}")

st.markdown("---")

# ====================== TABS ======================
st.markdown("## 🗂️ Active Technical Samples")

tab1, tab2, tab3 = st.tabs([
    "📋 Overall Run Pool",
    "🧪 Development Section (Scratch)",
    "🔄 Repeat Section (Existing/Production Beam)"
])

# Common function for editing
def handle_data_editor(df_filtered, key, df_master_original):
    edited = st.data_editor(
        df_filtered,
        use_container_width=True,
        num_rows="dynamic",
        key=key
    )
    
    if len(edited) != len(df_filtered):
        # Deletion happened
        remaining_snos = set(str(x).split('.')[0].strip() for x in edited["S.no"].dropna())
        df_master_original = df_master_original[
            (df_master_original["S.no"].astype(str).str.split('.').str[0].str.strip().isin(remaining_snos)) |
            (df_master_original["Status"] == "Submitted to marketing")
        ].copy()
        return df_master_original, True
    return df_master_original, False

# Tab 1: Overall
with tab1:
    st.markdown("#### Overall Active Samples")
    if len(df_running) == 0:
        st.info("Koi active sample nahi hai.")
    else:
        df_master, rerun_needed = handle_data_editor(df_running, "overall_editor", df_master)
        if rerun_needed:
            save_all_sheets(df_master, df_trials)
            st.rerun()

# Tab 2: Development
with tab2:
    st.markdown("#### Development Section")
    if len(df_dev) == 0:
        st.info("Development section mein koi data nahi.")
    else:
        df_master, rerun_needed = handle_data_editor(df_dev, "dev_editor", df_master)
        if rerun_needed:
            save_all_sheets(df_master, df_trials)
            st.rerun()

# Tab 3: Repeat
with tab3:
    st.markdown("#### Repeat Section")
    if len(df_repeat) == 0:
        st.info("Repeat section mein koi data nahi.")
    else:
        df_master, rerun_needed = handle_data_editor(df_repeat, "repeat_editor", df_master)
        if rerun_needed:
            save_all_sheets(df_master, df_trials)
            st.rerun()

st.markdown("---")

# ====================== MANUAL ENTRY ======================
st.markdown("## 📥 New Sample Manual Entry")

with st.form("manual_entry_form", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        s_no = st.text_input("S.no *", placeholder="101")
        brand = st.text_input("Brand Name")
        ref_id = st.text_input("Ref / Quality No *")
        finish = st.text_input("Finish")
        tr_code = st.text_input("TR Code")
        source = st.selectbox("Source", SOURCE_OPTIONS)
    
    with col2:
        req_date = st.date_input("Request Date", value=date.today())
        del_date = st.date_input("Delivery Date Target")
        status = st.selectbox("Status", STATUS_OPTIONS)
        remarks = st.text_area("Remarks")
    
    with col3:
        warp = st.text_input("Warp")
        warp_slub = st.text_input("Warp Slub")
        weft = st.text_input("Weft")
        shade = st.text_input("Shade")
        weave = st.text_input("Weave")
        reed = st.text_input("Reed")
        picks = st.text_input("Picks (PPI)")
        greige_mtrs = st.text_input("Greige Meters")
    
    submitted = st.form_submit_button("💾 Save to Database", use_container_width=True)
    
    if submitted:
        clean_sno = str(s_no).split('.')[0].strip()
        if not clean_sno or not ref_id.strip():
            st.error("S.no aur Ref dono likhna zaroori hai!")
        elif clean_sno in df_master["S.no"].astype(str).str.split('.').str[0].str.strip().values:
            st.error(f"S.no {clean_sno} already exist!")
        else:
            new_row = pd.DataFrame([{
                "S.no": clean_sno, "Brand": brand, "Ref": ref_id.strip(), "Finish": finish,
                "TR Code": tr_code, "Source": source, "Pending Days in Process": 0,
                "Request Date": str(req_date), "Current Date": str(date.today()),
                "Delivery date": str(del_date), "Ready date": "", "Status": status,
                "Alert": "Normal", "Warp": warp, "Warp Slub": warp_slub, "Weft": weft,
                "Shade": shade, "Weave": weave, "Reed": reed, "Picks": picks,
                "Greige Meters": greige_mtrs, "Remarks": remarks
            }])
            
            df_master = pd.concat([df_master, new_row], ignore_index=True)
            df_master = calculate_metrics_and_alerts(df_master)
            save_all_sheets(df_master, df_trials)
            st.success(f"Sample S.no {clean_sno} add ho gaya!")
            st.rerun()

# Archive
st.subheader("📁 Submitted to Marketing (Archive)")
df_submitted = df_master[df_master["Status"] == "Submitted to marketing"].copy()
st.dataframe(df_submitted, use_container_width=True)
