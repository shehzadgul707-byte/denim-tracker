import streamlit as st
import pandas as pd
from datetime import date
import os

# ========================= PAGE CONFIG =========================
st.set_page_config(page_title="Denim R&D Advanced Tracker", layout="wide")

EXCEL_FILE = "Denim_Master_Database.xlsx"

# Standard Options
SOURCE_OPTIONS = ["Existing", "Scratch", "Production beam"]
STATUS_OPTIONS = ["Yarn Demand", "Dyeing", "Weaving", "Finishing", "Inspection", "Submitted to marketing"]

COLUMNS_STRUCTURE = [
    "S.no", "Brand", "Ref", "Finish", "TR Code", "Source", "Pending Days in Process",
    "Request Date", "Current Date", "Delivery date", "Ready date", "Status", "Alert",
    "Warp", "Warp Slub", "Weft", "Shade", "Weave", "Reed", "Picks", "Greige Meters", "Remarks"
]

# ========================= SESSION STATE =========================
if "fresh_uploaded" not in st.session_state:
    st.session_state.fresh_uploaded = False

# ========================= HELPER FUNCTIONS =========================
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
    
    # Critical Fix: Force S.no as string
    df["S.no"] = df["S.no"].astype(str).str.strip().str.split('.').str[0]
    
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

# ========================= LOAD DATA =========================
if os.path.exists(EXCEL_FILE) and not st.session_state.fresh_uploaded:
    try:
        df_master = pd.read_excel(EXCEL_FILE, sheet_name="Master_Data")
        df_trials = pd.read_excel(EXCEL_FILE, sheet_name="Trial_Data")
    except:
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

# ========================= UI START =========================
st.title("👖 Denim Fabric R&D Production Pipeline Tracker")
st.markdown("---")

# KPI Boxes
st.markdown("### 📊 Live Floor Workload Counters")
c1, c2, c3 = st.columns(3)

with c1:
    st.markdown(f"""<div style='background-color:#1E3A8A;padding:20px;border-radius:10px;text-align:center;color:white;'>
                    <h2>Total Samples On Floor</h2>
                    <p style='font-size:38px;font-weight:bold;margin:0;'>{len(df_running)}</p>
                  </div>""", unsafe_allow_html=True)

with c2:
    st.markdown(f"""<div style='background-color:#0D9488;padding:20px;border-radius:10px;text-align:center;color:white;'>
                    <h2>Development Section</h2>
                    <p style='font-size:38px;font-weight:bold;margin:0;'>{len(df_dev)}</p>
                  </div>""", unsafe_allow_html=True)

with c3:
    st.markdown(f"""<div style='background-color:#B45309;padding:20px;border-radius:10px;text-align:center;color:white;'>
                    <h2>Repeat Section</h2>
                    <p style='font-size:38px;font-weight:bold;margin:0;'>{len(df_repeat)}</p>
                  </div>""", unsafe_allow_html=True)

st.markdown("---")

# Control Panel
st.markdown("## ⚙️ Control Panel")
if st.button("🚨 Wipe Out Everything (Reset Dashboard)", use_container_width=True):
    df_master, df_trials = load_all_sheets_clean()
    save_all_sheets(df_master, df_trials)
    st.session_state.fresh_uploaded = False
    st.success("Saara data delete kar diya gaya!")
    st.rerun()

# ====================== EXCEL UPLOAD (Most Important Fix) ======================
st.markdown("## 📥 Fresh Excel Upload Center")
uploaded_file = st.file_uploader("Excel File (.xlsx) select karein", type=["xlsx"], key="uploader_key")

if uploaded_file is not None:
    if st.button("🚀 Process & Replace All Pipeline Data", type="primary", use_container_width=True):
        try:
            imported_df = pd.read_excel(uploaded_file)
            
            required_cols = ["S.no", "Brand", "Ref", "Status"]
            missing = [col for col in required_cols if col not in imported_df.columns]
            
            if missing:
                st.error(f"Ye columns missing hain: {missing}")
            else:
                # ====================== STRONG FIXES ======================
                imported_df = imported_df.copy()
                
                # S.no ko forcefully string banao (Yeh error ka main solution hai)
                imported_df["S.no"] = (
                    imported_df["S.no"]
                    .fillna("")
                    .astype(str)
                    .str.split('.')
                    .str[0]
                    .str.strip()
                )
                
                imported_df["Ref"] = imported_df["Ref"].astype(str).str.strip()
                
                if "Ppi" in imported_df.columns and "Picks" not in imported_df.columns:
                    imported_df.rename(columns={"Ppi": "Picks"}, inplace=True)
                
                # Valid rows only
                valid_mask = (
                    imported_df["Ref"].notna() & 
                    (imported_df["Ref"] != "") & 
                    (imported_df["Ref"].str.lower() != "nan")
                )
                final_df = imported_df[valid_mask].copy()
                
                # Add missing columns
                for col in COLUMNS_STRUCTURE:
                    if col not in final_df.columns:
                        final_df[col] = ""
                
                final_df = final_df[COLUMNS_STRUCTURE].copy()
                final_df["S.no"] = final_df["S.no"].astype(str)
                
                final_df = calculate_metrics_and_alerts(final_df)
                
                # Save
                save_all_sheets(final_df, pd.DataFrame(columns=df_trials.columns))
                st.session_state.fresh_uploaded = True
                
                st.success(f"🎉 {len(final_df)} samples successfully load ho gaye!")
                st.rerun()
                
        except Exception as e:
            st.error(f"Upload Error: {str(e)}")

st.markdown("---")

# ====================== TABS ======================
st.markdown("## 🗂️ Active Samples List")

tab1, tab2, tab3 = st.tabs([
    "📋 Overall Run Pool",
    "🧪 Development Section",
    "🔄 Repeat Section"
])

def safe_data_editor(df_filtered, key):
    edited_df = st.data_editor(
        df_filtered,
        use_container_width=True,
        num_rows="dynamic",
        key=key
    )
    return edited_df

# Tab 1
with tab1:
    if len(df_running) == 0:
        st.info("Koi active sample nahi hai. Excel upload karein.")
    else:
        edited = safe_data_editor(df_running, "overall_editor")
        if len(edited) != len(df_running):
            remaining = set(str(x).strip() for x in edited["S.no"].dropna())
            df_master = df_master[
                (df_master["S.no"].astype(str).str.strip().isin(remaining)) |
                (df_master["Status"] == "Submitted to marketing")
            ].copy()
            save_all_sheets(df_master, df_trials)
            st.rerun()

# Tab 2
with tab2:
    if len(df_dev) == 0:
        st.info("Development section khali hai.")
    else:
        edited = safe_data_editor(df_dev, "dev_editor")
        if len(edited) != len(df_dev):
            remaining = set(str(x).strip() for x in edited["S.no"].dropna())
            df_master = df_master[
                (df_master["S.no"].astype(str).str.strip().isin(remaining)) |
                (df_master["Status"] == "Submitted to marketing")
            ].copy()
            save_all_sheets(df_master, df_trials)
            st.rerun()

# Tab 3
with tab3:
    if len(df_repeat) == 0:
        st.info("Repeat section khali hai.")
    else:
        edited = safe_data_editor(df_repeat, "repeat_editor")
        if len(edited) != len(df_repeat):
            remaining = set(str(x).strip() for x in edited["S.no"].dropna())
            df_master = df_master[
                (df_master["S.no"].astype(str).str.strip().isin(remaining)) |
                (df_master["Status"] == "Submitted to marketing")
            ].copy()
            save_all_sheets(df_master, df_trials)
            st.rerun()

st.markdown("---")

# ====================== MANUAL ENTRY ======================
st.markdown("## 📥 Manual New Sample Entry")
with st.form("manual_form", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        s_no = st.text_input("S.no *")
        brand = st.text_input("Brand")
        ref = st.text_input("Ref / Quality No *")
        finish = st.text_input("Finish")
        tr_code = st.text_input("TR Code")
        source = st.selectbox("Source", SOURCE_OPTIONS)
    
    with col2:
        req_date = st.date_input("Request Date", date.today())
        del_date = st.date_input("Delivery Date")
        status = st.selectbox("Status", STATUS_OPTIONS)
        remarks = st.text_area("Remarks")
    
    with col3:
        warp = st.text_input("Warp")
        warp_slub = st.text_input("Warp Slub")
        weft = st.text_input("Weft")
        shade = st.text_input("Shade")
        weave = st.text_input("Weave")
        reed = st.text_input("Reed")
        picks = st.text_input("Picks")
        greige = st.text_input("Greige Meters")
    
    if st.form_submit_button("💾 Save Sample", use_container_width=True):
        clean_sno = str(s_no).strip().split('.')[0]
        if not clean_sno or not ref.strip():
            st.error("S.no aur Ref likhna zaroori hai!")
        elif clean_sno in df_master["S.no"].astype(str).str.strip().values:
            st.error(f"S.no {clean_sno} already exists!")
        else:
            new_row = pd.DataFrame([{
                "S.no": clean_sno, "Brand": brand, "Ref": ref.strip(), "Finish": finish,
                "TR Code": tr_code, "Source": source, "Pending Days in Process": 0,
                "Request Date": str(req_date), "Current Date": str(date.today()),
                "Delivery date": str(del_date), "Ready date": "", "Status": status,
                "Alert": "Normal", "Warp": warp, "Warp Slub": warp_slub, "Weft": weft,
                "Shade": shade, "Weave": weave, "Reed": reed, "Picks": picks,
                "Greige Meters": greige, "Remarks": remarks
            }])
            
            df_master = pd.concat([df_master, new_row], ignore_index=True)
            df_master = calculate_metrics_and_alerts(df_master)
            save_all_sheets(df_master, df_trials)
            st.success(f"Sample {clean_sno} add ho gaya!")
            st.rerun()

# Archive
st.subheader("📁 Submitted to Marketing - Archive")
st.dataframe(df_master[df_master["Status"] == "Submitted to marketing"], use_container_width=True)
