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

# --- HELPER FUNCTIONS FOR EXCEL HANDLING ---
def load_all_sheets():
    """Loads Master and Trial data from Excel, initializing them if file does not exist."""
    if os.path.exists(EXCEL_FILE):
        try:
            df_master = pd.read_excel(EXCEL_FILE, sheet_name="Master_Data")
            df_trials = pd.read_excel(EXCEL_FILE, sheet_name="Trial_Data")
            
            # Standardize data to clean string format to avoid type drops
            df_master["S.no"] = df_master["S.no"].astype(str).apply(lambda x: str(x).split('.')[0].strip())
            df_master["Ref"] = df_master["Ref"].astype(str).str.strip()
            df_master["TR Code"] = df_master["TR Code"].astype(str).str.strip()
            
            # Ensure proper column headers exist for user requirements (Ppi vs Picks)
            if "Picks" not in df_master.columns and "Ppi" in df_master.columns:
                df_master.rename(columns={"Ppi": "Picks"}, inplace=True)
                
            return df_master, df_trials
        except Exception:
            pass
            
    # Standard precise structure requested by the user
    df_master = pd.DataFrame(columns=[
        "S.no", "Brand", "Ref", "Finish", "TR Code", "Source", "Pending Days in Process",
        "Request Date", "Current Date", "Delivery date", "Ready date", "Status", "Alert",
        "Warp", "Warp Slub", "Weft", "Shade", "Weave", "Reed", "Picks", "Greige Meters", "Remarks"
    ])
    df_trials = pd.DataFrame(columns=[
        "Sample ID", "Trial Number", "Parameters", "Status_OK", "Updated Date"
    ])
    
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master_Data", index=False)
        df_trials.to_excel(writer, sheet_name="Trial_Data", index=False)
        
    return df_master, df_trials

def save_all_sheets(df_master, df_trials):
    """Saves both dataframes into separate sheets in the same Excel file."""
    df_master["S.no"] = df_master["S.no"].astype(str).apply(lambda x: str(x).split('.')[0].strip())
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master_Data", index=False)
        df_trials.to_excel(writer, sheet_name="Trial_Data", index=False)

# Load current data state
df_master, df_trials = load_all_sheets()

# Helper to pull unique values for clean auto-suggestions
def get_suggestions(column_name):
    if len(df_master) > 0 and column_name in df_master.columns:
        return [str(x) for x in df_master[column_name].dropna().unique() if str(x).strip() != "" and str(x).lower() != "nan"]
    return []

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

# Categorize into Development vs Repeat based on industry source logic
# 'Scratch' is purely Development. 'Existing' and 'Production beam' fall under Repeat runs.
df_dev = df_running[df_running["Source"].str.lower() == "scratch"].copy()
df_repeat = df_running[df_running["Source"].str.lower().isin(["existing", "production beam"])].copy()

# --- TOP MAIN HEADINGS & METRIC BOXES ---
st.title("👖 Denim Fabric R&D Production Pipeline Tracker")
st.markdown("---")

# 1st Custom Container: Core Production Indicators
st.markdown("### 📊 Live Floor Workload Counters")
kpi_box1, kpi_box2, kpi_box3 = st.columns(3)

with kpi_box1:
    st.markdown(
        f"<div style='background-color:#1E3A8A; padding:20px; border-radius:10px; text-align:center; color:white;'>"
        f"<h2>Total Samples On Floor</h2>"
        f"<p style='font-size:35px; font-weight:bold; margin:0;'>{len(df_running)}</p>"
        f"</div>", 
        unsafe_with_html=True
    )

with kpi_box2:
    st.markdown(
        f"<div style='background-color:#0D9488; padding:20px; border-radius:10px; text-align:center; color:white;'>"
        f"<h2>Total Development Section</h2>"
        f"<p style='font-size:35px; font-weight:bold; margin:0;'>{len(df_dev)}</p>"
        f"</div>", 
        unsafe_with_html=True
    )

with kpi_box3:
    st.markdown(
        f"<div style='background-color:#B45309; padding:20px; border-radius:10px; text-align:center; color:white;'>"
        f"<h2>Total Repeat Section</h2>"
        f"<p style='font-size:35px; font-weight:bold; margin:0;'>{len(df_repeat)}</p>"
        f"</div>", 
        unsafe_with_html=True
    )

st.markdown("---")

# --- FILE UPLOAD BLOCK CENTER ---
with st.expander("📥 EXCEL FILE BULK DATA IMPORT CENTER", expanded=False):
    st.markdown("#### Database file upload karke pipeline sync karein")
    uploaded_file = st.file_uploader("Excel (.xlsx) sheet yahan upload karein:", type=["xlsx"])
    if uploaded_file is not None:
        if st.button("Process & Append Excel Rows", use_container_width=True):
            try:
                imported_df = pd.read_excel(uploaded_file)
                required_cols = ["S.no", "Brand", "Ref", "Status"]
                missing_cols = [c for c in required_cols if c not in imported_df.columns]
                
                if missing_cols:
                    st.error(f"Excel File mein yeh basic headers hona lazmi hain: {missing_cols}")
                else:
                    imported_df["S.no"] = imported_df["S.no"].astype(str).apply(lambda x: str(x).split('.')[0].strip())
                    imported_df["Ref"] = imported_df["Ref"].astype(str).str.strip()
                    if "Ppi" in imported_df.columns: imported_df.rename(columns={"Ppi": "Picks"}, inplace=True)
                    
                    # Row data filters
                    valid_mask = imported_df["Ref"].notna() & (imported_df["Ref"] != "")
                    filtered_imp = imported_df[valid_mask]
                    
                    new_records = filtered_imp[~filtered_imp["S.no"].astype(str).isin(df_master["S.no"])]
                    
                    if len(new_records) == 0:
                        st.warning("Excel mein koi naya unique S.no data nahi mila.")
                    else:
                        for col in df_master.columns:
                            if col not in new_records.columns: new_records[col] = ""
                        new_records = new_records[df_master.columns]
                        df_master = pd.concat([df_master, new_records], ignore_index=True)
                        df_master = calculate_metrics_and_alerts(df_master)
                        save_all_sheets(df_master, df_trials)
                        st.success(f"🎉 {len(new_records)} Naye samples successfully upload ho gaye!")
                        st.rerun()
            except Exception as e:
                st.error(f"Processing error: {str(e)}")

st.markdown("---")

# --- LIST OF SAMPLES CONTAINER BOX (3 PARTS) ---
st.markdown("## 🗂️ List of Active Technical Samples")
st.caption("🗑️ **Row Delete Karne Ka Tareeka:** Kisi bhi list ke extreme left pe box ko check/tick karein aur keyboard se **'Delete'** dabayein. Data automatic Excel se saaf ho jayega.")

tab_overall, tab_dev, tab_repeat = st.tabs([
    "📋 Part 1: Overall Run Pool (Dono Sections)", 
    "🧪 Part 2: Development Section Details", 
    "🔄 Part 3: Repeat Section Details"
])

# Part 1: Overall Segment Data
with tab_overall:
    st.markdown("#### Overall Horizontal Database Details")
    edited_overall_df = st.data_editor(
        df_running,
        use_container_width=True,
        num_rows="dynamic",
        key="overall_editor_view"
    )
    
    # Inline Deletion Synchronizer Loop
    if len(edited_overall_df) != len(df_running):
        remaining_snos = [str(x).split('.')[0].strip() for x in edited_overall_df["S.no"].dropna().tolist()]
        df_master["S.no_clean"] = df_master["S.no"].apply(lambda x: str(x).split('.')[0].strip())
        df_master = df_master[(df_master["S.no_clean"].isin(remaining_snos)) | (df_master["Status"] == "Submitted to marketing")].copy()
        df_master.drop(columns=["S.no_clean"], errors="ignore", inplace=True)
        save_all_sheets(df_master, df_trials)
        st.rerun()

# Part 2: Development Section Data
with tab_dev:
    st.markdown("#### Filtered View: Pure R&D Developments (Source: Scratch)")
    edited_dev_df = st.data_editor(
        df_dev,
        use_container_width=True,
        num_rows="dynamic",
        key="dev_editor_view"
    )
    
    if len(edited_dev_df) != len(df_dev):
        deleted_snos = set(df_dev["S.no"].tolist()) - set([str(x).split('.')[0].strip() for x in edited_dev_df["S.no"].dropna().tolist()])
        df_master = df_master[~df_master["S.no"].astype(str).apply(lambda x: x.split('.')[0].strip()).isin(deleted_snos)]
        save_all_sheets(df_master, df_trials)
        st.rerun()

# Part 3: Repeat Section Data
with tab_repeat:
    st.markdown("#### Filtered View: Repeat Runs (Source: Existing / Production Beam)")
    edited_repeat_df = st.data_editor(
        df_repeat,
        use_container_width=True,
        num_rows="dynamic",
        key="repeat_editor_view"
    )
    
    if len(edited_repeat_df) != len(df_repeat):
        deleted_snos = set(df_repeat["S.no"].tolist()) - set([str(x).split('.')[0].strip() for x in edited_repeat_df["S.no"].dropna().tolist()])
        df_master = df_master[~df_master["S.no"].astype(str).apply(lambda x: x.split('.')[0].strip()).isin(deleted_snos)]
        save_all_sheets(df_master, df_trials)
        st.rerun()

st.markdown("---")

# --- MANUAL NEW SAMPLE ENTRY INPUT FORM CONTAINER ---
st.markdown("## 📥 Box: New Sample Manual Entry Section")

with st.form("manual_sample_entry_container", clear_on_submit=True):
    st.markdown("#### Enter Quality Specifications Below to Add into Excel Directly:")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        s_no = st.text_input("S.no (Manual Primary Key)*")
        brand = st.text_input("Brand Name")
        ref_id = st.text_input("Ref / Quality No (Mandatory)*")
        finish_code = st.text_input("Finish Code")
        tr_code = st.text_input("TR Code")
        source = st.selectbox("Source (Select Scratch for Development Run):", SOURCE_OPTIONS)
        
    with col2:
        req_date_in = st.date_input("Request Date:", value=date.today())
        del_date_in = st.date_input("Delivery Date Target:")
        status_init = st.selectbox("Current Status Stage:", STATUS_OPTIONS)
        remarks = st.text_area("Remarks / Mill Instructions:")
        
    with col3:
        warp_txt = st.text_input("Warp Specs")
        w_slub_txt = st.text_input("Warp Slub")
        weft_txt = st.text_input("Weft Specs")
        shade_txt = st.text_input("Shade Type")
        weave_txt = st.text_input("Weave Pattern")
        reed_txt = st.text_input("Reed")
        picks_txt = st.text_input("Picks (PPI)")
        greige_mtrs = st.text_input("Greige Meters")
        
    submit_btn = st.form_submit_button("💾 Save Manual Record To Excel Database", use_container_width=True)
    
    if submit_btn:
        clean_s_no = str(s_no).split('.')[0].strip()
        if not clean_s_no or not ref_id.strip() or ref_id.strip().lower() == "nan":
            st.error("Fields 'S.no' aur 'Ref' Quality likhna zaroori hai. Inke bagair record block ho jayega!")
        elif clean_s_no in df_master["S.no"].values:
            st.error(f"S.no '{clean_s_no}' database mein pehle se chal raha hai! Duplicate block kar diya gaya.")
        else:
            # Build new structured record array precisely matching user layout
            new_row = {
                "S.no": clean_s_no, "Brand": brand, "Ref": ref_id.strip(), "Finish": finish_code,
                "TR Code": tr_code.strip(), "Source": source, "Pending Days in Process": 0,
                "Request Date": str(req_date_in), "Current Date": str(date.today()),
                "Delivery date": str(del_date_in), "Ready date": "", "Status": status_init,
                "Alert": "Normal", "Warp": warp_txt, "Warp Slub": w_slub_txt, "Weft": weft_txt,
                "Shade": shade_txt, "Weave": weave_txt, "Reed": reed_txt, "Picks": picks_txt,
                "Greige Meters": greige_mtrs, "Remarks": remarks
            }
            df_master = pd.concat([df_master, pd.DataFrame([new_row])], ignore_index=True)
            df_master = calculate_metrics_and_alerts(df_master)
            save_all_sheets(df_master, df_trials)
            st.success(f"🎉 Sample S.no {clean_s_no} master database mein add ho kar sync ho gaya!")
            st.rerun()

st.markdown("---")
# Archive table rendering for finalized logs at the very bottom
st.subheader("📁 Complete Closed Archive Pool (Submitted to Marketing)")
df_submitted = df_master[df_master["Status"] == "Submitted to marketing"].copy()
st.dataframe(df_submitted, use_container_width=True)
