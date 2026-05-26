import streamlit as st
import pandas as pd
from datetime import date
import os
import plotly.express as px

st.set_page_config(page_title="Denim R&D Advanced Tracker", layout="wide")

EXCEL_FILE = "Denim_Master_Database.xlsx"

STATUS_OPTIONS = ["Yarn Demand", "Dyeing", "Weaving", "Finishing", "Inspection", "Submitted to marketing"]

# Standard fixed columns structure to prevent missing column errors
COLUMNS_STRUCTURE = [
    "S.no", "Brand", "Ref", "Finish", "TR Code", "Source", "Pending Days in Process",
    "Request Date", "Current Date", "Delivery date", "Ready date", "Status", "Alert",
    "Warp", "Warp Slub", "Weft", "Shade", "Weave", "Reed", "Picks", "Greige Meters", "Remarks", "Category"
]

if "fresh_uploaded" not in st.session_state:
    st.session_state.fresh_uploaded = False

# ====================== HELPERS ======================
def clean_dataframe(df):
    df = df.copy()
    # Ensure all baseline columns exist
    for col in COLUMNS_STRUCTURE:
        if col not in df.columns:
            df[col] = ""
    # Standardize spaces and text formats
    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()
    return df[COLUMNS_STRUCTURE]

def save_all_sheets(df_master):
    df_master = clean_dataframe(df_master)
    # Automatically calculate Category dynamically based on Technical Source before saving
    for idx in df_master.index:
        src = str(df_master.at[idx, 'Source']).strip().lower()
        if src == "scratch":
            df_master.at[idx, 'Category'] = "Development"
        else:
            df_master.at[idx, 'Category'] = "Repeat"
            
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master_Data", index=False)

def calculate_metrics_and_alerts(df):
    if len(df) == 0:
        return df.copy()
    df = clean_dataframe(df)
    today = date.today()
    df['Current Date'] = today.strftime("%Y-%m-%d")
    
    for idx in df.index:
        # 1. Clear floating points from serial numbers
        sno_raw = str(df.at[idx, 'S.no']).split('.')[0].strip()
        df.at[idx, 'S.no'] = sno_raw
        
        # 2. Dynamic Category Routing
        src = str(df.at[idx, 'Source']).strip().lower()
        if src == "scratch":
            df.at[idx, 'Category'] = "Development"
        else:
            df.at[idx, 'Category'] = "Repeat"

        # 3. Running Days Calculation
        try:
            req_date = pd.to_datetime(df.at[idx, 'Request Date']).date()
            df.at[idx, 'Pending Days in Process'] = str((today - req_date).days)
        except:
            df.at[idx, 'Pending Days in Process'] = "0"
        
        # 4. Critical Alert System
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

def load_data():
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name="Master_Data", dtype=str)
            return clean_dataframe(df)
        except:
            return pd.DataFrame(columns=COLUMNS_STRUCTURE)
    else:
        return pd.DataFrame(columns=COLUMNS_STRUCTURE)

df_master = load_data()

# ====================== UPLOAD CENTER ======================
st.title("👖 Denim Fabric R&D Production Pipeline Tracker")
st.markdown("---")

cc_left, cc_right = st.columns([2, 4])
with cc_left:
    upload_type = st.radio("⚙️ Upload Action Mode:", ["Full Pipeline File", "Only Repeat Sampling Sheet"], horizontal=False)

with cc_right:
    uploaded_file = st.file_uploader("Excel Sheet Browse Karein (.xlsx):", type=["xlsx"])

if uploaded_file and st.button("🚀 Upload & Process Pipeline", type="primary", use_container_width=True):
    try:
        imported_df = pd.read_excel(uploaded_file, dtype=str)
        
        if "Marketing" in imported_df.columns:
            imported_df.rename(columns={"Marketing": "Ref"}, inplace=True)
        if "Ppi" in imported_df.columns and "Picks" not in imported_df.columns:
            imported_df.rename(columns={"Ppi": "Picks"}, inplace=True)

        if "Ref" not in imported_df.columns or "S.no" not in imported_df.columns:
            st.error("Excel File mein basic headers ('S.no' aur 'Ref') hona lazmi hain!")
            st.stop()

        imported_df = clean_dataframe(imported_df)
        final_df = calculate_metrics_and_alerts(imported_df)

        if upload_type == "Full Pipeline File":
            save_all_sheets(final_df)
            st.success(f"🎉 Complete pipeline reset successfully! Loaded {len(final_df)} samples.")
        else:
            # Safe cross merge for specific updates
            snos_to_update = final_df["S.no"].astype(str).str.strip().tolist()
            df_master = df_master[~df_master["S.no"].isin(snos_to_update)]
            df_master = pd.concat([df_master, final_df], ignore_index=True)
            df_master = calculate_metrics_and_alerts(df_master)
            save_all_sheets(df_master)
            st.success(f"✅ {len(final_df)} Repeat samples seamlessly merged/updated!")
            
        st.rerun()
    except Exception as e:
        st.error(f"Processing Error: {str(e)}")

st.markdown("---")

# ====================== DYNAMIC PIPELINE FILTERS ======================
df_master = calculate_metrics_and_alerts(df_master)

# Separate active running samples from complete marketing archive
df_running = df_master[df_master["Status"] != "Submitted to marketing"].copy()
df_dev = df_running[df_running["Category"].str.lower() == "development"].copy()
df_repeat = df_running[df_running["Category"].str.lower() == "repeat"].copy()

# ====================== DASHBOARD KPI & PLOTLY CHARTS ======================
st.markdown("### 📊 Live R&D Floor Workload Indicators")
kpi1, kpi2, kpi3 = st.columns([1, 2, 2])

with kpi1:
    st.metric(label="Total Active On Floor", value=str(len(df_running)))
    st.metric(label="🧪 Development Pool", value=str(len(df_dev)))
    st.metric(label="🔄 Repeat Pool", value=str(len(df_repeat)))

with kpi2:
    if len(df_running) > 0:
        fig_pie = px.pie(
            names=['Development Section', 'Repeat Run Section'], 
            values=[len(df_dev), len(df_repeat)], 
            title="Workload Share (%)", 
            color_discrete_sequence=['#0D9488', '#B45309'],
            hole=0.4
        )
        st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.info("No data available for charts.")

with kpi3:
    if len(df_running) > 0:
        fig_bar = px.bar(
            x=['Development Pool', 'Repeat Pool'], 
            y=[len(df_dev), len(df_repeat)], 
            text=[len(df_dev), len(df_repeat)], 
            title="Exact Roll Count On Floor",
            color=['Development Pool', 'Repeat Pool'],
            color_discrete_map={'Development Pool': '#0D9488', 'Repeat Pool': '#B45309'}
        )
        st.plotly_chart(fig_bar, use_container_width=True)

st.markdown("---")

# ====================== SAFE EDIT FORM SECTION ======================
st.subheader("✏️ Quick Search & Technical Edit Form")

if len(df_running) > 0:
    df_running["Display"] = df_running["S.no"].astype(str) + " - [Ref: " + df_running["Ref"].astype(str) + "] - " + df_running["Brand"].astype(str)
    selected_display = st.selectbox("Select Sample Target to Overwrite:", options=df_running["Display"].tolist())
    
    if selected_display:
        selected_sno = selected_display.split(" - ")[0].strip()
        sample_row = df_master[df_master["S.no"].astype(str) == selected_sno]
        
        if not sample_row.empty:
            sample = sample_row.iloc[0]
            
            with st.form("edit_sample_form_secure"):
                st.markdown(f"⚙️ **Editing Specifications For Sample S.no: `{selected_sno}`**")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    brand = st.text_input("Brand", value=sample.get("Brand", ""))
                    ref = st.text_input("Ref Quality No", value=sample.get("Ref", ""))
                    finish = st.text_input("Finish Type", value=sample.get("Finish", ""))
                    tr_code = st.text_input("TR Code", value=sample.get("TR Code", ""))
                    source = st.selectbox("Source", ["Scratch", "Existing", "Production beam"], 
                                          index=["scratch", "existing", "production beam"].index(str(sample.get("Source", "Existing")).lower()) if str(sample.get("Source")).lower() in ["scratch", "existing", "production beam"] else 1)
                
                with col2:
                    status = st.selectbox("Pipeline Status Stage", STATUS_OPTIONS, 
                                          index=STATUS_OPTIONS.index(sample.get("Status")) if sample.get("Status") in STATUS_OPTIONS else 0)
                    warp_txt = st.text_input("Warp Specs", value=sample.get("Warp", ""))
                    w_slub_txt = st.text_input("Warp Slub", value=sample.get("Warp Slub", ""))
                    weft_txt = st.text_input("Weft Specs", value=sample.get("Weft", ""))
                    shade_txt = st.text_input("Shade Type", value=sample.get("Shade", ""))
                
                with col3:
                    try: req_default = pd.to_datetime(sample.get("Request Date")).date()
                    except: req_default = date.today()
                    try: del_default = pd.to_datetime(sample.get("Delivery date")).date()
                    except: del_default = date.today()
                    
                    req_date = st.date_input("Request Date", value=req_default)
                    del_date = st.date_input("Delivery Date Target", value=del_default)
                    picks_txt = st.text_input("Picks (PPI)", value=sample.get("Picks", ""))
                    remarks = st.text_area("Remarks / Factory Instructions", value=sample.get("Remarks", ""))
                
                if st.form_submit_button("💾 Save Data Updates To Database File", use_container_width=True):
                    mask = df_master["S.no"].astype(str) == selected_sno
                    df_master.loc[mask, "Brand"] = brand
                    df_master.loc[mask, "Ref"] = ref
                    df_master.loc[mask, "Finish"] = finish
                    df_master.loc[mask, "TR Code"] = tr_code
                    df_master.loc[mask, "Source"] = source
                    df_master.loc[mask, "Status"] = status
                    df_master.loc[mask, "Warp"] = warp_txt
                    df_master.loc[mask, "Warp Slub"] = w_slub_txt
                    df_master.loc[mask, "Weft"] = weft_txt
                    df_master.loc[mask, "Shade"] = shade_txt
                    df_master.loc[mask, "Picks"] = picks_txt
                    df_master.loc[mask, "Request Date"] = str(req_date)
                    df_master.loc[mask, "Delivery date"] = str(del_date)
                    df_master.loc[mask, "Remarks"] = remarks
                    
                    df_master = calculate_metrics_and_alerts(df_master)
                    save_all_sheets(df_master)
                    st.success("🎉 Sample changes applied successfully!")
                    st.rerun()
else:
    st.info("Floor par koi active running target sample nahi mila.")

st.markdown("---")

# ====================== DATA EDITORS & TABS ======================
view = st.radio("🗂️ Select View Pipeline Section:", ["Overall Active View", "Development Section (Scratch)", "Repeat Run Section"], horizontal=True)

if view == "Overall Active View":
    st.subheader("📋 All Active R&D Samples Pool")
    if len(df_running) > 0:
        edited_overall = st.data_editor(df_running, use_container_width=True, num_rows="dynamic", key="overall_editor_sync")
        if len(edited_overall) != len(df_running):
            active_snos = edited_overall["S.no"].tolist()
            # Retain archived but filter out deleted live tracks
            df_master = df_master[(df_master["S.no"].isin(active_snos)) | (df_master["Status"] == "Submitted to marketing")].copy()
            save_all_sheets(df_master)
            st.rerun()
    else:
        st.info("Pipeline Table Empty. Upload an Excel File above to begin tracking.")

elif view == "Development Section (Scratch)":
    st.subheader("🧪 Pure Development Specs Filter")
    if len(df_dev) > 0:
        edited_dev = st.data_editor(df_dev, use_container_width=True, num_rows="dynamic", key="dev_editor_sync")
        if len(edited_dev) != len(df_dev):
            active_snos = edited_dev["S.no"].tolist()
            df_master = df_master[(df_master["S.no"].isin(active_snos)) | (df_master["Category"].str.lower() != "development") | (df_master["Status"] == "Submitted to marketing")].copy()
            save_all_sheets(df_master)
            st.rerun()
    else:
        st.info("Development pool khali hai (No Scratch Source matches found).")

elif view == "Repeat Run Section":
    st.subheader("🔄 Repeat Processing Run Specs Filter")
    if len(df_repeat) > 0:
        edited_repeat = st.data_editor(df_repeat, use_container_width=True, num_rows="dynamic", key="repeat_editor_sync")
        if len(edited_repeat) != len(df_repeat):
            active_snos = edited_repeat["S.no"].tolist()
            df_master = df_master[(df_master["S.no"].isin(active_snos)) | (df_master["Category"].str.lower() != "repeat") | (df_master["Status"] == "Submitted to marketing")].copy()
            save_all_sheets(df_master)
            st.rerun()
    else:
        st.info("Repeat pool khali hai.")

# Permanent Floor Archive View
st.markdown("---")
st.subheader("📁 Complete Closed Archive Pool (Submitted to Marketing)")
df_submitted = df_master[df_master["Status"] == "Submitted to marketing"].copy()
st.dataframe(df_submitted, use_container_width=True)
