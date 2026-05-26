import streamlit as st
import pandas as pd
from datetime import date
import os
import plotly.express as px

st.set_page_config(page_title="Denim R&D Advanced Tracker", layout="wide")

EXCEL_FILE = "Denim_Master_Database.xlsx"

STATUS_OPTIONS = ["Yarn Demand", "Dyeing", "Weaving", "Finishing", "Inspection", "Submitted to marketing"]

# Standard fixed columns structure precisely matching user specs
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
    for col in COLUMNS_STRUCTURE:
        if col not in df.columns:
            df[col] = ""
    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()
    return df[COLUMNS_STRUCTURE]

def save_all_sheets(df_master):
    df_master = clean_dataframe(df_master)
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master_Data", index=False)

def calculate_metrics_and_alerts(df):
    if len(df) == 0:
        return df.copy()
    df = clean_dataframe(df)
    today = date.today()
    df['Current Date'] = today.strftime("%Y-%m-%d")
    
    for idx in df.index:
        sno_raw = str(df.at[idx, 'S.no']).split('.')[0].strip()
        df.at[idx, 'S.no'] = sno_raw

        # Running Days Calculation
        try:
            req_val = df.at[idx, 'Request Date']
            if req_val and str(req_val).lower() not in ["nat", "nan", ""]:
                req_date = pd.to_datetime(req_val).date()
                df.at[idx, 'Pending Days in Process'] = str((today - req_date).days)
            else:
                df.at[idx, 'Pending Days in Process'] = "0"
        except:
            df.at[idx, 'Pending Days in Process'] = "0"
        
        # Critical Alert System
        status = str(df.at[idx, 'Status']).strip()
        if status != "Submitted to marketing":
            try:
                del_val = df.at[idx, 'Delivery date']
                if del_val and str(del_val).lower() not in ["nat", "nan", ""]:
                    del_date = pd.to_datetime(del_val).date()
                    days_left = (del_date - today).days
                    if days_left <= 3 and days_left >= 0:
                        df.at[idx, 'Alert'] = "⚠️ Critical"
                    elif days_left < 0:
                        df.at[idx, 'Alert'] = "🚨 Overdue"
                    else:
                        df.at[idx, 'Alert'] = "Normal"
                else:
                    df.at[idx, 'Alert'] = "No Delivery Date"
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
df_running = df_master[df_master["Status"] != "Submitted to marketing"].copy()

df_dev = df_running[df_running["Category"].str.lower().str.contains("development", na=False)].copy()
df_repeat = df_running[df_running["Category"].str.lower().str.contains("repeat", na=False)].copy()

# ====================== DASHBOARD KPI & CHARTS ======================
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

# ====================== THE PERFECT 23-PARAMETER EDIT FORM ======================
st.subheader("✏️ Search & Complete Technical Specification Editor")

if len(df_running) > 0:
    df_running["Display"] = df_running["S.no"].astype(str) + " - [Ref: " + df_running["Ref"].astype(str) + "] - " + df_running["Brand"].astype(str)
    selected_display = st.selectbox("Select Sample Target to Edit ALL Fields:", options=df_running["Display"].tolist())
    
    if selected_display:
        selected_sno = selected_display.split(" - ")[0].strip()
        sample_row = df_master[df_master["S.no"].astype(str) == selected_sno]
        
        if not sample_row.empty:
            sample = sample_row.iloc[0]
            
            with st.form("edit_sample_form_secure_all_23_fields"):
                st.markdown(f"⚙️ **Editing Specifications For Sample S.no: `{selected_sno}`**")
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.markdown("##### 📁 Section 1: Core Identity")
                    new_sno = st.text_input("S.no (Manual Serial)", value=sample.get("S.no", ""))
                    brand = st.text_input("Brand / Customer Name", value=sample.get("Brand", ""))
                    ref = st.text_input("Ref / Quality ID", value=sample.get("Ref", ""))
                    finish = st.text_input("Finish Type", value=sample.get("Finish", ""))
                    tr_code = st.text_input("TR Code", value=sample.get("TR Code", ""))
                    source = st.text_input("Source / Fabric Route", value=sample.get("Source", ""))

                with col2:
                    st.markdown("##### 📅 Section 2: Management & Dates")
                    status = st.selectbox("Pipeline Status Stage", STATUS_OPTIONS, 
                                          index=STATUS_OPTIONS.index(sample.get("Status")) if sample.get("Status") in STATUS_OPTIONS else 0)
                    
                    cat_val = str(sample.get("Category", "Development")).lower().strip()
                    cat_idx = 0 if "development" in cat_val else 1
                    category = st.selectbox("Category Group", ["Development", "Repeat"], index=cat_idx)
                    
                    # Date values parsing safely
                    req_val = sample.get("Request Date")
                    del_val = sample.get("Delivery date")
                    
                    try: req_default = pd.to_datetime(req_val).date() if req_val and str(req_val).lower() not in ["nat", "nan", ""] else date.today()
                    except: req_default = date.today()
                    try: del_default = pd.to_datetime(del_val).date() if del_val and str(del_val).lower() not in ["nat", "nan", ""] else date.today()
                    except: del_default = date.today()
                    
                    req_date = st.date_input("Request Date (Entry)", value=req_default)
                    del_date = st.date_input("Delivery Date Target", value=del_default)
                    
                    ready_date = st.text_input("Ready Date (Mill Production Out)", value=sample.get("Ready date", ""))
                    current_date_str = st.text_input("Current Tracking Date", value=sample.get("Current Date", ""))

                with col3:
                    st.markdown("##### 🧵 Section 3: Yarn Technical Specs")
                    warp_txt = st.text_input("Warp Specs (Count)", value=sample.get("Warp", ""))
                    w_slub_txt = st.text_input("Warp Slub Info", value=sample.get("Warp Slub", ""))
                    weft_txt = st.text_input("Weft Specs (Count)", value=sample.get("Weft", ""))
                    shade_txt = st.text_input("Shade Type / Indigo Tone", value=sample.get("Shade", ""))
                    pending_days = st.text_input("Pending Days in Process", value=sample.get("Pending Days in Process", "0"))
                    alert_state = st.text_input("Alert Status Notification", value=sample.get("Alert", "Normal"))

                with col4:
                    st.markdown("##### 📐 Section 4: Fabric Structure")
                    weave_txt = st.text_input("Weave Pattern (e.g. 3/1 RHT)", value=sample.get("Weave", ""))
                    reed_txt = st.text_input("Reed", value=sample.get("Reed", ""))
                    picks_txt = st.text_input("Picks (PPI)", value=sample.get("Picks", ""))
                    greige_mtrs = st.text_input("Greige Meters", value=sample.get("Greige Meters", ""))
                
                st.markdown("---")
                remarks = st.text_area("Remarks / Mill Production Instructions:", value=sample.get("Remarks", ""))
                
                if st.form_submit_button("💾 Save All 23 Parameter Updates To Excel Database", use_container_width=True):
                    # Delete old instance tracking row to avoid duplications
                    df_master = df_master[df_master["S.no"].astype(str) != selected_sno].copy()
                    
                    # Generate fresh dictionary holding exact 23 fields structure mapping
                    updated_row = {
                        "S.no": str(new_sno).split('.')[0].strip(),
                        "Brand": str(brand).strip(),
                        "Ref": str(ref).strip(),
                        "Finish": str(finish).strip(),
                        "TR Code": str(tr_code).strip(),
                        "Source": str(source).strip(),
                        "Pending Days in Process": str(pending_days).strip(),
                        "Request Date": str(req_date),
                        "Current Date": str(current_date_str).strip(),
                        "Delivery date": str(del_date),
                        "Ready date": str(ready_date).strip(),
                        "Status": str(status).strip(),
                        "Alert": str(alert_state).strip(),
                        "Warp": str(warp_txt).strip(),
                        "Warp Slub": str(w_slub_txt).strip(),
                        "Weft": str(weft_txt).strip(),
                        "Shade": str(shade_txt).strip(),
                        "Weave": str(weave_txt).strip(),
                        "Reed": str(reed_txt).strip(),
                        "Picks": str(picks_txt).strip(),
                        "Greige Meters": str(greige_mtrs).strip(),
                        "Remarks": str(remarks).strip(),
                        "Category": str(category).strip()
                    }
                    
                    df_master = pd.concat([df_master, pd.DataFrame([updated_row])], ignore_index=True)
                    df_master = calculate_metrics_and_alerts(df_master)
                    save_all_sheets(df_master)
                    st.success(f"🎉 All 23 specifications for Sample S.no {new_sno} updated successfully!")
                    st.rerun()
else:
    st.info("Floor par koi active running target sample nahi mila.")

st.markdown("---")

# ====================== DATA EDITORS & TABS ======================
view = st.radio("🗂 *Select View Pipeline Section:*", ["Overall Active View", "Development Section (Category)", "Repeat Run Section (Category)"], horizontal=True)

if view == "Overall Active View":
    st.subheader("📋 All Active R&D Samples Pool")
    if len(df_running) > 0:
        edited_overall = st.data_editor(df_running, use_container_width=True, num_rows="dynamic", key="overall_editor_sync")
        if len(edited_overall) != len(df_running):
            active_snos = edited_overall["S.no"].tolist()
            df_master = df_master[(df_master["S.no"].isin(active_snos)) | (df_master["Status"] == "Submitted to marketing")].copy()
            save_all_sheets(df_master)
            st.rerun()
    else:
        st.info("Pipeline Table Empty. Upload an Excel File above to begin tracking.")

elif view == "Development Section (Category)":
    st.subheader("🧪 Pure Development Specs Filter (Based strictly on Category)")
    if len(df_dev) > 0:
        edited_dev = st.data_editor(df_dev, use_container_width=True, num_rows="dynamic", key="dev_editor_sync")
        if len(edited_dev) != len(df_dev):
            active_snos = edited_dev["S.no"].tolist()
            df_master = df_master[(df_master["S.no"].isin(active_snos)) | (~df_master["Category"].str.lower().str.contains("development", na=False)) | (df_master["Status"] == "Submitted to marketing")].copy()
            save_all_sheets(df_master)
            st.rerun()
    else:
        st.info("Development pool khali hai (No samples found with Category containing 'Development').")

elif view == "Repeat Run Section (Category)":
    st.subheader("🔄 Repeat Processing Run Specs Filter (Based strictly on Category)")
    if len(df_repeat) > 0:
        edited_repeat = st.data_editor(df_repeat, use_container_width=True, num_rows="dynamic", key="repeat_editor_sync")
        if len(edited_repeat) != len(df_repeat):
            active_snos = edited_repeat["S.no"].tolist()
            df_master = df_master[(df_master["S.no"].isin(active_snos)) | (~df_master["Category"].str.lower().str.contains("repeat", na=False)) | (df_master["Status"] == "Submitted to marketing")].copy()
            save_all_sheets(df_master)
            st.rerun()
    else:
        st.info("Repeat pool khali hai (No samples found with Category containing 'Repeat').")

# Permanent Floor Archive View
st.markdown("---")
st.subheader("📁 Complete Closed Archive Pool (Submitted to Marketing)")
df_submitted = df_master[df_master["Status"] == "Submitted to marketing"].copy()
st.dataframe(df_submitted, use_container_width=True)
