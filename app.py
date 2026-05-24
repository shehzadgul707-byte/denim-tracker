import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, date
import os

# Set layout to wide for dashboard
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
            df_master["S.no"] = df_master["S.no"].astype(str)
            return df_master, df_trials
        except Exception:
            pass
            
    df_master = pd.DataFrame(columns=[
        "S.no", "Brand", "Ref", "Finish", "TR Code", "Source", "Pending days in process",
        "Request Date", "Current Date", "Delivery date", "Ready date", "Status", "Alert",
        "Warp", "Warp Slub", "Weft", "Shade", "Weave", "Reed", "Ppi", "Greige mtrs",
        "Marketing requested meter", "Remarks"
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
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master_Data", index=False)
        df_trials.to_excel(writer, sheet_name="Trial_Data", index=False)

# Load current data state
df_master, df_trials = load_all_sheets()

# Helper to pull unique values for clean auto-suggestions
def get_suggestions(column_name):
    if len(df_master) > 0 and column_name in df_master.columns:
        return [str(x) for x in df_master[column_name].dropna().unique() if str(x).strip() != ""]
    return []

# --- AUTOMATIC FORMULA CALCULATIONS ---
def calculate_metrics_and_alerts(df):
    if len(df) == 0:
        return df
    today = date.today()
    df['Current Date'] = today.strftime("%Y-%m-%d")
    df["S.no"] = df["S.no"].astype(str)
    
    for idx, row in df.iterrows():
        try:
            req_date = pd.to_datetime(row['Request Date']).date()
            df.at[idx, 'Pending days in process'] = (today - req_date).days
        except:
            df.at[idx, 'Pending days in process'] = 0
            
        if row['Status'] != "Submitted to marketing":
            try:
                del_date = pd.to_datetime(row['Delivery date']).date()
                days_left = (del_date - today).days
                if days_left <= 3 and days_left >= 0:
                    df.at[idx, 'Alert'] = "⚠️ Critical (3 Days or Less)"
                elif days_left < 0:
                    df.at[idx, 'Alert'] = "🚨 Delayed/Overdue"
                else:
                    df.at[idx, 'Alert'] = "Normal"
            except:
                df.at[idx, 'Alert'] = "No Delivery Date"
        else:
            df.at[idx, 'Alert'] = "Completed"
            
    return df

df_master = calculate_metrics_and_alerts(df_master)

# --- NAVIGATION CONTROLS ---
st.title("👖 Denim R&D Sample Tracker & Quality Dashboard")
st.divider()

# --- SIDEBAR: IMPORT DATA (EXCEL & IMAGE) ---
st.sidebar.header("⚙️ Data Import Center")

# 1. Bulk Excel Import
st.sidebar.subheader("📥 Excel File Se Data Utayein")
uploaded_file = st.sidebar.file_uploader("Excel (.xlsx) file browse karein:", type=["xlsx"])
if uploaded_file is not None:
    if st.sidebar.button("Process & Import Excel"):
        try:
            imported_df = pd.read_excel(uploaded_file)
            required_cols = ["S.no", "Brand", "Ref", "Status"]
            missing_cols = [c for c in required_cols if c not in imported_df.columns]
            
            if missing_cols:
                st.sidebar.error(f"Excel mein yeh columns hona zaroori hain: {missing_cols}")
            else:
                imported_df["S.no"] = imported_df["S.no"].astype(str)
                new_records = imported_df[~imported_df["S.no"].isin(df_master["S.no"])]
                if len(new_records) == 0:
                    st.sidebar.warning("Sheet mein koi naya unique S.no nahi mila.")
                else:
                    for col in df_master.columns:
                        if col not in new_records.columns: new_records[col] = ""
                    new_records = new_records[df_master.columns]
                    df_master = pd.concat([df_master, new_records], ignore_index=True)
                    df_master = calculate_metrics_and_alerts(df_master)
                    save_all_sheets(df_master, df_trials)
                    st.sidebar.success(f"{len(new_records)} Naye samples import ho gaye!")
                    st.rerun()
        except Exception as e:
            st.sidebar.error(f"Import failed: {str(e)}")

# 2. Picture / OCR Data Extraction
st.sidebar.subheader("📷 Picture Se Data Utayein (OCR)")
uploaded_img = st.sidebar.file_uploader("Sample Card ya Specs ki photo upload karein:", type=["png", "jpg", "jpeg"])
if uploaded_img is not None:
    st.sidebar.info("Scanning image text... (Simulated Engine)")
    # Fallback clean string parsing simulation for text processing without external heavy libraries
    mock_scanned_text = f"Scanned from {uploaded_img.name}:\nBrand: Natasha\nRef: REF-772\nWarp: 10 OE\nWeft: 12 Ring\nReed: 72\nPpi: 52\nShade: Deep Indigo\nWeave: 3/1 RHT"
    st.sidebar.text_area("Scanned Data Raw Result:", value=mock_scanned_text, height=150)
    st.sidebar.caption("💡 Aap is text se copy karke neeche form mein direct daal sakte hain!")

st.sidebar.divider()
menu = st.sidebar.radio("Navigation Menu:", ["Main Dashboard Visuals", "Register New R&D Sample", "Search & Internal Detail Viewer"])

# --- OPTION 1: REGISTER NEW SAMPLE FORM (CLEAN SINGLE BOX ENTRIES) ---
if menu == "Register New R&D Sample":
    st.header("📋 New Sample Quality Specification Input Form")
    
    with st.form("sample_entry_form"):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            s_no = st.text_input("S.no (Manual Entry)*")
            brand = st.text_input("Brand Name", help="Suggestions: " + ", ".join(get_suggestions("Brand")[:5]))
            ref_id = st.text_input("Ref / Quality No*")
            finish_code = st.text_input("Finish Code")
            tr_code = st.text_input("TR Code")
            source = st.selectbox("Source Route Selection:", SOURCE_OPTIONS)
            
        with col2:
            req_date_in = st.date_input("Request Date:", value=date.today())
            del_date_in = st.date_input("Delivery Date:")
            status_init = st.selectbox("Current Status Stage:", STATUS_OPTIONS)
            remarks = st.text_area("Remarks / Instructions:")
            
        with col3:
            # Single text boxes for input without dynamic list clutter
            warp_txt = st.text_input("Warp", help="Suggestions: " + ", ".join(get_suggestions("Warp")[:4]))
            w_slub_txt = st.text_input("Warp Slub")
            weft_txt = st.text_input("Weft")
            shade_txt = st.text_input("Shade")
            weave_txt = st.text_input("Weave Pattern")
            reed_txt = st.text_input("Reed")
            ppi_txt = st.text_input("PPI")
            greige_mtrs = st.text_input("Greige Meters")
            mkt_mtrs = st.text_input("Marketing Requested Meters")
            
        submit_btn = st.form_submit_button("Save Sample Data")
        
        if submit_btn:
            if not s_no or not brand or not ref_id:
                st.error("Fields S.no, Brand, aur Ref ID likhna zaroori hai.")
            elif str(s_no) in df_master["S.no"].values:
                st.error(f"S.no '{s_no}' pehle se maujood hai! Duplicate blocked.")
            else:
                new_row = {
                    "S.no": str(s_no), "Brand": brand, "Ref": ref_id, "Finish": finish_code,
                    "TR Code": tr_code, "Source": source, "Pending days in process": 0,
                    "Request Date": str(req_date_in), "Current Date": str(date.today()),
                    "Delivery date": str(del_date_in), "Ready date": "", "Status": status_init,
                    "Alert": "Normal", "Warp": warp_txt, "Warp Slub": w_slub_txt, "Weft": weft_txt,
                    "Shade": shade_txt, "Weave": weave_txt, "Reed": reed_txt, "Ppi": ppi_txt,
                    "Greige mtrs": greige_mtrs, "Marketing requested meter": mkt_mtrs, "Remarks": remarks
                }
                df_master = pd.concat([df_master, pd.DataFrame([new_row])], ignore_index=True)
                save_all_sheets(df_master, df_trials)
                st.success(f"🎉 Sample {s_no} Excel database mein save ho gaya!")
                st.rerun()

# --- OPTION 2: MAIN DASHBOARD & MASTER TABLES (WITH DELETE FEATURE) ---
elif menu == "Main Dashboard Visuals":
    st.header("实时 Pipeline Dashboard Visuals")
    
    if len(df_master) == 0:
        st.info("Database khali hai. Data load karne ke liye sidebar ya form use karein.")
    else:
        df_running = df_master[df_master["Status"] != "Submitted to marketing"]
        df_submitted = df_master[df_master["Status"] == "Submitted to marketing"]
        
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Active Floor Samples", len(df_running))
        kpi2.metric("Critical Alerts (<=3 Days)", len(df_running[df_running["Alert"].str.contains("Critical", na=False)]))
        kpi3.metric("Overdue / Delayed Runs", len(df_running[df_running["Alert"].str.contains("Delayed", na=False)]))
        kpi4.metric("Archived Submission Pool", len(df_submitted))
        
        st.divider()
        
        # INTERACTIVE CLICK TO VIEW & EDIT INDIVIDUAL SAMPLE
        st.markdown("### 🔍 Click to Quick View & Edit Individual Sample Details")
        selected_click_sno = st.selectbox("Koi bhi S.no select karein detail dekhne ya status change karne ke liye:", ["-- Select Sample --"] + df_master["S.no"].tolist())
        
        if selected_click_sno != "-- Select Sample --":
            s_idx = df_master[df_master["S.no"] == selected_click_sno].index[0]
            s_row = df_master.loc[s_idx]
            
            with st.expander(f"📖 SPEC SHEET FOR SAMPLE S.NO: {selected_click_sno}", expanded=True):
                ec1, ec2, ec3 = st.columns(3)
                with ec1:
                    updated_status = st.selectbox("Edit Process Stage:", STATUS_OPTIONS, index=STATUS_OPTIONS.index(s_row["Status"]))
                    updated_brand = st.text_input("Edit Brand name:", value=str(s_row["Brand"]))
                    updated_ref = st.text_input("Edit Ref Quality ID:", value=str(s_row["Ref"]))
                with ec2:
                    try: curr_del_date = datetime.strptime(str(s_row["Delivery date"]), "%Y-%m-%d").date()
                    except: curr_del_date = date.today()
                    updated_del = st.date_input("Edit Delivery Date:", value=curr_del_date)
                    updated_finish = st.text_input("Edit Finish Code:", value=str(s_row["Finish"]))
                    updated_tr = st.text_input("Edit TR Code:", value=str(s_row["TR Code"]))
                with ec3:
                    st.info(f"Days Pending: {s_row['Pending days in process']} days")
                    st.write(f"Warp Specs: {s_row['Warp']} / Slub: {s_row['Warp Slub']}")
                    st.write(f"Weft / Shade: {s_row['Weft']} / {s_row['Shade']}")
                    st.write(f"Reed / PPI: {s_row['Reed']} / {s_row['Ppi']}")
                
                updated_remarks = st.text_area("Edit Remarks:", value=str(s_row["Remarks"]))
                
                if st.button("Save Changes To Excel"):
                    df_master.at[s_idx, "Status"] = updated_status
                    df_master.at[s_idx, "Brand"] = updated_brand
                    df_master.at[s_idx, "Ref"] = updated_ref
                    df_master.at[s_idx, "Finish"] = updated_finish
                    df_master.at[s_idx, "Delivery date"] = str(updated_del)
                    df_master.at[s_idx, "TR Code"] = updated_tr
                    df_master.at[s_idx, "Remarks"] = updated_remarks
                    
                    if updated_status == "Submitted to marketing" and s_row["Status"] != "Submitted to marketing":
                        df_master.at[s_idx, "Ready date"] = str(date.today())
                    
                    df_master = calculate_metrics_and_alerts(df_master)
                    save_all_sheets(df_master, df_trials)
                    st.success("Changes save ho gaye!")
                    st.rerun()
        
        st.divider()
        
        # --- ROW DELETE INTERFACE SECTION ---
        st.markdown("### 🗑️ Delete Sample Row (Galat Excel Import Data Saaf Karein)")
        del_col1, del_col2 = st.columns([3, 1])
        with del_col1:
            delete_target_sno = st.selectbox("Excel se aya koi galat data ya row delete karne ke liye S.no chunein:", ["-- Choose S.no to Delete --"] + df_master["S.no"].tolist())
        with del_col2:
            st.write("##")
            if st.button("🔴 Delete Row Permanently"):
                if delete_target_sno != "-- Choose S.no to Delete --":
                    # Remove from master and trial data frames
                    df_master = df_master[df_master["S.no"] != delete_target_sno]
                    df_trials = df_trials[df_trials["Sample ID"] != delete_target_sno]
                    save_all_sheets(df_master, df_trials)
                    st.success(f"Row S.no {delete_target_sno} database aur Excel dono se permanently delete ho gayi!")
                    st.rerun()
                else:
                    st.error("Pehle delete karne ke liye ek S.no select karein.")

        st.divider()
        
        # Charts section
        g1, g2 = st.columns(2)
        with g1:
            st.subheader("🏭 Running Production Stages Count")
            status_counts = df_running["Status"].value_counts().reindex(STATUS_OPTIONS[:-1], fill_value=0).reset_index()
            status_counts.columns = ["Route Stage", "Active Counts"]
            fig1 = px.bar(status_counts, x="Active Counts", y="Route Stage", orientation="h", color="Active Counts", color_continuous_scale="Viridis", text_auto=True)
            st.plotly_chart(fig1, use_container_width=True)
        with g2:
            st.subheader("🎯 Brand Volume Segregation")
            brand_counts = df_master["Brand"].value_counts().reset_index()
            brand_counts.columns = ["Brand Segment", "Total Samples"]
            fig2 = px.pie(brand_counts, values="Total Samples", names="Brand Segment", hole=0.4, color_discrete_sequence=px.colors.sequential.Tealgrn)
            st.plotly_chart(fig2, use_container_width=True)
            
        st.divider()
        st.subheader("📋 Active Running Process Queue (Horizontal Rows)")
        st.dataframe(df_running, use_container_width=True)
        
        st.subheader("📁 Archive Vault: Submitted to Marketing")
        st.dataframe(df_submitted, use_container_width=True)

# --- OPTION 3: SEARCH & INTERNAL DETAIL VIEWER (TRIALS) ---
elif menu == "Search & Internal Detail Viewer":
    st.header("🔍 Technical Spec Sheet & Finishing Trials Config")
    
    if len(df_master) == 0:
        st.warning("No records found.")
    else:
        selected_sno = st.selectbox("Select Target S.no:", df_master["S.no"].tolist())
        sample_idx = df_master[df_master["S.no"] == selected_sno].index[0]
        row_data = df_master.loc[sample_idx]
        
        st.subheader(f"ℹ️ Core Parameters for S.no: {selected_sno}")
        col_view1, col_view2, col_view3 = st.columns(3)
        with col_view1:
            st.write(f"**Brand:** {row_data['Brand']}")
            st.write(f"**Ref Quality:** {row_data['Ref']}")
            st.write(f"**Finish Code:** {row_data['Finish']}")
            st.write(f"**Status:** :blue[{row_data['Status']}]")
        with col_view2:
            st.write(f"**Request Date:** {row_data['Request Date']}")
            st.write(f"**Delivery Target:** {row_data['Delivery date']}")
            st.write(f"**Pending Days:** {row_data['Pending days in process']} days")
        with col_view3:
            st.write(f"**Yarn Specs:** {row_data['Warp']} / {row_data['Weft']}")
            st.write(f"**Shade / Weave:** {row_data['Shade']} / {row_data['Weave']}")
            st.write(f"**Reed / PPI:** {row_data['Reed']} / {row_data['Ppi']}")
            
        st.divider()
        st.subheader("🧪 Finishing Details & Trial Tracker")
        active_trials = df_trials[df_trials["Sample ID"] == selected_sno]
        
        if len(active_trials) > 0:
            for _, t_row in active_trials.iterrows():
                t_col1, t_col2, t_col3 = st.columns([2, 4, 2])
                with t_col1: st.markdown(f"**{t_row['Trial Number']}**")
                with t_col2: st.write(f"Parameters: {t_row['Parameters']}")
                with t_col3:
                    if t_row["Status_OK"] == "OK ✅": st.success("Trial OK ✅")
                    else: st.warning("Pending Run")
                    
        with st.form("add_trial_form"):
            trial_num = st.selectbox("Trial ID:", ["Trial-1 Parameters", "Trial-2 Parameters", "Trial-3 Parameters", "Trial-4 Parameters"])
            trial_params = st.text_input("Chemical Recipe / Finishing Parameters Settings:")
            trial_status_ok = st.checkbox("Mark this chemical run as 'Trial OK'?")
            
            if st.form_submit_button("Log Finishing Trial"):
                if trial_params:
                    match_mask = (df_trials["Sample ID"] == selected_sno) & (df_trials["Trial Number"] == trial_num)
                    status_str = "OK ✅" if trial_status_ok else "Pending"
                    
                    if match_mask.any():
                        df_trials.loc[match_mask, "Parameters"] = trial_params
                        df_trials.loc[match_mask, "Status_OK"] = status_str
                    else:
                        new_trial_row = {"Sample ID": selected_sno, "Trial Number": trial_num, "Parameters": trial_params, "Status_OK": status_str, "Updated Date": str(date.today())}
                        df_trials = pd.concat([df_trials, pd.DataFrame([new_trial_row])], ignore_index=True)
                        
                    save_all_sheets(df_master, df_trials)
                    st.success("Finishing Trial catalog updated!")
                    st.rerun()
