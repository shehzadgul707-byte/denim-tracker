import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, date
import os

# Set layout to wide for dashboard
st.set_page_config(page_title="Denim R&D Advanced Tracker", layout="wide")

EXCEL_FILE = "Denim_Master_Database.xlsx"

# Standard fixed dropdown lists
BRAND_OPTIONS = ["Natasha", "Tommy", "David", "Guess", "Non-Nomination"]
SOURCE_OPTIONS = ["Existing", "Scratch", "Production beam"]
STATUS_OPTIONS = ["Yarn Demand", "Dyeing", "Weaving", "Finishing", "Inspection", "Submitted to marketing"]

# --- HELPER FUNCTIONS FOR EXCEL HANDLING ---
def load_all_sheets():
    """Loads Master and Trial data from Excel, initializing them if file does not exist."""
    if os.path.exists(EXCEL_FILE):
        try:
            df_master = pd.read_excel(EXCEL_FILE, sheet_name="Master_Data")
            df_trials = pd.read_excel(EXCEL_FILE, sheet_name="Trial_Data")
            # Ensure S.no is treated as string to avoid merge/search mismatches
            df_master["S.no"] = df_master["S.no"].astype(str)
            return df_master, df_trials
        except Exception:
            pass
            
    # Initialize empty dataframes with your exact columns if file not found
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

# Helper to pull unique values for dynamic dropdown options
def get_dynamic_options(column_name, default_list=None):
    if default_list is None:
        default_list = []
    if len(df_master) > 0 and column_name in df_master.columns:
        existing_vals = df_master[column_name].dropna().unique().tolist()
        return sorted(list(set(default_list + [str(x) for x in existing_vals if str(x).strip() != ""])))
    return sorted(default_list)

# --- AUTOMATIC FORMULA CALCULATIONS ---
def calculate_metrics_and_alerts(df):
    """Calculates running pending days and alerts based on Delivery dates dynamically."""
    if len(df) == 0:
        return df
    
    today = date.today()
    df['Current Date'] = today.strftime("%Y-%m-%d")
    df["S.no"] = df["S.no"].astype(str)
    
    for idx, row in df.iterrows():
        # Pending Days logic from Request Date
        try:
            req_date = pd.to_datetime(row['Request Date']).date()
            df.at[idx, 'Pending days in process'] = (today - req_date).days
        except:
            df.at[idx, 'Pending days in process'] = 0
            
        # Alert logic (3 days remaining check)
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

# Apply calculations instantly on loaded data
df_master = calculate_metrics_and_alerts(df_master)

# --- NAVIGATION CONTROLS ---
st.title("👖 Denim R&D Sample Tracker & Quality Dashboard")
st.markdown("Advanced manufacturing workflow pipeline matching production specifications with interactive sample history logs.")
st.divider()

# --- SIDEBAR: EXCEL IMPORT & CONFIG ---
st.sidebar.header("⚙️ Data Management")

# Bulk Import Feature
st.sidebar.subheader("📥 Import Bulk Data from Excel")
uploaded_file = st.sidebar.file_uploader("Upload an Excel file to merge/import data:", type=["xlsx"])

if uploaded_file is not None:
    if st.sidebar.button("Process & Import Excel Data"):
        try:
            imported_df = pd.read_excel(uploaded_file)
            # Standardize columns to match system database
            required_cols = ["S.no", "Brand", "Ref", "Status"]
            missing_cols = [c for c in required_cols if c not in imported_df.columns]
            
            if missing_cols:
                st.sidebar.error(f"Excel file missing important columns: {missing_cols}")
            else:
                # Merge or append based on S.no
                imported_df["S.no"] = imported_df["S.no"].astype(str)
                
                # Filter out entries that already exist to avoid crashes, or combine
                new_records = imported_df[~imported_df["S.no"].isin(df_master["S.no"])]
                
                if len(new_records) == 0:
                    st.sidebar.warning("No new unique S.no found in uploaded sheet.")
                else:
                    # Fill missing columns with blank defaults to match schema
                    for col in df_master.columns:
                        if col not in new_records.columns:
                            new_records[col] = ""
                            
                    new_records = new_records[df_master.columns] # Reorder
                    df_master = pd.concat([df_master, new_records], ignore_index=True)
                    df_master = calculate_metrics_and_alerts(df_master)
                    save_all_sheets(df_master, df_trials)
                    st.sidebar.success(f"Successfully imported {len(new_records)} new rows from Excel!")
                    st.rerun()
        except Exception as e:
            st.sidebar.error(f"Import failed: {str(e)}")

st.sidebar.divider()
menu = st.sidebar.radio("Navigation Menu:", ["Main Dashboard Visuals", "Register New R&D Sample", "Search & Internal Detail Viewer"])

# --- OPTION 1: REGISTER NEW SAMPLE FORM ---
if menu == "Register New R&D Sample":
    st.header("📋 New Sample Quality Specification Input Form")
    
    dynamic_brands = get_dynamic_options("Brand", BRAND_OPTIONS)
    dynamic_warps = get_dynamic_options("Warp")
    dynamic_slubs = get_dynamic_options("Warp Slub")
    dynamic_wefts = get_dynamic_options("Weft")
    dynamic_shades = get_dynamic_options("Shade")
    dynamic_weaves = get_dynamic_options("Weave")
    dynamic_reeds = get_dynamic_options("Reed")
    dynamic_ppis = get_dynamic_options("Ppi")
    
    with st.form("sample_entry_form"):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            s_no = st.text_input("S.no (Manual Entry)*")
            brand_sel = st.selectbox("Select Brand Name:", ["-- Manual Entry --"] + dynamic_brands)
            brand_manual = st.text_input("Brand Manual Entry:")
            final_brand = brand_manual.strip() if brand_sel == "-- Manual Entry --" else brand_sel
            ref_id = st.text_input("Ref / Quality No (Manual)")
            finish_code = st.text_input("Finish Code (Manual)")
            tr_code = st.text_input("TR Code (Manual)")
            source = st.selectbox("Source Route Selection:", SOURCE_OPTIONS)
            
        with col2:
            req_date_in = st.date_input("Request Date:", value=date.today())
            del_date_in = st.date_input("Delivery Date:")
            status_init = st.selectbox("Current Status Stage:", STATUS_OPTIONS)
            remarks = st.text_area("Remarks / Instructions:")
            
        with col3:
            warp_sel = st.selectbox("Warp Yarn:", ["-- New Option --"] + dynamic_warps)
            warp_txt = st.text_input("New Warp:") if warp_sel == "-- New Option --" else warp_sel
            
            w_slub_sel = st.selectbox("Warp Slub:", ["-- New Option --"] + dynamic_slubs)
            w_slub_txt = st.text_input("New Slub:") if w_slub_sel == "-- New Option --" else w_slub_sel
            
            weft_sel = st.selectbox("Weft Yarn:", ["-- New Option --"] + dynamic_wefts)
            weft_txt = st.text_input("New Weft:") if weft_sel == "-- New Option --" else weft_sel
            
            shade_sel = st.selectbox("Shade Target:", ["-- New Option --"] + dynamic_shades)
            shade_txt = st.text_input("New Shade:") if shade_sel == "-- New Option --" else shade_sel
            
            weave_sel = st.selectbox("Weave Pattern:", ["-- New Option --"] + dynamic_weaves)
            weave_txt = st.text_input("New Weave:") if weave_sel == "-- New Option --" else weave_sel
            
            reed_sel = st.selectbox("Reed Count:", ["-- New Option --"] + dynamic_reeds)
            reed_txt = st.text_input("New Reed:") if reed_sel == "-- New Option --" else reed_sel
            
            ppi_sel = st.selectbox("PPI:", ["-- New Option --"] + dynamic_ppis)
            ppi_txt = st.text_input("New PPI:") if ppi_sel == "-- New Option --" else ppi_sel
            
            greige_mtrs = st.text_input("Greige Meters:")
            mkt_mtrs = st.text_input("Marketing Requested Meters:")
            
        submit_btn = st.form_submit_button("Save Sample")
        
        if submit_btn:
            if not s_no or not final_brand or not ref_id:
                st.error("Fields S.no, Brand, and Ref ID are mandatory.")
            elif str(s_no) in df_master["S.no"].values:
                st.error(f"S.no '{s_no}' already exists! Duplicates blocked.")
            else:
                new_row = {
                    "S.no": str(s_no), "Brand": final_brand, "Ref": ref_id, "Finish": finish_code,
                    "TR Code": tr_code, "Source": source, "Pending days in process": 0,
                    "Request Date": str(req_date_in), "Current Date": str(date.today()),
                    "Delivery date": str(del_date_in), "Ready date": "", "Status": status_init,
                    "Alert": "Normal", "Warp": warp_txt, "Warp Slub": w_slub_txt, "Weft": weft_txt,
                    "Shade": shade_txt, "Weave": weave_txt, "Reed": reed_txt, "Ppi": ppi_txt,
                    "Greige mtrs": greige_mtrs, "Marketing requested meter": mkt_mtrs, "Remarks": remarks
                }
                df_master = pd.concat([df_master, pd.DataFrame([new_row])], ignore_index=True)
                save_all_sheets(df_master, df_trials)
                st.success(f"🎉 Sample {s_no} saved successfully!")
                st.rerun()

# --- OPTION 2: MAIN DASHBOARD & MASTER TABLES ---
elif menu == "Main Dashboard Visuals":
    st.header("实时 Pipeline Dashboard Visuals")
    
    if len(df_master) == 0:
        st.info("Database empty. Use sidebar spreadsheet uploader or register form to load data.")
    else:
        df_running = df_master[df_master["Status"] != "Submitted to marketing"]
        df_submitted = df_master[df_master["Status"] == "Submitted to marketing"]
        
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Active Floor Samples", len(df_running))
        kpi2.metric("Critical Alerts (<=3 Days)", len(df_running[df_running["Alert"].str.contains("Critical", na=False)]))
        kpi3.metric("Overdue / Delayed Runs", len(df_running[df_running["Alert"].str.contains("Delayed", na=False)]))
        kpi4.metric("Archived Submission Pool", len(df_submitted))
        
        st.divider()
        
        # INTERACTIVE "CLICK TO VIEW & EDIT INDIVIDUAL SAMPLE" OVERLAY
        st.markdown("### 🔍 Click to Quick View & Edit Individual Sample Details")
        all_samples_dropdown = df_master["S.no"].tolist()
        selected_click_sno = st.selectbox("Select any S.no to read profile details or live-update status on floor instantly:", ["-- Select Sample --"] + all_samples_dropdown)
        
        if selected_click_sno != "-- Select Sample --":
            s_idx = df_master[df_master["S.no"] == selected_click_sno].index[0]
            s_row = df_master.loc[s_idx]
            
            with st.expander(f"📖 OPEN SPEC SHEET FOR SAMPLE S.NO: {selected_click_sno}", expanded=True):
                ec1, ec2, ec3 = st.columns(3)
                with ec1:
                    updated_status = st.selectbox("Edit Current Process Stage:", STATUS_OPTIONS, index=STATUS_OPTIONS.index(s_row["Status"]), key="quick_status")
                    updated_brand = st.text_input("Edit Brand name:", value=str(s_row["Brand"]), key="quick_brand")
                    updated_ref = st.text_input("Edit Ref Quality ID:", value=str(s_row["Ref"]), key="quick_ref")
                    updated_finish = st.text_input("Edit Finish Code:", value=str(s_row["Finish"]), key="quick_finish")
                with ec2:
                    try:
                        curr_del_date = datetime.strptime(str(s_row["Delivery date"]), "%Y-%m-%d").date()
                    except:
                        curr_del_date = date.today()
                    updated_del = st.date_input("Edit Delivery Target Date:", value=curr_del_date, key="quick_del")
                    updated_tr = st.text_input("Edit TR Code:", value=str(s_row["TR Code"]), key="quick_tr")
                    updated_greige = st.text_input("Edit Greige Meters:", value=str(s_row["Greige mtrs"]), key="quick_greige")
                with ec3:
                    st.info(f"Days Pending in Process: {s_row['Pending days in process']} days")
                    st.write(f"Warp Specs: {s_row['Warp']} ({s_row['Warp Slub']})")
                    st.write(f"Weft / Shade Target: {s_row['Weft']} / {s_row['Shade']}")
                    st.write(f"Reed / PPI Setup: {s_row['Reed']} / {s_row['Ppi']}")
                
                updated_remarks = st.text_area("Edit Sample Remarks:", value=str(s_row["Remarks"]), key="quick_rem")
                
                if st.button("Save Changes to Excel Database", key="quick_save_btn"):
                    df_master.at[s_idx, "Status"] = updated_status
                    df_master.at[s_idx, "Brand"] = updated_brand
                    df_master.at[s_idx, "Ref"] = updated_ref
                    df_master.at[s_idx, "Finish"] = updated_finish
                    df_master.at[s_idx, "Delivery date"] = str(updated_del)
                    df_master.at[s_idx, "TR Code"] = updated_tr
                    df_master.at[s_idx, "Greige mtrs"] = updated_greige
                    df_master.at[s_idx, "Remarks"] = updated_remarks
                    
                    if updated_status == "Submitted to marketing" and s_row["Status"] != "Submitted to marketing":
                        df_master.at[s_idx, "Ready date"] = str(date.today())
                    
                    df_master = calculate_metrics_and_alerts(df_master)
                    save_all_sheets(df_master, df_trials)
                    st.success(f"Changes for Sample {selected_click_sno} updated successfully!")
                    st.rerun()
        
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
            st.write(f"**TR Code:** {row_data['TR Code']}")
            st.write(f"**Status:** :blue[{row_data['Status']}]")
        with col_view2:
            st.write(f"**Request Date:** {row_data['Request Date']}")
            st.write(f"**Delivery Target:** {row_data['Delivery date']}")
            st.write(f"**Pending Days:** {row_data['Pending days in process']} days")
            st.write(f"**Alert Status:** :orange[{row_data['Alert']}]")
        with col_view3:
            st.write(f"**Yarn Specs:** {row_data['Warp']} / {row_data['Weft']}")
            st.write(f"**Shade / Weave:** {row_data['Shade']} / {row_data['Weave']}")
            st.write(f"**Reed / PPI:** {row_data['Reed']} / {row_data['Ppi']}")
            st.write(f"**Meters:** {row_data['Greige mtrs']} Mtrs")
            
        st.divider()
        
        # Trials logger engine
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
