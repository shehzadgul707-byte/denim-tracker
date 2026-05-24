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
    
    # Save template structure
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master_Data", index=False)
        df_trials.to_excel(writer, sheet_name="Trial_Data", index=False)
        
    return df_master, df_trials

def save_all_sheets(df_master, df_trials):
    """Saves both dataframes into separate sheets in the same Excel file horizontally aligned."""
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master_Data", index=False)
        df_trials.to_excel(writer, sheet_name="Trial_Data", index=False)

# Load current data state
df_master, df_trials = load_all_sheets()

# Helper to pull unique values for dynamic dynamic dropdown options
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
    
    for idx, row in df.iterrows():
        # Pending Days logic from Request Date
        try:
            req_date = pd.to_datetime(row['Request Date']).date()
            df.at[idx, 'Pending days in process'] = (today - req_date).days
        except:
            df.at[idx, 'Pending days in process'] = 0
            
        # Alert logic (3 days remaining check and not submitted)
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

menu = st.sidebar.radio("Navigation Menu:", ["Main Dashboard Visuals", "Register New R&D Sample", "Search & Internal Detail Viewer"])

# --- OPTION 1: REGISTER NEW SAMPLE FORM ---
if menu == "Register New R&D Sample":
    st.header("📋 New Sample Quality Specification Input Form")
    st.markdown("Enter technical attributes below. Fields like fabric specs dynamically adapt based on previously stored parameters.")
    
    # Dynamic Lists fetched straight from Excel data history
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
            
            brand_sel = st.selectbox("Select Brand Name (Or choose manual entry below):", ["-- Manual Entry --"] + dynamic_brands)
            brand_manual = st.text_input("Brand Manual Entry (If not found in dropdown list):")
            final_brand = brand_manual.strip() if brand_sel == "-- Manual Entry --" else brand_sel
            
            ref_id = st.text_input("Ref / Quality No (Manual)")
            finish_code = st.text_input("Finish Code (Manual)")
            tr_code = st.text_input("TR Code (Manual)")
            source = st.selectbox("Source Route Selection:", SOURCE_OPTIONS)
            
        with col2:
            req_date_in = st.date_input("Request Date (Sample Start):", value=date.today())
            del_date_in = st.date_input("Delivery Date (Committed to Marketing via Email):")
            status_init = st.selectbox("Current Processing Stage Status:", STATUS_OPTIONS)
            remarks = st.text_area("R&D Sample Remarks / Special Instructions:")
            
        with col3:
            # Fabric Spec Dynamic Check-boxes and Manual Entries combined
            warp_sel = st.selectbox("Warp Yarn Spec:", ["-- New Option --"] + dynamic_warps)
            warp_txt = st.text_input("Enter New Warp:") if warp_sel == "-- New Option --" else warp_sel
            
            w_slub_sel = st.selectbox("Warp Slub Profile:", ["-- New Option --"] + dynamic_slubs)
            w_slub_txt = st.text_input("Enter New Slub Profile:") if w_slub_sel == "-- New Option --" else w_slub_sel
            
            weft_sel = st.selectbox("Weft Yarn Spec:", ["-- New Option --"] + dynamic_wefts)
            weft_txt = st.text_input("Enter New Weft:") if weft_sel == "-- New Option --" else weft_sel
            
            shade_sel = st.selectbox("Color Shade Target:", ["-- New Option --"] + dynamic_shades)
            shade_txt = st.text_input("Enter New Shade:") if shade_sel == "-- New Option --" else shade_sel
            
            weave_sel = st.selectbox("Weave Pattern Construction:", ["-- New Option --"] + dynamic_weaves)
            weave_txt = st.text_input("Enter New Weave Type:") if weave_sel == "-- New Option --" else weave_sel
            
            reed_sel = st.selectbox("Reed Count:", ["-- New Option --"] + dynamic_reeds)
            reed_txt = st.text_input("Enter New Reed:") if reed_sel == "-- New Option --" else reed_sel
            
            ppi_sel = st.selectbox("Picks Per Inch (PPI):", ["-- New Option --"] + dynamic_ppis)
            ppi_txt = st.text_input("Enter New PPI Count:") if ppi_sel == "-- New Option --" else ppi_sel
            
            greige_mtrs = st.text_input("Greige Meters:")
            mkt_mtrs = st.text_input("Marketing Requested Meters:")
            
        submit_btn = st.form_submit_button("Save Sample and Auto-Update Specs List")
        
        if submit_btn:
            if not s_no or not final_brand or not ref_id:
                st.error("Error: S.no, Brand, and Ref ID fields mandatory.")
            elif s_no in df_master["S.no"].astype(str).values:
                st.error(f"S.no code '{s_no}' already exists in the database. Duplicate entries blocked.")
            else:
                new_row = {
                    "S.no": s_no, "Brand": final_brand, "Ref": ref_id, "Finish": finish_code,
                    "TR Code": tr_code, "Source": source, "Pending days in process": 0,
                    "Request Date": str(req_date_in), "Current Date": str(date.today()),
                    "Delivery date": str(del_date_in), "Ready date": "", "Status": status_init,
                    "Alert": "Normal", "Warp": warp_txt, "Warp Slub": w_slub_txt, "Weft": weft_txt,
                    "Shade": shade_txt, "Weave": weave_txt, "Reed": reed_txt, "Ppi": ppi_txt,
                    "Greige mtrs": greige_mtrs, "Marketing requested meter": mkt_mtrs, "Remarks": remarks
                }
                
                df_master = pd.concat([df_master, pd.DataFrame([new_row])], ignore_index=True)
                save_all_sheets(df_master, df_trials)
                st.success(f"🎉 New Sample {s_no} with custom technical constraints saved directly into Excel horizontal pipeline sheets!")
                st.rerun()

# --- OPTION 2: MAIN LIVE DASHBOARD ---
elif menu == "Main Dashboard Visuals":
    st.header("📊 Real-Time Operations Pipeline Tracker")
    
    if len(df_master) == 0:
        st.info("Database is empty. Please register R&D samples from side menu to populate operational charts.")
    else:
        # Separate running items vs submitted archive items
        df_running = df_master[df_master["Status"] != "Submitted to marketing"]
        df_submitted = df_master[df_master["Status"] == "Submitted to marketing"]
        
        # --- TOP LEVEL SUMMARY NUMBERS ---
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Active Floor Samples", len(df_running))
        kpi2.metric("Critical Alerts (<=3 Days)", len(df_running[df_running["Alert"].str.contains("Critical", na=False)]))
        kpi3.metric("Overdue / Delayed Runs", len(df_running[df_running["Alert"].str.contains("Delayed", na=False)]))
        kpi4.metric("Archived Submission Pool", len(df_submitted))
        
        st.divider()
        
        # --- GRAPHICAL SEGREGATIONS ---
        g1, g2 = st.columns(2)
        
        with g1:
            st.subheader("🏭 Status Stage Distribution (Active Floor)")
            status_counts = df_running["Status"].value_counts().reindex(STATUS_OPTIONS[:-1], fill_value=0).reset_index()
            status_counts.columns = ["Route Stage", "Active Counts"]
            fig1 = px.bar(status_counts, x="Active Counts", y="Route Stage", orientation="h", 
                          color="Active Counts", color_continuous_scale="Viridis", text_auto=True)
            st.plotly_chart(fig1, use_container_width=True)
            
        with g2:
            st.subheader("🎯 Brand Volume Breakdown")
            brand_counts = df_master["Brand"].value_counts().reset_index()
            brand_counts.columns = ["Brand Segment", "Total Samples Developed"]
            fig2 = px.pie(brand_counts, values="Total Samples Developed", names="Brand Segment", 
                          hole=0.4, color_discrete_sequence=px.colors.sequential.Tealgrn)
            st.plotly_chart(fig2, use_container_width=True)
            
        st.divider()
        
        # --- SUBMISSION HISTORY BY MONTH GRAPH ---
        st.subheader("📅 Monthly Marketing Submission Trends (Graphical Representation)")
        if len(df_submitted) > 0:
            df_submitted_copy = df_submitted.copy()
            df_submitted_copy["Ready date"] = pd.to_datetime(df_submitted_copy["Ready date"], errors='coerce')
            df_submitted_copy["Month"] = df_submitted_copy["Ready date"].dt.strftime('%B %Y')
            
            month_df = df_submitted_copy.groupby("Month").size().reset_index(name="Samples Transferred")
            fig_month = px.line(month_df, x="Month", y="Samples Transferred", title="Total Output Handover by Month", 
                                markers=True, line_shape="spline", color_discrete_sequence=["#2E8B57"])
            st.plotly_chart(fig_month, use_container_width=True)
        else:
            st.info("No samples historical records found in submission pools yet to construct graphical charts.")
            
        st.divider()
        
        # --- EXCEL EXPORT MASTER HORIZONTAL TABLES ---
        st.subheader("📋 Active Running Process Queue (Synced with Master Excel)")
        st.dataframe(df_running, use_container_width=True)
        
        st.subheader("📁 Archive Vault: Submitted to Marketing Section")
        if len(df_submitted) > 0:
            st.dataframe(df_submitted, use_container_width=True)
        else:
            st.caption("Archive table currently empty. Move status to 'Submitted to marketing' inside individual viewer to view records here.")

# --- OPTION 3: SEARCH & INTERNAL DETAIL VIEWER (TRIALS MECHANISM) ---
elif menu == "Search & Internal Detail Viewer":
    st.header("🔍 Technical Spec Sheet & Finishing Trials Configuration")
    
    if len(df_master) == 0:
        st.warning("No tracked sample records available.")
    else:
        sample_list = df_master["S.no"].astype(str).tolist()
        selected_sno = st.selectbox("Select Target S.no to check profile & add Finishing Trials:", sample_list)
        
        # Extract row entry metrics
        sample_idx = df_master[df_master["S.no"].astype(str) == selected_sno].index[0]
        row_data = df_master.loc[sample_idx]
        
        # Display layout details structured horizontally
        st.subheader(f"ℹ️ Core Parameters for Sample S.no: {selected_sno} [Brand: {row_data['Brand']}]")
        
        col_view1, col_view2, col_view3 = st.columns(3)
        with col_view1:
            st.write(f"**Ref Quality:** {row_data['Ref']}")
            st.write(f"**Finish Code:** {row_data['Finish']}")
            st.write(f"**TR Code:** {row_data['TR Code']}")
            st.write(f"**Source Variant:** {row_data['Source']}")
            st.write(f"**Current Status:** :blue[{row_data['Status']}]")
            
        with col_view2:
            st.write(f"**Request Date:** {row_data['Request Date']}")
            st.write(f"**Delivery Target:** {row_data['Delivery date']}")
            st.write(f"**Pending Days:** {row_data['Pending days in process']} days")
            st.write(f"**Current System Date:** {row_data['Current Date']}")
            st.write(f"**Alert Status:** :orange[{row_data['Alert']}]")
            
        with col_view3:
            st.write(f"**Yarn Warp/Slub:** {row_data['Warp']} / {row_data['Warp Slub']}")
            st.write(f"**Yarn Weft:** {row_data['Weft']}")
            st.write(f"**Shade Target:** {row_data['Shade']}")
            st.write(f"**Construction Weave:** {row_data['Weave']}")
            st.write(f"**Reed / PPI:** {row_data['Reed']} / {row_data['Ppi']}")
            st.write(f"**Meters (Greige/Req):** {row_data['Greige mtrs']} / {row_data['Marketing requested meter']}")
            
        st.text_area("Master Remarks:", value=str(row_data['Remarks']), disabled=True)
        
        st.divider()
        
        # --- FINISHING DETAIL SECTION (TRIALS TRACKING) ---
        st.subheader("🧪 Finishing Detail Section (Trials Logger)")
        
        # Fetch relevant sub-trial rows
        active_trials = df_trials[df_trials["Sample ID"].astype(str) == selected_sno]
        
        if len(active_trials) == 0:
            st.info("No finishing trials documented for this quality number yet.")
        else:
            # Render trials dynamically showing text parameters and check-boxes status
            for _, t_row in active_trials.iterrows():
                t_col1, t_col2, t_col3 = st.columns([1, 4, 2])
                with t_col1:
                    st.markdown(f"**{t_row['Trial Number']}**")
                with t_col2:
                    st.write(f"**Parameters Log:** {t_row['Parameters']}")
                with t_col3:
                    if t_row["Status_OK"] == "OK ✅":
                        st.success("Status: Trial OK ✅")
                    else:
                        st.warning("Status: Running / Pending")
                        
        st.markdown("#### Append New Finishing Trial Parameter Runtime Logs")
        with st.form("add_trial_form"):
            trial_num = st.selectbox("Trial ID to Record:", ["Trial-1 Parameters", "Trial-2 Parameters", "Trial-3 Parameters", "Trial-4 Parameters"])
            trial_params = st.text_input("Enter Finishing Parameters / Chemical Recipe Settings:")
            trial_status_ok = st.checkbox("Mark this specific chemical run as 'Trial OK'?")
            
            submit_trial = st.form_submit_button("Log Finishing Trial to Database")
            if submit_trial:
                if not trial_params:
                    st.error("Please add parameter technical logs.")
                else:
                    # Check if trial already logged to prevent duplicates, or overwrite
                    match_mask = (df_trials["Sample ID"].astype(str) == selected_sno) & (df_trials["Trial Number"] == trial_num)
                    status_str = "OK ✅" if trial_status_ok else "Pending"
                    
                    if match_mask.any():
                        df_trials.loc[match_mask, "Parameters"] = trial_params
                        df_trials.loc[match_mask, "Status_OK"] = status_str
                        df_trials.loc[match_mask, "Updated Date"] = str(date.today())
                    else:
                        new_trial_row = {
                            "Sample ID": selected_sno, "Trial Number": trial_num,
                            "Parameters": trial_params, "Status_OK": status_str,
                            "Updated Date": str(date.today())
                        }
                        df_trials = pd.concat([df_trials, pd.DataFrame([new_trial_row])], ignore_index=True)
                        
                    save_all_sheets(df_master, df_trials)
                    st.success("Finishing Trial catalog updated successfully inside connected spreadsheet matrices.")
                    st.rerun()
                    
        st.divider()
        
        # --- UPDATE STAGE STATUS OR SIGN-OFF AS SUBMITTED ---
        st.subheader("🔄 Change Dynamic Route Stage Status")
        new_stage_update = st.selectbox("Assign New Processing Stage:", STATUS_OPTIONS, index=STATUS_OPTIONS.index(row_data["Status"]))
        
        if st.button("Commit Status Change"):
            df_master.at[sample_idx, "Status"] = new_stage_update
            if new_stage_update == "Submitted to marketing":
                df_master.at[sample_idx, "Ready date"] = str(date.today())
                st.balloons()
                
            save_all_sheets(df_master, df_trials)
            st.success(f"Sample {selected_sno} pipeline route updated to: '{new_stage_update}'")
            st.rerun()
