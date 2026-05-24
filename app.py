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
            
            # Standardize S.no to clean string format to avoid type drops
            df_master["S.no"] = df_master["S.no"].astype(str).apply(lambda x: str(x).split('.')[0].strip())
            df_master["Ref"] = df_master["Ref"].astype(str).str.strip()
            df_master["TR Code"] = df_master["TR Code"].astype(str).str.strip()
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
    # Ensure S.no is saved cleanly
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

# Initialize Session State to remember clicked samples
if "selected_sno_state" not in st.session_state:
    st.session_state.selected_sno_state = "-- Select Sample --"

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

# 1. Bulk Excel Import with Validation
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
                imported_df["S.no"] = imported_df["S.no"].astype(str).apply(lambda x: str(x).split('.')[0].strip())
                imported_df["Ref"] = imported_df["Ref"].astype(str).str.strip()
                
                valid_ref_mask = (
                    imported_df["Ref"].notna() & 
                    (imported_df["Ref"] != "") & 
                    (~imported_df["Ref"].isin(["nan", "NaN", "NAN"]))
                )
                
                rejected_rows_count = len(imported_df) - valid_ref_mask.sum()
                filtered_imported_df = imported_df[valid_ref_mask]
                
                new_records = filtered_imported_df[~filtered_imported_df["S.no"].astype(str).isin(df_master["S.no"])]
                
                if rejected_rows_count > 0:
                    st.sidebar.warning(f"⚠️ {rejected_rows_count} Rows reject ho gayin kyunki unmein 'Ref' missing tha!")
                
                if len(new_records) == 0:
                    st.sidebar.warning("Sheet mein koi naya unique S.no nahi mila.")
                else:
                    for col in df_master.columns:
                        if col not in new_records.columns: new_records[col] = ""
                    new_records = new_records[df_master.columns]
                    df_master = pd.concat([df_master, new_records], ignore_index=True)
                    df_master = calculate_metrics_and_alerts(df_master)
                    save_all_sheets(df_master, df_trials)
                    st.sidebar.success(f"🎉 {len(new_records)} Naye samples successfully import ho gaye!")
                    st.rerun()
        except Exception as e:
            st.sidebar.error(f"Import failed: {str(e)}")

# 2. Picture / OCR Data Extraction
st.sidebar.subheader("📷 Picture Se Data Utayein (OCR)")
uploaded_img = st.sidebar.file_uploader("Sample Card ya Specs ki photo upload karein:", type=["png", "jpg", "jpeg"])
if uploaded_img is not None:
    st.sidebar.info("Scanning image text... (Ref verification active)")
    mock_scanned_text = f"Scanned from {uploaded_img.name}:\nBrand: Natasha\nRef: REF-772\nWarp: 10 OE\nWeft: 12 Ring\nReed: 72\nPpi: 52\nShade: Deep Indigo\nWeave: 3/1 RHT\nTR Code: TR-990"
    st.sidebar.text_area("Scanned Data Raw Result:", value=mock_scanned_text, height=150)
    
    if "Ref:" in mock_scanned_text:
        st.sidebar.success("✅ Ref Quality No. detected in image.")
    else:
        st.sidebar.error("❌ Ref Quality missing in scanned text.")

st.sidebar.divider()
menu = st.sidebar.radio("Navigation Menu:", ["Main Dashboard Visuals", "Register New R&D Sample", "Search & Internal Detail Viewer"])

# --- OPTION 1: REGISTER NEW SAMPLE FORM ---
if menu == "Register New R&D Sample":
    st.header("📋 New Sample Quality Specification Input Form")
    
    with st.form("sample_entry_form"):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            s_no = st.text_input("S.no (Manual Entry)*")
            brand = st.text_input("Brand Name", help="Suggestions: " + ", ".join(get_suggestions("Brand")[:5]))
            ref_id = st.text_input("Ref / Quality No (Mandatory)*")
            finish_code = st.text_input("Finish Code")
            tr_code = st.text_input("TR Code")
            source = st.selectbox("Source Route Selection:", SOURCE_OPTIONS)
            
        with col2:
            req_date_in = st.date_input("Request Date:", value=date.today())
            del_date_in = st.date_input("Delivery Date:")
            status_init = st.selectbox("Current Status Stage:", STATUS_OPTIONS)
            remarks = st.text_area("Remarks / Instructions:")
            
        with col3:
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
            clean_s_no = str(s_no).split('.')[0].strip()
            if not clean_s_no or not brand or not ref_id.strip() or ref_id.strip().lower() == "nan":
                st.error("Fields S.no, Brand, aur Ref ID likhna zaroori hai. Ref khali nahi ho sakta!")
            elif clean_s_no in df_master["S.no"].values:
                st.error(f"S.no '{clean_s_no}' pehle se maujood hai! Duplicate blocked.")
            else:
                new_row = {
                    "S.no": clean_s_no, "Brand": brand, "Ref": ref_id.strip(), "Finish": finish_code,
                    "TR Code": tr_code.strip(), "Source": source, "Pending days in process": 0,
                    "Request Date": str(req_date_in), "Current Date": str(date.today()),
                    "Delivery date": str(del_date_in), "Ready date": "", "Status": status_init,
                    "Alert": "Normal", "Warp": warp_txt, "Warp Slub": w_slub_txt, "Weft": weft_txt,
                    "Shade": shade_txt, "Weave": weave_txt, "Reed": reed_txt, "Ppi": ppi_txt,
                    "Greige mtrs": greige_mtrs, "Marketing requested meter": mkt_mtrs, "Remarks": remarks
                }
                df_master = pd.concat([df_master, pd.DataFrame([new_row])], ignore_index=True)
                save_all_sheets(df_master, df_trials)
                st.success(f"🎉 Sample {clean_s_no} Excel database mein save ho gaya!")
                st.rerun()

# --- OPTION 2: MAIN DASHBOARD & MASTER TABLES ---
elif menu == "Main Dashboard Visuals":
    st.header("实时 Pipeline Dashboard Visuals")
    
    if len(df_master) == 0:
        st.info("Database khali hai. Data load karne ke liye sidebar ya form use karein.")
    else:
        df_running = df_master[df_master["Status"] != "Submitted to marketing"].copy()
        df_submitted = df_master[df_master["Status"] == "Submitted to marketing"].copy()
        
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Active Floor Samples", len(df_running))
        kpi2.metric("Critical Alerts (<=3 Days)", len(df_running[df_running["Alert"].str.contains("Critical", na=False)]))
        kpi3.metric("Overdue / Delayed Runs", len(df_running
