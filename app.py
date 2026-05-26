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

if "current_view" not in st.session_state:
    st.session_state.current_view = "Overall"

# ====================== HELPER FUNCTIONS ======================
def read_excel_as_string(file_path_or_buffer):
    return pd.read_excel(file_path_or_buffer, dtype=str)

def clean_dataframe(df):
    df = df.copy()
    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()
    return df

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

# KPI Boxes
k1, k2, k3 = st.columns(3)
with k1:
    if st.button(f"**Total Samples On Floor**\n\n**{len(df_running)}**", use_container_width=True):
        st.session_state.current_view = "Overall"
with k2:
    if st.button(f"**Development Section**\n\n**{len(df_dev)}**", use_container_width=True):
        st.session_state.current_view = "Development"
with k3:
    if st.button(f"**Repeat Section**\n\n**{len(df_repeat)}**", use_container_width=True):
        st.session_state.current_view = "Repeat"

st.markdown("---")

# ====================== UPLOAD SECTION (New Feature) ======================
st.markdown("## 📥 Excel Upload Center")

upload_type = st.radio(
    "Kis type ki file upload kar rahe hain?",
    options=["Full Pipeline File (Development + Repeat)", "Only Repeat Sampling Sheet"],
    horizontal=True
)

uploaded_file = st.file_uploader("Excel File Select Karein", type=["xlsx"], key="uploader")

if uploaded_file is not None:
    if st.button("🚀 Process & Upload Data", type="primary", use_container_width=True):
        try:
            imported_df = read_excel_as_string(uploaded_file)
            imported_df = clean_dataframe(imported_df)
            
            if "Ref" not in imported_df.columns:
                st.error("Ref column zaroori hai")
            else:
                if "Ppi" in imported_df.columns and "Picks" not in imported_df.columns:
                    imported_df.rename(columns={"Ppi": "Picks"}, inplace=True)
                
                valid_df = imported_df[imported_df["Ref"].str.strip() != ""].copy()
                
                for col in COLUMNS_STRUCTURE:
                    if col not in valid_df.columns:
                        valid_df[col] = ""
                
                final_df = valid_df[COLUMNS_STRUCTURE].copy()
                final_df = calculate_metrics_and_alerts(final_df)

                if upload_type == "Full Pipeline File (Development + Repeat)":
                    # Full Replace
                    save_all_sheets(final_df)
                    st.success(f"🎉 Full Pipeline - {len(final_df)} samples loaded!")
                
                else:  # Only Repeat Sampling Sheet
                    # Append / Update Repeat samples
                    repeat_mask = final_df["Source"].str.contains("existing|production", case=False, na=False)
                    repeat_samples = final_df[repeat_mask].copy()
                    
                    if len(repeat_samples) > 0:
                        # Remove old repeat samples with same S.no
                        existing_snos = repeat_samples["S.no"].astype(str).str.strip().tolist()
                        df_master = df_master[~df_master["S.no"].astype(str).str.strip().isin(existing_snos)]
                        
                        # Add new repeat samples
                        df_master = pd.concat([df_master, repeat_samples], ignore_index=True)
                        df_master = calculate_metrics_and_alerts(df_master)
                        save_all_sheets(df_master)
                        st.success(f"✅ Only Repeat Sheet - {len(repeat_samples)} samples added/updated!")
                    else:
                        st.warning("Is file mein koi Repeat sample nahi mila.")

                st.rerun()
                
        except Exception as e:
            st.error(f"Error: {str(e)}")

st.markdown("---")

# Dynamic View (Same as before)
st.markdown(f"### Current View: **{st.session_state.current_view}**")

if st.session_state.current_view == "Overall":
    st.subheader("📋 All Active Samples")
    if len(df_running) > 0:
        st.data_editor(df_running, use_container_width=True, num_rows="dynamic", key="overall")
    else:
        st.info("No active samples")

elif st.session_state.current_view == "Development":
    st.subheader("🧪 Development Section")
    if len(df_dev) > 0:
        st.data_editor(df_dev, use_container_width=True, num_rows="dynamic", key="dev")
    else:
        st.info("Development section khali hai")

elif st.session_state.current_view == "Repeat":
    st.subheader("🔄 Repeat Section")
    if len(df_repeat) > 0:
        st.data_editor(df_repeat, use_container_width=True, num_rows="dynamic", key="repeat_editor")
        
        st.markdown("### 🔄 Transfer to Development")
        df_repeat_display = df_repeat.copy()
        df_repeat_display["Display"] = df_repeat_display["Ref"] + "  |  S.no: " + df_repeat_display["S.no"].astype(str)
        
        selected_display = st.multiselect(
            "Transfer karne ke liye samples select karein:",
            options=df_repeat_display["Display"].tolist(),
            default=[],
        )
        
        if st.button("🔄 Selected Samples ko Development mein Transfer Karein", type="primary", use_container_width=True):
            if selected_display:
                selected_snos = [item.split("S.no: ")[-1].strip() for item in selected_display]
                mask = df_master["S.no"].astype(str).str.strip().isin(selected_snos)
                df_master.loc[mask, "Source"] = "Scratch"
                save_all_sheets(df_master)
                st.success(f"✅ {len(selected_snos)} samples Development mein transfer ho gaye!")
                st.rerun()
            else:
                st.warning("Koi sample select nahi kiya!")
    else:
        st.info("Repeat section khali hai.")

# Archive
st.subheader("📁 Submitted to Marketing (Archive)")
st.dataframe(df_master[df_master["Status"] == "Submitted to marketing"], use_container_width=True)
