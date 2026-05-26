import streamlit as st
import pandas as pd
from datetime import date
import os

st.set_page_config(page_title="Denim R&D Advanced Tracker", layout="wide")

EXCEL_FILE = "Denim_Master_Database.xlsx"

if "fresh_uploaded" not in st.session_state:
    st.session_state.fresh_uploaded = False

if "current_view" not in st.session_state:
    st.session_state.current_view = "Overall"

# ====================== HELPERS ======================
def read_excel_as_string(file):
    return pd.read_excel(file, dtype=str)

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

def load_data():
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name="Master_Data", dtype=str)
            df = clean_dataframe(df)
            # Agar Category column nahi hai to add kar do
            if "Category" not in df.columns:
                df["Category"] = ""
            return df
        except:
            df = pd.DataFrame()
            df["Category"] = ""
            return df
    else:
        df = pd.DataFrame()
        df["Category"] = ""
        return df

df_master = load_data()

# ====================== UI ======================
st.title("👖 Denim Fabric R&D Production Pipeline Tracker")
st.markdown("---")

upload_type = st.radio("Upload Type", 
    ["Full Pipeline File", "Only Repeat Sampling Sheet"], horizontal=True)

uploaded_file = st.file_uploader("Excel File Select Karein", type=["xlsx"])

if uploaded_file and st.button("🚀 Upload & Process", type="primary", use_container_width=True):
    try:
        imported_df = read_excel_as_string(uploaded_file)
        imported_df = clean_dataframe(imported_df)

        if "Marketing" in imported_df.columns:
            imported_df.rename(columns={"Marketing": "Ref"}, inplace=True)

        if "Ref" not in imported_df.columns:
            st.error("Ref ya Marketing column nahi mila")
            st.stop()

        if "Ppi" in imported_df.columns and "Picks" not in imported_df.columns:
            imported_df.rename(columns={"Ppi": "Picks"}, inplace=True)

        # Add missing columns including Category
        for col in list(df_master.columns) + ["Category"]:
            if col not in imported_df.columns:
                imported_df[col] = ""

        final_df = imported_df[df_master.columns].copy() if len(df_master) > 0 else imported_df
        final_df = calculate_metrics_and_alerts(final_df)

        if upload_type == "Full Pipeline File":
            save_all_sheets(final_df)
            st.success(f"Full Pipeline Loaded: {len(final_df)} samples")
        else:
            existing_snos = final_df["S.no"].astype(str).str.strip().tolist()
            df_master = df_master[~df_master["S.no"].astype(str).str.strip().isin(existing_snos)]
            df_master = pd.concat([df_master, final_df], ignore_index=True)
            df_master = calculate_metrics_and_alerts(df_master)
            save_all_sheets(df_master)
            st.success(f"✅ Repeat Sheet Loaded: {len(final_df)} samples")
        
        st.rerun()

    except Exception as e:
        st.error(f"Error: {str(e)}")

st.markdown("---")

# ====================== CLASSIFICATION ======================
df_running = df_master[df_master["Status"] != "Submitted to marketing"].copy() if len(df_master) > 0 else pd.DataFrame()

# Safe Category Check
if "Category" in df_running.columns:
    df_dev = df_running[df_running["Category"].str.contains("development", case=False, na=False)].copy()
    df_repeat = df_running[df_running["Category"].str.contains("repeat", case=False, na=False)].copy()
else:
    df_dev = pd.DataFrame()
    df_repeat = df_running.copy()

# KPI
c1, c2, c3 = st.columns(3)
with c1: st.metric("Total On Floor", len(df_running))
with c2: st.metric("Development", len(df_dev))
with c3: st.metric("Repeat", len(df_repeat))

st.info("**Classification:** 'Category' column ke basis par (Development / Repeat)")

# Views with Delete Option
view = st.radio("Section", ["Overall", "Development", "Repeat"], horizontal=True, key="view_selector")

if view == "Overall":
    st.subheader("📋 All Active Samples")
    if len(df_running) > 0:
        edited = st.data_editor(df_running, use_container_width=True, num_rows="dynamic", key="overall")
        if len(edited) < len(df_running):
            remaining = edited["S.no"].astype(str).str.strip().tolist()
            df_master = df_master[df_master["S.no"].astype(str).str.strip().isin(remaining) | 
                                (df_master["Status"] == "Submitted to marketing")]
            save_all_sheets(df_master)
            st.rerun()

elif view == "Development":
    st.subheader("🧪 Development Section")
    if len(df_dev) > 0:
        st.data_editor(df_dev, use_container_width=True, num_rows="dynamic", key="dev")

elif view == "Repeat":
    st.subheader("🔄 Repeat Section")
    if len(df_repeat) > 0:
        st.data_editor(df_repeat, use_container_width=True, num_rows="dynamic", key="repeat")

# Archive
st.subheader("📁 Archive (Submitted to Marketing)")
st.dataframe(df_master[df_master["Status"] == "Submitted to marketing"], use_container_width=True)
