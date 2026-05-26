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

# ====================== LOAD DATA ======================
def load_data():
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name="Master_Data", dtype=str)
            df = clean_dataframe(df)
        except:
            df = pd.DataFrame(columns=["S.no", "Brand", "Ref", "Finish", "TR Code", "Source", "Pending Days in Process",
                                       "Request Date", "Current Date", "Delivery date", "Ready date", "Status", "Alert",
                                       "Warp", "Warp Slub", "Weft", "Shade", "Weave", "Reed", "Picks", "Greige Meters", "Remarks"])
    else:
        df = pd.DataFrame(columns=["S.no", "Brand", "Ref", "Finish", "TR Code", "Source", "Pending Days in Process",
                                   "Request Date", "Current Date", "Delivery date", "Ready date", "Status", "Alert",
                                   "Warp", "Warp Slub", "Weft", "Shade", "Weave", "Reed", "Picks", "Greige Meters", "Remarks"])
    return df

df_master = load_data()

# ====================== UPLOAD ======================
st.title("👖 Denim Fabric R&D Production Pipeline Tracker")
st.markdown("---")

upload_type = st.radio("Upload Type", 
    ["Full Pipeline File (Development + Repeat)", "Only Repeat Sampling Sheet"], 
    horizontal=True)

uploaded_file = st.file_uploader("Excel File", type=["xlsx"])

if uploaded_file and st.button("🚀 Upload & Process", type="primary"):
    try:
        imported_df = read_excel_as_string(uploaded_file)
        imported_df = clean_dataframe(imported_df)

        # Repeat File Special Handling
        if upload_type == "Only Repeat Sampling Sheet":
            if "Marketing" in imported_df.columns:
                imported_df.rename(columns={"Marketing": "Ref"}, inplace=True)
                st.success("Marketing → Ref convert ho gaya")

        if "Ref" not in imported_df.columns:
            st.error("Ref ya Marketing column nahi mila")
            st.stop()

        if "Ppi" in imported_df.columns and "Picks" not in imported_df.columns:
            imported_df.rename(columns={"Ppi": "Picks"}, inplace=True)

        # Add missing columns
        for col in df_master.columns:
            if col not in imported_df.columns:
                imported_df[col] = ""

        final_df = imported_df[df_master.columns].copy()
        final_df = calculate_metrics_and_alerts(final_df)

        if upload_type == "Full Pipeline File (Development + Repeat)":
            save_all_sheets(final_df)
            st.success(f"Full Pipeline Loaded: {len(final_df)} rows")
        else:
            # === REPEAT FILE LOGIC ===
            # Source ko flexible banao
            final_df["Source"] = final_df["Source"].str.lower()
            
            repeat_samples = final_df[final_df["Source"].str.contains("existing|production|repeat", na=False)].copy()
            
            st.write("**Debug:** Repeat samples found =", len(repeat_samples))   # ← Yeh line debug ke liye

            if len(repeat_samples) > 0:
                # Remove old entries
                old_snos = repeat_samples["S.no"].astype(str).str.strip().tolist()
                df_master = df_master[~df_master["S.no"].astype(str).str.strip().isin(old_snos)]
                
                # Add new
                df_master = pd.concat([df_master, repeat_samples], ignore_index=True)
                df_master = calculate_metrics_and_alerts(df_master)
                save_all_sheets(df_master)
                st.success(f"✅ {len(repeat_samples)} Repeat samples successfully added!")
            else:
                st.error("Koi Repeat sample nahi mila. Source column mein 'Existing' ya 'Production' hona chahiye.")

        st.rerun()

    except Exception as e:
        st.error(f"Error: {str(e)}")

st.markdown("---")

# ====================== DISPLAY ======================
df_running = df_master[df_master["Status"] != "Submitted to marketing"].copy()
df_dev = df_running[df_running["Source"].str.contains("scratch", case=False, na=False)].copy()
df_repeat = df_running[df_running["Source"].str.contains("existing|production", case=False, na=False)].copy()

st.markdown("### 📊 Live Counters")
c1, c2, c3 = st.columns(3)
with c1: st.metric("Total On Floor", len(df_running))
with c2: st.metric("Development", len(df_dev))
with c3: st.metric("Repeat", len(df_repeat))

# Show Repeat Section Directly for Debugging
st.subheader("🔄 Repeat Section (Current Data)")
if len(df_repeat) > 0:
    st.dataframe(df_repeat, use_container_width=True)
else:
    st.info("Repeat section mein abhi koi data nahi hai.")

st.subheader("📁 Full Master Data (Debug)")
st.dataframe(df_master.head(20), use_container_width=True)
