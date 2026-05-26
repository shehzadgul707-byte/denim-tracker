import streamlit as st
import pandas as pd
from datetime import date
import os
import plotly.express as px

st.set_page_config(page_title="Denim R&D Advanced Tracker", layout="wide")

EXCEL_FILE = "Denim_Master_Database.xlsx"

STATUS_OPTIONS = ["Yarn Demand", "Dyeing", "Weaving", "Finishing", "Inspection", "Submitted to marketing"]

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
            if "Category" not in df.columns:
                df["Category"] = ""
            return df
        except:
            df = pd.DataFrame(columns=["S.no","Brand","Ref","Finish","TR Code","Source",
                                       "Pending Days in Process","Request Date","Current Date",
                                       "Delivery date","Ready date","Status","Alert","Warp",
                                       "Warp Slub","Weft","Shade","Weave","Reed","Picks",
                                       "Greige Meters","Remarks","Category"])
            return df
    else:
        df = pd.DataFrame(columns=["S.no","Brand","Ref","Finish","TR Code","Source",
                                   "Pending Days in Process","Request Date","Current Date",
                                   "Delivery date","Ready date","Status","Alert","Warp",
                                   "Warp Slub","Weft","Shade","Weave","Reed","Picks",
                                   "Greige Meters","Remarks","Category"])
        return df

df_master = load_data()

# ====================== UPLOAD ======================
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

        for col in df_master.columns:
            if col not in imported_df.columns:
                imported_df[col] = ""

        final_df = imported_df[df_master.columns].copy()
        final_df = calculate_metrics_and_alerts(final_df)

        if upload_type == "Full Pipeline File":
            save_all_sheets(final_df)
            st.success(f"Full Pipeline Loaded: {len(final_df)} samples")
        else:
            if "Category" in final_df.columns:
                final_df["Category"] = final_df["Category"].str.strip()
            else:
                final_df["Category"] = "Repeat"

            new_repeat = final_df[final_df["Category"].str.contains("repeat", case=False, na=False)].copy()
            
            if len(new_repeat) > 0:
                snos_to_update = new_repeat["S.no"].astype(str).str.strip().tolist()
                df_master = df_master[~(
                    df_master["S.no"].astype(str).str.strip().isin(snos_to_update) & 
                    df_master["Category"].str.contains("repeat", case=False, na=False)
                )]
                df_master = pd.concat([df_master, new_repeat], ignore_index=True)
                df_master = calculate_metrics_and_alerts(df_master)
                save_all_sheets(df_master)
                st.success(f"✅ {len(new_repeat)} Repeat samples updated!")
            else:
                st.warning("Koi Repeat sample nahi mila.")

        st.rerun()

    except Exception as e:
        st.error(f"Error: {str(e)}")

st.markdown("---")

# ====================== GRAPHS ======================
df_running = df_master[df_master["Status"] != "Submitted to marketing"].copy() if len(df_master) > 0 else pd.DataFrame()
df_dev = df_running[df_running.get("Category", pd.Series("")).str.contains("development", case=False, na=False)].copy()
df_repeat = df_running[df_running.get("Category", pd.Series("")).str.contains("repeat", case=False, na=False)].copy()

col1, col2, col3 = st.columns([1, 2, 2])
with col1:
    st.metric("**Total On Floor**", len(df_running))
with col2:
    fig_pie = px.pie(names=['Development', 'Repeat'], values=[len(df_dev), len(df_repeat)], 
                     title="Development vs Repeat", color_discrete_sequence=['#0D9488', '#B45309'])
    st.plotly_chart(fig_pie, use_container_width=True)
with col3:
    fig_bar = px.bar(x=['Development', 'Repeat'], y=[len(df_dev), len(df_repeat)], 
                     text=[len(df_dev), len(df_repeat)], title="Bar Chart",
                     color_discrete_sequence=['#0D9488', '#B45309'])
    st.plotly_chart(fig_bar, use_container_width=True)

st.markdown("---")

# ====================== EDIT SAMPLE (Fixed) ======================
st.subheader("✏️ Edit Any Sample")

all_active = df_running.copy()
if len(all_active) > 0:
    all_active["Display"] = all_active["S.no"].astype(str) + " - " + all_active["Ref"].astype(str)
    
    selected_display = st.selectbox("Select Sample to Edit", options=all_active["Display"].tolist())
    
    if selected_display:
        selected_sno = selected_display.split(" - ")[0].strip()
        sample_row = df_master[df_master["S.no"].astype(str) == selected_sno]
        
        if not sample_row.empty:
            sample = sample_row.iloc[0]
            
            with st.form("edit_sample_form"):
                st.write(f"**Editing Sample: S.no {selected_sno}**")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    brand = st.text_input("Brand", value=sample.get("Brand", ""))
                    ref = st.text_input("Ref", value=sample.get("Ref", ""))
                    finish = st.text_input("Finish", value=sample.get("Finish", ""))
                    tr_code = st.text_input("TR Code", value=sample.get("TR Code", ""))
                
                with col2:
                    source = st.selectbox("Source", ["Existing", "Scratch", "Production beam"], 
                                        index=1)
                    category = st.selectbox("Category", ["Development", "Repeat"], 
                                          index=0 if str(sample.get("Category", "")).lower() == "development" else 1)
                    status = st.selectbox("Status", STATUS_OPTIONS, 
                                        index=STATUS_OPTIONS.index(sample.get("Status", "Yarn Demand")) 
                                        if sample.get("Status") in STATUS_OPTIONS else 0)
                
                with col3:
                    # Safe Date Handling
                    try:
                        req_default = pd.to_datetime(sample.get("Request Date")).date()
                    except:
                        req_default = date.today()
                    
                    try:
                        del_default = pd.to_datetime(sample.get("Delivery date")).date()
                    except:
                        del_default = date.today()
                    
                    req_date = st.date_input("Request Date", value=req_default)
                    del_date = st.date_input("Delivery Date", value=del_default)
                    remarks = st.text_area("Remarks", value=sample.get("Remarks", ""))
                
                if st.form_submit_button("💾 Save Changes", use_container_width=True):
                    mask = df_master["S.no"].astype(str) == selected_sno
                    df_master.loc[mask, "Brand"] = brand
                    df_master.loc[mask, "Ref"] = ref
                    df_master.loc[mask, "Finish"] = finish
                    df_master.loc[mask, "TR Code"] = tr_code
                    df_master.loc[mask, "Source"] = source
                    df_master.loc[mask, "Category"] = category
                    df_master.loc[mask, "Status"] = status
                    df_master.loc[mask, "Request Date"] = str(req_date)
                    df_master.loc[mask, "Delivery date"] = str(del_date)
                    df_master.loc[mask, "Remarks"] = remarks
                    
                    df_master = calculate_metrics_and_alerts(df_master)
                    save_all_sheets(df_master)
                    st.success("✅ Sample Updated Successfully!")
                    st.rerun()
else:
    st.info("No active samples to edit.")

# ====================== TABS ======================
view = st.radio("View Section", ["Overall", "Development", "Repeat"], horizontal=True)

if view == "Overall":
    st.subheader("📋 All Active Samples")
    if len(df_running) > 0:
        st.data_editor(df_running, use_container_width=True, num_rows="dynamic", key="overall")

elif view == "Development":
    st.subheader("🧪 Development Section")
    if len(df_dev) > 0:
        st.data_editor(df_dev, use_container_width=True, num_rows="dynamic", key="dev")

elif view == "Repeat":
    st.subheader("🔄 Repeat Section")
    if len(df_repeat) > 0:
        st.data_editor(df_repeat, use_container_width=True, num_rows="dynamic", key="repeat")

st.subheader("📁 Archive (Submitted to Marketing)")
st.dataframe(df_master[df_master["Status"] == "Submitted to marketing"], use_container_width=True)
