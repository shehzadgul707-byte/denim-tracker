import streamlit as st
import pandas as pd
import plotly.express as px
import os

# Set page layout to wide for dashboard look
st.set_page_config(page_title="Denim R&D Sample Tracker", layout="wide")

EXCEL_FILE = "Sample_Master_DB.xlsx"
STAGES = ["1. Yarn", "2. Dyeing", "3. Weaving", "4. Finishing", "5. Inspection", "6. Washing", "7. Fabric Submitted to Marketing"]
SAMPLE_TYPES = ["New Development", "Repeat Sampling"]
SOURCES = ["Marketing Request", "Counter"]

# --- DATA FUNCTIONS ---
def load_data():
    """Loads data from Excel, or creates a template if it doesn't exist."""
    if os.path.exists(EXCEL_FILE):
        return pd.read_excel(EXCEL_FILE)
    else:
        # Create an empty template with correct columns
        df = pd.DataFrame(columns=[
            "Sample ID", "Sample Type", "Request Source", 
            "Brand / Customer", "Current Stage", "Target Date"
        ])
        df.to_excel(EXCEL_FILE, index=False)
        return df

def save_data(df):
    """Saves the current dataframe back to the Excel file."""
    df.to_excel(EXCEL_FILE, index=False)

# Load data into session memory
df = load_data()

# --- APP HEADER ---
st.title("👖 Denim R&D Sample Status Dashboard")
st.markdown("Track development routes from yarn processing to final marketing handover.")
st.divider()

# --- SIDEBAR: ADD / UPDATE SAMPLES ---
st.sidebar.header("🕹️ R&D Control Panel")
action = st.sidebar.radio("Choose Action:", ["View Dashboard", "Add New Sample", "Update Existing Sample Stage"])

if action == "Add New Sample":
    st.sidebar.subheader("✨ Register New Sample")
    new_id = st.sidebar.text_input("Sample ID / Quality No:")
    new_type = st.sidebar.selectbox("Sample Type:", SAMPLE_TYPES)
    new_source = st.sidebar.selectbox("Request Source:", SOURCES)
    new_brand = st.sidebar.text_input("Brand / Customer Name:")
    new_stage = st.sidebar.selectbox("Initial Stage:", STAGES)
    new_date = st.sidebar.date_input("Target Delivery Date:")
    
    if st.sidebar.button("Save New Sample to Excel"):
        if new_id and new_brand:
            if new_id in df["Sample ID"].values:
                st.sidebar.error(f"Error: Sample ID '{new_id}' already exists!")
            else:
                new_row = {
                    "Sample ID": new_id, "Sample Type": new_type, 
                    "Request Source": new_source, "Brand / Customer": new_brand, 
                    "Current Stage": new_stage, "Target Date": str(new_date)
                }
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                save_data(df)
                st.sidebar.success(f"🎉 {new_id} saved successfully!")
                st.rerun()
        else:
            st.sidebar.warning("Please fill out Sample ID and Brand fields.")

elif action == "Update Existing Sample Stage":
    st.sidebar.subheader("🔄 Update Sample Route Stage")
    if len(df) == 0:
        st.sidebar.warning("No samples available to update.")
    else:
        select_id = st.sidebar.selectbox("Select Sample ID to Update:", df["Sample ID"].tolist())
        current_row = df[df["Sample ID"] == select_id].iloc[0]
        
        st.sidebar.info(f"Current Stage: {current_row['Current Stage']}")
        next_stage = st.sidebar.selectbox("Move to Stage:", STAGES, index=STAGES.index(current_row['Current Stage']))
        
        if st.sidebar.button("Update Status"):
            df.loc[df["Sample ID"] == select_id, "Current Stage"] = next_stage
            save_data(df)
            st.sidebar.success(f"Moved {select_id} to {next_stage}!")
            st.rerun()

# --- MAIN DASHBOARD INTERFACE ---
if len(df) == 0:
    st.info("👋 Welcome! The Excel database is empty. Use the sidebar on the left to add your first denim sample.")
else:
    # --- TOP ROW: KPI CARDS ---
    total_samples = len(df)
    active_samples = len(df[df["Current Stage"] != STAGES[-1]])
    repeat_samples = len(df[df["Sample Type"] == "Repeat Sampling"])
    marketing_reqs = len(df[df["Request Source"] == "Marketing Request"])
    counters = len(df[df["Request Source"] == "Counter"])
    
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    kpi1.metric("Total Logged", total_samples)
    kpi2.metric("Active in R&D", active_samples, delta_color="inverse")
    kpi3.metric("Repeat Samplings ⚠️", repeat_samples)
    kpi4.metric("Marketing Requests", marketing_reqs)
    kpi5.metric("Counter Samples", counters)
    
    st.divider()
    
    # --- MIDDLE ROW: CHARTS ---
    chart_col1, chart_col2 = st.columns(2)
    
    with chart_col1:
        st.subheader("📊 Active Samples across Production Route")
        # Ensure charts show all stages, even if count is 0
        stage_counts = df[df["Current Stage"] != STAGES[-1]]["Current Stage"].value_counts().reindex(STAGES[:-1], fill_value=0).reset_index()
        stage_counts.columns = ["R&D Stage", "Number of Samples"]
        
        fig_route = px.bar(
            stage_counts, x="Number of Samples", y="R&D Stage", 
            orientation='h', color="Number of Samples",
            color_continuous_scale="Blugrn", text_auto=True
        )
        fig_route.update_layout(yaxis={'categoryorder':'trace'}, showlegend=False)
        st.plotly_chart(fig_route, use_container_width=True)
        
    with chart_col2:
        st.subheader("🎯 Sample Distribution by Brand")
        brand_counts = df["Brand / Customer"].value_counts().reset_index()
        brand_counts.columns = ["Brand", "Sample Count"]
        
        fig_brand = px.pie(
            brand_counts, values="Sample Count", names="Brand", 
            hole=0.4, color_discrete_sequence=px.colors.sequential.Deep
        )
        st.plotly_chart(fig_brand, use_container_width=True)
        
    st.divider()
    
    # --- BOTTOM ROW: RAW LIVE DATA TABLE ---
    st.subheader("📋 Master Live Tracking Sheet (Synced with Excel)")
    
    # Simple search bar to filter the sheet
    search_query = st.text_input("🔍 Quick Search by Brand or Sample ID:")
    if search_query:
        filtered_df = df[df["Sample ID"].str.contains(search_query, case=False) | df["Brand / Customer"].str.contains(search_query, case=False)]
        st.dataframe(filtered_df, use_container_width=True)
    else:
        st.dataframe(df, use_container_width=True)
