import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import io

# ==========================================
# 1. Page Configuration
# ==========================================
st.set_page_config(
    page_title="T&C Resource Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 2. Dummy Data Generator (For Demo Purposes)
# ==========================================
@st.cache_data
def generate_dummy_data():
    """Generates synthetic data if no Excel file is uploaded."""
    data = [
        {"Activity": "JRL-Correspondence", "Line": "JRL", "Role": "SIG", "Assigned": "Farid", "Week": "W27", "Month": "Jul", "Load": 0.6},
        {"Activity": "JRL-Through Test", "Line": "JRL", "Role": "SIG", "Assigned": "Farid", "Week": "W27", "Month": "Jul", "Load": 0.6},
        {"Activity": "CRL-Dynamic", "Line": "CRL", "Role": "ATC", "Assigned": "Daniel", "Week": "W27", "Month": "Jul", "Load": 0.8},
        {"Activity": "JRL-Test Running", "Line": "JRL", "Role": "ATC", "Assigned": "Marcus", "Week": "W28", "Month": "Jul", "Load": 1.0},
        {"Activity": "TGD-Power Up", "Line": "TGD", "Role": "Comms", "Assigned": "Damuel", "Week": "W28", "Month": "Jul", "Load": 0.5},
        {"Activity": "CRL-Test Running", "Line": "CRL", "Role": "Subcon", "Assigned": "TBC", "Week": "W27", "Month": "Jul", "Load": 1.0},
        {"Activity": "JRL-OSIT", "Line": "JRL", "Role": "Subcon", "Assigned": "TBC", "Week": "W29", "Month": "Jul", "Load": 2.0},
        {"Activity": "CRL-OSIT", "Line": "CRL", "Role": "ATC", "Assigned": "Daniel", "Week": "W28", "Month": "Jul", "Load": 0.4},
        {"Activity": "JRL-Correspondence", "Line": "JRL", "Role": "SIG", "Assigned": "Farid", "Week": "W28", "Month": "Jul", "Load": 0.4},
    ]
    return pd.DataFrame(data)

# ==========================================
# 3. Core Processing Logic
# ==========================================
def process_data(df, view_mode):
    """Aggregates data and calculates clashes based on the selected view mode."""
    
    # Define period column and threshold based on toggle
    period_col = "Week" if view_mode == "Weekly" else "Month"
    threshold = 1.0 if view_mode == "Weekly" else 4.0

    # Separate assigned personnel from TBC
    tbc_df = df[df["Assigned"] == "TBC"]
    assigned_df = df[df["Assigned"] != "TBC"]

    # Aggregate Load by Person and Period
    grouped = assigned_df.groupby(["Assigned", period_col, "Role"])["Load"].sum().reset_index()

    # Determine Status
    def get_status(load, thresh):
        if load > thresh: return "Clash"
        if load == thresh: return "Fully Loaded"
        return "Available"

    grouped["Status"] = grouped["Load"].apply(lambda x: get_status(x, threshold))
    grouped["Capacity Threshold"] = threshold
    grouped["Excess Load"] = grouped["Load"].apply(lambda x: round(max(0, x - threshold), 2))
    grouped["Available Capacity"] = grouped["Load"].apply(lambda x: round(max(0, threshold - x), 2))

    return df, grouped, tbc_df, period_col, threshold

# ==========================================
# 4. Sidebar Controls
# ==========================================
st.sidebar.title("Controls & Filters")

# The crucial toggle requested for Week/Month view
view_mode = st.sidebar.radio(
    "Select View Mode",
    ["Weekly", "Monthly"],
    help="Weekly threshold is 1.0. Monthly threshold is 4.0."
)

st.sidebar.divider()
uploaded_file = st.sidebar.file_uploader("Upload Manpower Plan (Excel)", type=["xlsx", "xls"])

# Load data (uploaded or dummy)
if uploaded_file is not None:
    raw_df = pd.read_excel(uploaded_file)
else:
    raw_df = generate_dummy_data()
    st.sidebar.info("Currently displaying synthetic demo data. Upload an Excel file to see your data.")

# ==========================================
# 5. Dashboard Layout & KPIs
# ==========================================
st.title("T&C Resource Planning & Allocation Dashboard")

# Process the data
df, grouped_df, tbc_df, period_col, threshold = process_data(raw_df, view_mode)

# Calculate KPIs
total_activities = len(df)
total_load = round(df["Load"].sum(), 1)
clash_count = len(grouped_df[grouped_df["Status"] == "Clash"])
tbc_count = len(tbc_df)

# Display KPIs
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total Activities", total_activities)
col2.metric("Total Manpower Demand", f"{total_load} Load")
col3.metric("Overloaded Personnel", clash_count, delta="Clashes", delta_color="inverse")
col4.metric("TBC / Missing Roles", tbc_count, delta="Gaps", delta_color="inverse")

st.divider()

# ==========================================
# 6. Charts & Visualizations
# ==========================================
st.subheader(f"Resource Loading Overview ({view_mode})")

chart_col1, chart_col2 = st.columns(2)

with chart_col1:
    # Chart 1: Total Demand by Period
    demand_by_period = df.groupby(period_col)["Load"].sum().reset_index()
    fig1 = px.bar(
        demand_by_period, 
        x=period_col, 
        y="Load", 
        title="Total Manpower Demand by Period",
        text_auto=True,
        color_discrete_sequence=["#1f77b4"]
    )
    fig1.update_layout(xaxis_title=period_col, yaxis_title="Total Load")
    st.plotly_chart(fig1, use_container_width=True)

with chart_col2:
    # Chart 2: Demand by Discipline (Role)
    demand_by_role = df.groupby([period_col, "Role"])["Load"].sum().reset_index()
    fig2 = px.bar(
        demand_by_role, 
        x=period_col, 
        y="Load", 
        color="Role", 
        title="Manpower Demand by Discipline",
        barmode="stack"
    )
    fig2.update_layout(xaxis_title=period_col, yaxis_title="Total Load")
    st.plotly_chart(fig2, use_container_width=True)

st.divider()

# ==========================================
# 7. Actionable Reports (Clashes & Gaps)
# ==========================================
st.subheader("Actionable Insights")

tab1, tab2, tab3 = st.tabs(["🔴 Clash Report", "⚪ TBC Gaps", "🟢 Available Capacity"])

with tab1:
    clashes = grouped_df[grouped_df["Status"] == "Clash"].drop(columns=["Available Capacity"])
    if not clashes.empty:
        st.warning(f"Found {len(clashes)} instance(s) where personnel exceed the {threshold} load threshold.")
        st.dataframe(clashes.style.highlight_max(subset=['Excess Load'], color='#ff4b4b'), use_container_width=True)
    else:
        st.success("No manpower clashes found in this view!")

with tab2:
    if not tbc_df.empty:
        st.info(f"There are {len(tbc_df)} unassigned (TBC) activities requiring resource allocation.")
        st.dataframe(tbc_df, use_container_width=True)
    else:
        st.success("All roles are currently assigned.")

with tab3:
    available = grouped_df[grouped_df["Status"] == "Available"].drop(columns=["Excess Load", "Capacity Threshold"])
    if not available.empty:
        st.write("Personnel with capacity remaining in their period:")
        st.dataframe(available, use_container_width=True)
    else:
        st.write("No available capacity found.")

# ==========================================
# 8. Export Functionality
# ==========================================
st.sidebar.divider()
st.sidebar.subheader("Export Reports")

# Convert the grouped processed data back to Excel in memory
output = io.BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    grouped_df.to_excel(writer, index=False, sheet_name='Manpower Loading')
    tbc_df.to_excel(writer, index=False, sheet_name='TBC Gaps')
    df.to_excel(writer, index=False, sheet_name='Raw Data')

processed_data = output.getvalue()

st.sidebar.download_button(
    label="📥 Download Updated Plan (Excel)",
    data=processed_data,
    file_name="Resource_Allocation_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
