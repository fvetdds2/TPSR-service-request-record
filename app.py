import streamlit as st
import pandas as pd
import plotly.express as px

# -----------------------------
# Load Data
# -----------------------------
file_path = "cost recovery record from 2025.xlsx"
df = pd.read_excel(file_path)

# Clean column names
df.columns = df.columns.str.strip().str.replace(" ", "_")

# Convert date
df['Required_Date'] = pd.to_datetime(df['Required_Date'], errors='coerce')

# -----------------------------
# Sidebar Filters
# -----------------------------
st.sidebar.header("Filters")

status_filter = st.sidebar.multiselect(
    "Select Status",
    options=df['Status'].dropna().unique(),
    default=df['Status'].dropna().unique()
)

requester_filter = st.sidebar.multiselect(
    "Select Requester",
    options=df['Requester_Name'].dropna().unique(),
    default=df['Requester_Name'].dropna().unique()
)

df_filtered = df[
    (df['Status'].isin(status_filter)) &
    (df['Requester_Name'].isin(requester_filter))
]

# -----------------------------
# Title
# -----------------------------
st.title("Service Dashboard")

# -----------------------------
# Metrics
# -----------------------------
completed = df_filtered[df_filtered['Status'] == 'Completed'].shape[0]
pending = df_filtered[df_filtered['Status'] == 'Pending'].shape[0]
total_cost = df_filtered['Cost_Recovery'].sum()

col1, col2, col3 = st.columns(3)

col1.metric("Completed Services", completed)
col2.metric("Pending Services", pending)
col3.metric("Total Cost Recovery", f"${total_cost:,.2f}")

# -----------------------------
# Cost Recovery Over Time
# -----------------------------
st.subheader("Cost Recovery Over Time")

cost_time = df_filtered.groupby(
    df_filtered['Required_Date'].dt.to_period('M')
)['Cost_Recovery'].sum().reset_index()

cost_time['Required_Date'] = cost_time['Required_Date'].astype(str)

fig1 = px.line(cost_time, x='Required_Date', y='Cost_Recovery', markers=True)
st.plotly_chart(fig1, use_container_width=True)

# -----------------------------
# Requester Distribution
# -----------------------------
st.subheader("Services by Requester")

requester_count = df_filtered['Requester_Name'].value_counts().reset_index()
requester_count.columns = ['Requester_Name', 'Count']

fig2 = px.bar(requester_count, x='Requester_Name', y='Count')
st.plotly_chart(fig2, use_container_width=True)

# -----------------------------
# Services by Type (Columns C–L)
# -----------------------------
st.subheader("Service Types Distribution")

# Adjust column names here if needed
service_columns = df.columns[2:12]  # Columns C–L

service_counts = df_filtered[service_columns].sum().reset_index()
service_counts.columns = ['Service_Type', 'Total']

fig3 = px.bar(service_counts, x='Service_Type', y='Total')
st.plotly_chart(fig3, use_container_width=True)

# -----------------------------
# Status Breakdown Pie
# -----------------------------
st.subheader("Status Breakdown")

status_counts = df_filtered['Status'].value_counts().reset_index()
status_counts.columns = ['Status', 'Count']

fig4 = px.pie(status_counts, names='Status', values='Count')
st.plotly_chart(fig4, use_container_width=True)

# -----------------------------
# Raw Data
# -----------------------------
st.subheader("Raw Data")
st.dataframe(df_filtered)
