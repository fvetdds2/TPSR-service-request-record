import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(
    page_title="TPSR Core Service Analytics",
    layout="wide"
)

# --------------------------------------------------
# ENTERPRISE STYLING (Mobile + Elderly Friendly)
# --------------------------------------------------
st.markdown("""
<style>
html, body, [class*="css"]  {
    font-size: 18px !important;
}

h1 {
    color: #1F3A8A;
    font-weight: 700;
}

h2, h3 {
    color: #1F3A8A;
    font-weight: 600;
}

.section-box {
    background-color: #F3F4F6;
    padding: 20px;
    border-radius: 12px;
    margin-bottom: 25px;
}

</style>
""", unsafe_allow_html=True)

# --------------------------------------------------
# PASSCODE (Use Streamlit secrets in production)
# --------------------------------------------------
PASSCODE = "TPSR2025"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 TPSR Core Service Analytics")
    entered = st.text_input("Enter Access Code", type="password")
    if st.button("Unlock"):
        if entered == PASSCODE:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect access code.")
    st.stop()

# --------------------------------------------------
# LOAD DATA
# --------------------------------------------------
@st.cache_data
def load_data():
    df = pd.read_excel(
        "cost_recovery_record_from_2025.xlsx",
        engine="openpyxl"
    )

    df["Required_Date"] = pd.to_datetime(df["Required_Date"], errors="coerce")
    df["Month_Year"] = df["Required_Date"].dt.to_period("M")
    df["Month_Label"] = df["Required_Date"].dt.strftime("%b %Y")

    df["Cancer_Related_Project"] = (
        df["Cancer_Related_Project"]
        .astype(str)
        .str.strip()
        .str.capitalize()
    )

    return df

df = load_data()

# --------------------------------------------------
# SIDEBAR FILTERS
# --------------------------------------------------
st.sidebar.header("Filters")

status_filter = st.sidebar.multiselect(
    "Select Status",
    df["Status"].dropna().unique(),
    default=df["Status"].dropna().unique()
)

requester_filter = st.sidebar.multiselect(
    "Select Requester",
    df["Requester_Name"].dropna().unique(),
    default=df["Requester_Name"].dropna().unique()
)

df_filtered = df[
    df["Status"].isin(status_filter) &
    df["Requester_Name"].isin(requester_filter)
]

# --------------------------------------------------
# HEADER
# --------------------------------------------------
st.title("Translational Pathology Shared Resource")
st.subheader("Core Service Request Activity & Cost Recovery Overview")
st.divider()

# --------------------------------------------------
# STATUS DISTRIBUTION
# --------------------------------------------------
st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.subheader("Service Request Status Distribution")

status_counts = (
    df_filtered["Status"]
    .value_counts()
    .reset_index()
)

status_counts.columns = ["Status", "Number of Requests"]

fig_status = px.bar(
    status_counts,
    x="Status",
    y="Number of Requests",
    color="Status",
    text="Number of Requests",
    color_discrete_map={
        "Completed": "#16A34A",
        "Pending": "#F59E0B",
        "Unknown": "#9CA3AF"
    }
)

fig_status.update_traces(textposition="outside")
fig_status.update_layout(
    xaxis_title="Request Status",
    yaxis_title="Number of Requests",
    showlegend=False,
    font=dict(size=18)
)

st.plotly_chart(fig_status, use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

# --------------------------------------------------
# REQUESTS BY INVESTIGATOR
# --------------------------------------------------
st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.subheader("Requests by Investigator")

req_counts = (
    df_filtered
    .groupby("Requester_Name")
    .size()
    .reset_index(name="Number of Requests")
    .sort_values("Number of Requests", ascending=False)
)

fig_req = px.bar(
    req_counts,
    x="Requester_Name",
    y="Number of Requests",
    text="Number of Requests",
    color_discrete_sequence=["#1F3A8A"]
)

fig_req.update_traces(textposition="outside")
fig_req.update_layout(
    xaxis_title="Investigator",
    yaxis_title="Number of Requests",
    showlegend=False,
    font=dict(size=18)
)

st.plotly_chart(fig_req, use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

# --------------------------------------------------
# SERVICE TYPE DISTRIBUTION
# --------------------------------------------------
service_cols = [
    "FFPE processing & Embedding",
    "FFPE sectioning & H&E stain",
    "Frozen sectioning-unstained slide",
    "Frozen sectioning & H&E stain",
    "Frozen sectioning-step section",
    "Repository FFPE sectioning-unstained slide",
    "histology tissue collection vials",
    "histopathology support (hr)"
]

st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.subheader("Service Type Utilization")

svc_totals = (
    df_filtered[service_cols]
    .sum()
    .reset_index()
)

svc_totals.columns = ["Service Type", "Total Units"]
svc_totals = svc_totals.sort_values("Total Units", ascending=False)

fig_svc = px.bar(
    svc_totals,
    x="Service Type",
    y="Total Units",
    text="Total Units",
    color_discrete_sequence=["#0F766E"]
)

fig_svc.update_traces(textposition="outside")
fig_svc.update_layout(
    xaxis_title="Service Type",
    yaxis_title="Total Units",
    showlegend=False,
    font=dict(size=18)
)

st.plotly_chart(fig_svc, use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

# --------------------------------------------------
# MONTHLY SERVICE VOLUME
# --------------------------------------------------
st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.subheader("Monthly Service Volume")

df_filtered["Total_Service_Units"] = df_filtered[service_cols].sum(axis=1)

monthly_services = (
    df_filtered
    .groupby(["Month_Year", "Month_Label"])["Total_Service_Units"]
    .sum()
    .reset_index()
    .sort_values("Month_Year")
)

fig_month = px.bar(
    monthly_services,
    x="Month_Label",
    y="Total_Service_Units",
    text="Total_Service_Units",
    color_discrete_sequence=["#1F3A8A"]
)

fig_month.update_traces(textposition="outside")
fig_month.update_layout(
    xaxis_title="Month",
    yaxis_title="Total Service Units",
    showlegend=False,
    font=dict(size=18)
)

st.plotly_chart(fig_month, use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

# --------------------------------------------------
# MONTHLY COST RECOVERY
# --------------------------------------------------
st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.subheader("Monthly Cost Recovery")

monthly_cost = (
    df_filtered
    .groupby(["Month_Year", "Month_Label"])["Cost_Recovery"]
    .sum()
    .reset_index()
    .sort_values("Month_Year")
)

fig_cost = px.line(
    monthly_cost,
    x="Month_Label",
    y="Cost_Recovery",
    markers=True,
    color_discrete_sequence=["#0F766E"]
)

fig_cost.update_layout(
    xaxis_title="Month",
    yaxis_title="Cost Recovery ($)",
    font=dict(size=18)
)

st.plotly_chart(fig_cost, use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

# --------------------------------------------------
# CANCER RELATED PROJECT DISTRIBUTION
# --------------------------------------------------
st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.subheader("Cancer-Related Project Distribution")

cancer_counts = (
    df_filtered["Cancer_Related_Project"]
    .value_counts()
    .reset_index()
)

cancer_counts.columns = ["Cancer Related", "Number of Projects"]

fig_cancer = px.bar(
    cancer_counts,
    x="Cancer Related",
    y="Number of Projects",
    text="Number of Projects",
    color_discrete_map={
        "Yes": "#1F3A8A",
        "No": "#0F766E",
        "Unknown": "#9CA3AF"
    }
)

fig_cancer.update_traces(textposition="outside")
fig_cancer.update_layout(
    xaxis_title="Cancer Related Project",
    yaxis_title="Number of Projects",
    showlegend=False,
    font=dict(size=18)
)

st.plotly_chart(fig_cancer, use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

# --------------------------------------------------
# RAW DATA TABLE
# --------------------------------------------------
with st.expander("View Detailed Data Table"):
    st.dataframe(df_filtered, use_container_width=True)
)