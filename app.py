import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="TPSR Analytics Platform",
    layout="wide",
    page_icon="📊"
)

# =========================================================
# DESIGN SYSTEM (Integrity / Accuracy Theme)
# =========================================================
PRIMARY = "#1e3a8a"     # Deep institutional blue
SUCCESS = "#059669"
WARNING = "#d97706"
DANGER  = "#b91c1c"
GRAY    = "#6b7280"
LIGHT_BG = "#f8fafc"

STATUS_COLORS = {
    "Completed": SUCCESS,
    "Pending": WARNING,
    "Unknown": GRAY,
}

# =========================================================
# GLOBAL STYLING
# =========================================================
st.markdown(f"""
<style>
[data-testid="stAppViewContainer"] {{
    background-color: {LIGHT_BG};
}}

.block-container {{
    padding-top: 2rem;
}}

.metric-card {{
    background: white;
    padding: 24px;
    border-radius: 14px;
    box-shadow: 0 4px 14px rgba(0,0,0,0.04);
}}

.metric-title {{
    font-size: 0.85rem;
    color: {GRAY};
}}

.metric-value {{
    font-size: 1.9rem;
    font-weight: 700;
    color: {PRIMARY};
    margin-top: 8px;
}}

.section-title {{
    font-size: 1.3rem;
    font-weight: 600;
    margin-top: 30px;
    margin-bottom: 15px;
    color: {PRIMARY};
}}
</style>
""", unsafe_allow_html=True)

# =========================================================
# AUTHENTICATION (Replace with OAuth in Production)
# =========================================================
PASSCODE = "TPSR2025"

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("Secure Access")
    code = st.text_input("Enter Access Code", type="password")
    if st.button("Access Platform"):
        if code == PASSCODE:
            st.session_state.auth = True
            st.rerun()
        else:
            st.error("Invalid code.")
    st.stop()

# =========================================================
# DATA LOADING
# =========================================================
@st.cache_data
def load_data():
    df = pd.read_excel("cost_recovery_record_from_2025.xlsx")

    df["Required_Date"] = pd.to_datetime(df["Required_Date"], errors="coerce")
    df["Month"] = df["Required_Date"].dt.to_period("M")
    df["Month_Label"] = df["Required_Date"].dt.strftime("%b %Y")

    df["Status"] = df["Status"].fillna("Unknown")
    df["Cost_Recovery"] = pd.to_numeric(df["Cost_Recovery"], errors="coerce").fillna(0)

    service_cols = df.columns[4:12].tolist()

    df["Cancer_Related_Project"] = (
        df["Cancer_Related_Project"]
        .fillna("Unknown")
        .astype(str)
        .str.capitalize()
    )

    return df, service_cols

df, service_cols = load_data()

# =========================================================
# SIDEBAR FILTERS
# =========================================================
st.sidebar.header("Filters")

status_filter = st.sidebar.multiselect(
    "Status",
    options=df["Status"].unique(),
    default=df["Status"].unique()
)

requester_filter = st.sidebar.multiselect(
    "Requester",
    options=sorted(df["Requester_Name"].dropna().unique()),
    default=sorted(df["Requester_Name"].dropna().unique())
)

df = df[
    df["Status"].isin(status_filter) &
    df["Requester_Name"].isin(requester_filter)
]

# =========================================================
# HEADER
# =========================================================
st.title("MMC Translational Pathology Shared Resource")
st.caption("Commercial Analytics Platform • 2025–2026")

# =========================================================
# KPI STRIP
# =========================================================
completed = int((df["Status"] == "Completed").sum())
pending = int((df["Status"] == "Pending").sum())
total_cost = df["Cost_Recovery"].sum()
total_units = int(df[service_cols].sum().sum())

def metric_card(title, value):
    st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">{title}</div>
            <div class="metric-value">{value}</div>
        </div>
    """, unsafe_allow_html=True)

c1, c2, c3, c4 = st.columns(4)
with c1:
    metric_card("Completed Requests", completed)
with c2:
    metric_card("Pending Requests", pending)
with c3:
    metric_card("Total Cost Recovery", f"${total_cost:,.0f}")
with c4:
    metric_card("Total Service Units", total_units)

# =========================================================
# STATUS & CANCER DISTRIBUTION
# =========================================================
st.markdown('<div class="section-title">Operational Overview</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    status_counts = df["Status"].value_counts().reset_index()
    status_counts.columns = ["Status", "Count"]

    fig_status = px.pie(
        status_counts,
        names="Status",
        values="Count",
        hole=0.5,
        color="Status",
        color_discrete_map=STATUS_COLORS,
    )
    fig_status.update_layout(template="plotly_white")
    st.plotly_chart(fig_status, use_container_width=True)

with col2:
    cancer_counts = df["Cancer_Related_Project"].value_counts().reset_index()
    cancer_counts.columns = ["Cancer", "Count"]

    fig_cancer = px.pie(
        cancer_counts,
        names="Cancer",
        values="Count",
        hole=0.5,
        color_discrete_sequence=[PRIMARY, GRAY, DANGER],
    )
    fig_cancer.update_layout(template="plotly_white")
    st.plotly_chart(fig_cancer, use_container_width=True)

# =========================================================
# REQUESTER PERFORMANCE
# =========================================================
st.markdown('<div class="section-title">Requester Performance</div>', unsafe_allow_html=True)

req = (
    df.groupby("Requester_Name")
    .agg(Requests=("Status", "count"),
         Revenue=("Cost_Recovery", "sum"))
    .reset_index()
    .sort_values("Requests", ascending=False)
)

fig_req = px.bar(
    req,
    x="Requester_Name",
    y="Requests",
    color_discrete_sequence=[PRIMARY],
)
fig_req.update_layout(template="plotly_white", xaxis_tickangle=-30)
st.plotly_chart(fig_req, use_container_width=True)

# =========================================================
# MONTHLY SERVICE TREND
# =========================================================
st.markdown('<div class="section-title">Monthly Service Trend</div>', unsafe_allow_html=True)

df["Total_Units"] = df[service_cols].sum(axis=1)

monthly = (
    df.groupby(["Month", "Month_Label"])
    ["Total_Units"]
    .sum()
    .reset_index()
    .sort_values("Month")
)

fig_month = px.area(
    monthly,
    x="Month_Label",
    y="Total_Units",
    color_discrete_sequence=[PRIMARY]
)

fig_month.update_layout(
    template="plotly_white",
    hovermode="x unified"
)

st.plotly_chart(fig_month, use_container_width=True)

# =========================================================
# COST TREND
# =========================================================
st.markdown('<div class="section-title">Revenue Trend</div>', unsafe_allow_html=True)

cost_trend = (
    df.groupby(["Month", "Month_Label"])
    ["Cost_Recovery"]
    .sum()
    .reset_index()
    .sort_values("Month")
)

fig_cost = px.line(
    cost_trend,
    x="Month_Label",
    y="Cost_Recovery",
    markers=True,
    color_discrete_sequence=[PRIMARY]
)

fig_cost.update_layout(
    template="plotly_white",
    yaxis_tickprefix="$",
    hovermode="x unified"
)

st.plotly_chart(fig_cost, use_container_width=True)

# =========================================================
# DATA EXPORT
# =========================================================
st.markdown('<div class="section-title">Data Access</div>', unsafe_allow_html=True)

csv = df.to_csv(index=False).encode("utf-8")

st.download_button(
    "Download Filtered Dataset",
    csv,
    "TPSR_filtered_data.csv",
    "text/csv"
)