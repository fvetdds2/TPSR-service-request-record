import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl
from datetime import datetime

# ==========================================================
# CONFIGURATION
# ==========================================================

st.set_page_config(
    page_title="TPSR CoreSight™",
    page_icon="🧬",
    layout="wide",
)

PRIMARY_BLUE = "#1E3A8A"
SUCCESS_GREEN = "#16A34A"
WARNING_AMBER = "#F59E0B"
DANGER_RED = "#DC2626"
NEUTRAL_GRAY = "#6B7280"

# ==========================================================
# GLOBAL STYLING (Silicon Valley Minimal Design)
# ==========================================================

st.markdown("""
<style>
.main {background-color:#F9FAFB;}
.block-container {padding-top:2rem;}

h1 {font-weight:700; letter-spacing:-0.5px;}
h2, h3 {font-weight:600;}

.metric-card {
    background:white;
    padding:1.4rem;
    border-radius:18px;
    box-shadow:0 6px 18px rgba(0,0,0,0.05);
}

.section-card {
    background:white;
    padding:2rem;
    border-radius:22px;
    box-shadow:0 8px 25px rgba(0,0,0,0.05);
    margin-bottom:2rem;
}
</style>
""", unsafe_allow_html=True)

# ==========================================================
# AUTHENTICATION LAYER
# ==========================================================

PASSCODE = "TPSR2025"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("<h1 style='text-align:center'>TPSR CoreSight™</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center;color:gray'>Enterprise Laboratory Intelligence Platform</p>", unsafe_allow_html=True)
    
    entered = st.text_input("Enter Access Key", type="password")
    if st.button("Secure Login"):
        if entered == PASSCODE:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Invalid access key.")
    st.stop()

# ==========================================================
# DATA LOADER
# ==========================================================
@st.cache_data
def load_data():
    wb = openpyxl.load_workbook("cost_recovery_record_from_2025.xlsx")
    ws = wb.active

    header_map = {cell.column: cell.value for cell in ws[1] if cell.value}
    service_cols = [header_map[i] for i in range(5, 13)]

    def safe_float(value):
        """Enterprise-safe numeric conversion."""
        if value is None:
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, datetime):
            return 0.0
        try:
            return float(str(value).strip())
        except (ValueError, TypeError):
            return 0.0

    records = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue

        record = {
            "Requester_Name": row[0] if len(row) > 0 else None,
            "Required_Date": pd.to_datetime(row[1], errors="coerce") if len(row) > 1 else None,
            "Status": row[2] if len(row) > 2 and row[2] else "Unknown",
            "Cost_Recovery": safe_float(row[3]) if len(row) > 3 else 0.0,
            "Cancer_Related_Project": (
                str(row[13]).capitalize()
                if len(row) > 13 and row[13]
                else "Unknown"
            ),
        }

        # SAFE SERVICE COLUMN PARSING
        for i, col in enumerate(service_cols):
            index = 4 + i
            if index < len(row):
                record[col] = safe_float(row[index])
            else:
                record[col] = 0.0

        records.append(record)

    df = pd.DataFrame(records)

    df["Month"] = df["Required_Date"].dt.to_period("M")
    df["Month_Label"] = df["Required_Date"].dt.strftime("%b %Y")

    return df, service_cols

# ==========================================================
# HEADER
# ==========================================================

st.markdown("<h1>TPSR CoreSight™</h1>", unsafe_allow_html=True)
st.markdown("<p style='color:#6B7280'>Translational Pathology Shared Resource Intelligence Platform</p>", unsafe_allow_html=True)
st.divider()

# ==========================================================
# FILTER PANEL
# ==========================================================
st.write("DF exists?", "df" in locals())
st.sidebar.header("Global Filters")

status_filter = st.sidebar.multiselect(
    "Project Status",
    df["Status"].unique(),
    default=df["Status"].unique()
)

requester_filter = st.sidebar.multiselect(
    "Requester",
    df["Requester_Name"].unique(),
    default=df["Requester_Name"].unique()
)

df_filtered = df[
    df["Status"].isin(status_filter) &
    df["Requester_Name"].isin(requester_filter)
]

# ==========================================================
# NAVIGATION TABS
# ==========================================================

tab1, tab2, tab3 = st.tabs([
    "Executive Overview",
    "Service Analytics",
    "Financial Intelligence"
])

# ==========================================================
# TAB 1 — EXECUTIVE OVERVIEW
# ==========================================================

with tab1:

    completed = (df_filtered["Status"]=="Completed").sum()
    pending = (df_filtered["Status"]=="Pending").sum()
    revenue = df_filtered["Cost_Recovery"].sum()
    total_units = df_filtered[service_cols].sum().sum()

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.metric("Completed Projects", completed)
        st.markdown('</div>', unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.metric("Pending Projects", pending)
        st.markdown('</div>', unsafe_allow_html=True)

    with c3:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.metric("Total Revenue", f"${revenue:,.0f}")
        st.markdown('</div>', unsafe_allow_html=True)

    with c4:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.metric("Total Slides Processed", int(total_units))
        st.markdown('</div>', unsafe_allow_html=True)

# ==========================================================
# TAB 2 — SERVICE ANALYTICS
# ==========================================================

with tab2:

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Service Volume Distribution")

    svc_totals = df_filtered[service_cols].sum().reset_index()
    svc_totals.columns = ["Service", "Units"]

    fig = px.bar(
        svc_totals,
        x="Service",
        y="Units",
        text="Units",
        color_discrete_sequence=[PRIMARY_BLUE]
    )

    fig.update_layout(showlegend=False)
    st.plotly_chart(fig, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================================
# TAB 3 — FINANCIAL INTELLIGENCE
# ==========================================================

with tab3:

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Revenue Trend")

    revenue_trend = (
        df_filtered.groupby(["Month","Month_Label"])["Cost_Recovery"]
        .sum().reset_index()
        .sort_values("Month")
    )

    fig2 = px.line(
        revenue_trend,
        x="Month_Label",
        y="Cost_Recovery",
        markers=True,
        color_discrete_sequence=[PRIMARY_BLUE]
    )

    st.plotly_chart(fig2, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================================
# DATA TABLE
# ==========================================================

with st.expander("Operational Dataset"):
    st.dataframe(df_filtered, use_container_width=True)
