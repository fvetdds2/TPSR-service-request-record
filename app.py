import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl
from datetime import datetime, timedelta

st.set_page_config(page_title="TPSR Service Request Dashboard", layout="wide")

# -----------------------------
# Load & Clean Data
# -----------------------------
@st.cache_data
def load_data():
    # Use openpyxl to read raw values so Excel date-formatted numbers are handled correctly
    wb = openpyxl.load_workbook("cost_recovery_record_from_2025.xlsx")
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    headers = list(rows[0])

    # Rename None headers (columns beyond named ones) to generic labels
    unnamed_count = 0
    clean_headers = []
    for h in headers:
        if h is None:
            unnamed_count += 1
            clean_headers.append(f"Service_Extra_{unnamed_count}")
        else:
            clean_headers.append(h)

    df = pd.DataFrame(rows[1:], columns=clean_headers)

    # Fix Required_Date
    df["Required_Date"] = pd.to_datetime(df["Required_Date"], errors="coerce")

    # Fix FFPE Processing & Embedding — Excel stored a count as a date serial (e.g. 121 → 1900-05-01)
    # Convert datetime back to Excel serial day number (days since 1899-12-30)
    def fix_excel_date_as_number(val):
        if isinstance(val, datetime):
            return (val - datetime(1899, 12, 30)).days
        return val

    col_ffpe = "FFPE processing & Embedding"
    df[col_ffpe] = df[col_ffpe].apply(fix_excel_date_as_number)

    # Numeric conversion for service columns
    service_cols = [
        "FFPE processing & Embedding",
        "FFPE sectioning & H&E stain",
        "Frozen sectioning-unstained slide",
        "Frozen sectioning & H&E stain",
        "Repository FFPE sectioning-unstained slide",
    ]
    for col in service_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["Cost_Recovery"] = pd.to_numeric(df["Cost_Recovery"], errors="coerce").fillna(0)
    df["Status"] = df["Status"].fillna("Unknown")

    # Month-Year label for grouping (e.g. "Sep 2025")
    df["Month_Year"] = df["Required_Date"].dt.to_period("M")
    df["Month_Year_Label"] = df["Required_Date"].dt.strftime("%b %Y")

    return df, service_cols


df, service_cols = load_data()

# -----------------------------
# Sidebar Filters
# -----------------------------
st.sidebar.header("🔍 Filters")

status_options = df["Status"].unique().tolist()
status_filter = st.sidebar.multiselect(
    "Status", options=status_options, default=status_options
)

requester_options = df["Requester_Name"].dropna().unique().tolist()
requester_filter = st.sidebar.multiselect(
    "Requester", options=requester_options, default=requester_options
)

df_filtered = df[
    (df["Status"].isin(status_filter)) &
    (df["Requester_Name"].isin(requester_filter))
]

# -----------------------------
# Title
# -----------------------------
st.title("MMC Translational Pathology Shared Resource Core Service Request Dashboard")
st.caption("Cost Recovery Record – 2025/2026")

st.divider()

# -----------------------------
# Metrics Row
# -----------------------------
completed = df_filtered[df_filtered["Status"] == "Completed"].shape[0]
pending   = df_filtered[df_filtered["Status"] == "Pending"].shape[0]
total_cost = df_filtered["Cost_Recovery"].sum()
total_services = df_filtered[service_cols].sum().sum()

col1, col2, col3, col4 = st.columns(4)
col1.metric("✅ Completed", completed)
col2.metric("⏳ Pending", pending)
col3.metric("💰 Total Cost Recovery", f"${total_cost:,.2f}")
col4.metric("🧪 Total Service Units", int(total_services))

st.divider()

# -----------------------------
# Row 1: Status Pie | Services by Requester
# -----------------------------
row1_left, row1_right = st.columns(2)

with row1_left:
    st.subheader("Status Breakdown")
    status_counts = df_filtered["Status"].value_counts().reset_index()
    status_counts.columns = ["Status", "Count"]
    fig_pie = px.pie(
        status_counts,
        names="Status",
        values="Count",
        color="Status",
        color_discrete_map={"Completed": "#2ecc71", "Pending": "#e67e22", "Unknown": "#95a5a6"},
        hole=0.4,
    )
    fig_pie.update_traces(textinfo="label+percent+value")
    st.plotly_chart(fig_pie, use_container_width=True)

with row1_right:
    st.subheader("Requests by Requester")
    req_df = (
        df_filtered.groupby("Requester_Name")
        .agg(Count=("Status", "count"), Cost=("Cost_Recovery", "sum"))
        .reset_index()
        .sort_values("Count", ascending=False)
    )
    fig_req = px.bar(
        req_df,
        x="Requester_Name",
        y="Count",
        color="Cost",
        color_continuous_scale="Blues",
        labels={"Requester_Name": "Requester", "Count": "No. of Requests", "Cost": "Cost ($)"},
        text="Count",
    )
    fig_req.update_layout(xaxis_tickangle=-30, coloraxis_colorbar_title="Cost ($)")
    fig_req.update_traces(textposition="outside")
    st.plotly_chart(fig_req, use_container_width=True)

# -----------------------------
# Row 2: Service Types Bar Chart (Column E to I)
# -----------------------------
st.subheader("📊 Service Types Distribution (Columns E – I)")

service_totals = df_filtered[service_cols].sum().reset_index()
service_totals.columns = ["Service_Type", "Total_Units"]
service_totals = service_totals.sort_values("Total_Units", ascending=False)

# Short labels for readability
short_labels = {
    "FFPE processing & Embedding":                  "FFPE Process & Embed",
    "FFPE sectioning & H&E stain":                  "FFPE Section & H&E",
    "Frozen sectioning & H&E slide":                "Frozen Section & H&E",
    "Repository FFPE sectioning & H&E stain":       "Repo FFPE H&E",
    "Repository FFPE sectioning-unstained slide":   "Repo FFPE Unstained",
}
service_totals["Short_Label"] = service_totals["Service_Type"].map(short_labels)

fig_svc = px.bar(
    service_totals,
    x="Short_Label",
    y="Total_Units",
    color="Short_Label",
    color_discrete_sequence=px.colors.qualitative.Safe,
    labels={"Short_Label": "Service Type", "Total_Units": "Total Units"},
    text="Total_Units",
)
fig_svc.update_traces(textposition="outside")
fig_svc.update_layout(showlegend=False, xaxis_title="Service Type", yaxis_title="Total Units")
st.plotly_chart(fig_svc, use_container_width=True)

# Per-requester breakdown of service types
st.subheader("Service Units per Requester by Type")
per_req = df_filtered[["Requester_Name"] + service_cols].copy()
per_req_melt = per_req.melt(
    id_vars="Requester_Name",
    value_vars=service_cols,
    var_name="Service_Type",
    value_name="Units",
)
per_req_melt["Service_Type"] = per_req_melt["Service_Type"].map(short_labels)
per_req_melt = per_req_melt[per_req_melt["Units"] > 0]

if not per_req_melt.empty:
    fig_grouped = px.bar(
        per_req_melt,
        x="Requester_Name",
        y="Units",
        color="Service_Type",
        barmode="group",
        labels={"Requester_Name": "Requester", "Units": "Units", "Service_Type": "Service"},
        color_discrete_sequence=px.colors.qualitative.Safe,
        text="Units",
    )
    fig_grouped.update_layout(xaxis_tickangle=-30)
    fig_grouped.update_traces(textposition="outside")
    st.plotly_chart(fig_grouped, use_container_width=True)
else:
    st.info("No service unit data to display for the selected filters.")

# -----------------------------
# Total Service Count by Month & Year
# -----------------------------
st.subheader("📅 Total Service Count by Month")

# Group total service units by Month_Year, preserving sort order
svc_by_month = (
    df_filtered.copy()
    .assign(Total_Units=df_filtered[service_cols].sum(axis=1))
    .groupby(["Month_Year", "Month_Year_Label"], as_index=False)["Total_Units"]
    .sum()
    .sort_values("Month_Year")
)

if not svc_by_month.empty and svc_by_month["Total_Units"].sum() > 0:
    fig_month = px.bar(
        svc_by_month,
        x="Month_Year_Label",
        y="Total_Units",
        color="Total_Units",
        color_continuous_scale="Teal",
        labels={"Month_Year_Label": "Month", "Total_Units": "Total Service Units"},
        text="Total_Units",
    )
    fig_month.update_traces(textposition="outside")
    fig_month.update_layout(
        xaxis_title="Month / Year",
        yaxis_title="Total Service Units",
        coloraxis_showscale=False,
    )
    st.plotly_chart(fig_month, use_container_width=True)
else:
    st.info("No service unit data available for the selected filters.")

# Per-service-type count by Month
st.subheader("Service Type Count by Month")

svc_melt_month = (
    df_filtered[["Month_Year", "Month_Year_Label"] + service_cols]
    .melt(id_vars=["Month_Year", "Month_Year_Label"], value_vars=service_cols,
          var_name="Service_Type", value_name="Units")
)
svc_melt_month["Service_Type"] = svc_melt_month["Service_Type"].map(short_labels)
svc_melt_month = (
    svc_melt_month[svc_melt_month["Units"] > 0]
    .groupby(["Month_Year", "Month_Year_Label", "Service_Type"], as_index=False)["Units"]
    .sum()
    .sort_values("Month_Year")
)

if not svc_melt_month.empty:
    fig_svc_month = px.bar(
        svc_melt_month,
        x="Month_Year_Label",
        y="Units",
        color="Service_Type",
        barmode="stack",
        labels={"Month_Year_Label": "Month", "Units": "Units", "Service_Type": "Service"},
        color_discrete_sequence=px.colors.qualitative.Safe,
        text="Units",
    )
    fig_svc_month.update_traces(textposition="inside")
    fig_svc_month.update_layout(xaxis_title="Month / Year", yaxis_title="Service Units")
    st.plotly_chart(fig_svc_month, use_container_width=True)
else:
    st.info("No service unit breakdown available.")

# -----------------------------
# Row 3: Cost Recovery Over Time
# -----------------------------
st.subheader("💵 Cost Recovery Over Time")

cost_time = (
    df_filtered.groupby(["Month_Year", "Month_Year_Label"])["Cost_Recovery"]
    .sum()
    .reset_index()
    .sort_values("Month_Year")
)
cost_time["Required_Date"] = cost_time["Month_Year_Label"]

if len(cost_time) > 0:
    fig_time = px.line(
        cost_time,
        x="Required_Date",
        y="Cost_Recovery",
        markers=True,
        labels={"Required_Date": "Month", "Cost_Recovery": "Cost Recovery ($)"},
        color_discrete_sequence=["#3498db"],
    )
    fig_time.update_traces(line_width=2.5, marker_size=8)
    fig_time.update_layout(xaxis_title="Month", yaxis_title="Cost ($)")
    st.plotly_chart(fig_time, use_container_width=True)
else:
    st.info("No date data available.")

# -----------------------------
# Raw Data
# -----------------------------
with st.expander("📋 Raw Data Table"):
    display_df = df_filtered[["Requester_Name", "Month_Year_Label", "Status", "Cost_Recovery"] + service_cols].copy()
    display_df = display_df.rename(columns={"Month_Year_Label": "Month / Year"})
    st.dataframe(
        display_df.style.format({"Cost_Recovery": "${:,.2f}"}),
        use_container_width=True,
    )
