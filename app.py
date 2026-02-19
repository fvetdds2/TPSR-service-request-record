import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl
from datetime import datetime

st.set_page_config(page_title="TPSR Service Request Dashboard", layout="wide")

# ─────────────────────────────────────────
# Load & Clean Data
# ─────────────────────────────────────────
@st.cache_data
def load_data():
    wb = openpyxl.load_workbook("cost_recovery_record_from_2025.xlsx")
    ws = wb.active

    # Build header map from row 1: col_number -> header name
    header_map = {cell.column: cell.value for cell in ws[1] if cell.value}

    # Service columns = E(5) through L(12); M(13) = specie (not a service)
    service_col_nums = [5, 6, 7, 8, 9, 10, 11, 12]
    service_cols     = [header_map[c] for c in service_col_nums]

    # Col E (FFPE processing & Embedding) is sometimes stored as an Excel
    # date serial. Convert it back to the integer count (days since 1899-12-30).
    def excel_date_to_int(val):
        if isinstance(val, datetime):
            return (val - datetime(1899, 12, 30)).days
        return val

    records = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        if all(cell.value is None for cell in row):
            continue

        vals = {cell.column: cell.value for cell in row if cell.value is not None}

        name   = vals.get(1)
        date   = vals.get(2)
        status = vals.get(3)
        cost   = vals.get(4, 0) or 0

        svc = {}
        for col_num, col_name in zip(service_col_nums, service_cols):
            raw = vals.get(col_num)
            raw = excel_date_to_int(raw)
            svc[col_name] = float(raw) if isinstance(raw, (int, float)) else 0.0

        records.append({
            "Requester_Name": name,
            "Required_Date":  pd.to_datetime(date, errors="coerce"),
            "Status":         status or "Unknown",
            "Cost_Recovery":  float(cost) if isinstance(cost, (int, float)) else 0.0,
            **svc,
        })

    df = pd.DataFrame(records)
    df["Month_Year"]       = df["Required_Date"].dt.to_period("M")
    df["Month_Year_Label"] = df["Required_Date"].dt.strftime("%b %Y")

    return df, service_cols


df, service_cols = load_data()

# Short display labels for charts
short_labels = {
    "FFPE processing & Embedding":               "FFPE Process & Embed",
    "FFPE sectioning & H&E stain":               "FFPE Section & H&E",
    "Frozen sectioning-unstained slide":         "Frozen Unstained",
    "Frozen sectioning & H&E stain":             "Frozen H&E",
    "Frozen sectioning-step section":            "Frozen Step Section",
    "Repository FFPE sectioning-unstained slide":"Repo FFPE Unstained",
    "histology tissue collection vials":         "Histology Vials",
    "histopathology support (hr)":               "Histopath Support (hr)",
}

# ─────────────────────────────────────────
# Sidebar Filters
# ─────────────────────────────────────────
st.sidebar.header("🔍 Filters")

status_options   = sorted(df["Status"].unique().tolist())
status_filter    = st.sidebar.multiselect("Status",    options=status_options,   default=status_options)

requester_options = sorted(df["Requester_Name"].dropna().unique().tolist())
requester_filter  = st.sidebar.multiselect("Requester", options=requester_options, default=requester_options)

df_filtered = df[
    df["Status"].isin(status_filter) &
    df["Requester_Name"].isin(requester_filter)
]

# ─────────────────────────────────────────
# Title
# ─────────────────────────────────────────
st.title("MMC Translational Pathology Shared Resource Core Service Request Dashboard")
st.caption("Cost Recovery Record – 2025 / 2026")
st.divider()

# ─────────────────────────────────────────
# Metrics
# ─────────────────────────────────────────
completed      = int((df_filtered["Status"] == "Completed").sum())
pending        = int((df_filtered["Status"] == "Pending").sum())
total_cost     = df_filtered["Cost_Recovery"].sum()
total_services = int(df_filtered[service_cols].sum().sum())

c1, c2, c3, c4 = st.columns(4)
c1.metric("✅ Completed",           completed)
c2.metric("⏳ Pending",              pending)
c3.metric("💰 Total Cost Recovery",  f"${total_cost:,.2f}")
c4.metric("🧪 Total Service Units",  total_services)
st.divider()

# ─────────────────────────────────────────
# Row 1 — Status Pie | Requests by Requester
# ─────────────────────────────────────────
left, right = st.columns(2)

with left:
    st.subheader("Status Breakdown")
    sc = df_filtered["Status"].value_counts().reset_index()
    sc.columns = ["Status", "Count"]
    fig_pie = px.pie(
        sc, names="Status", values="Count", hole=0.4,
        color="Status",
        color_discrete_map={"Completed": "#2ecc71", "Pending": "#e67e22", "Unknown": "#95a5a6"},
    )
    fig_pie.update_traces(textinfo="label+percent+value")
    st.plotly_chart(fig_pie, use_container_width=True)

with right:
    st.subheader("Requests by Requester")
    req_df = (
        df_filtered
        .groupby("Requester_Name")
        .agg(Count=("Status", "count"), Cost=("Cost_Recovery", "sum"))
        .reset_index()
        .sort_values("Count", ascending=False)
    )
    fig_req = px.bar(
        req_df, x="Requester_Name", y="Count",
        color="Cost", color_continuous_scale="Blues", text="Count",
        labels={"Requester_Name": "Requester", "Count": "No. of Requests", "Cost": "Cost ($)"},
    )
    fig_req.update_traces(textposition="outside")
    fig_req.update_layout(xaxis_tickangle=-30, coloraxis_colorbar_title="Cost ($)")
    st.plotly_chart(fig_req, use_container_width=True)

# ─────────────────────────────────────────
# Service Types Distribution (Cols E–L)
# ─────────────────────────────────────────
st.subheader("📊 Service Types Distribution (Columns E – L)")

svc_totals = df_filtered[service_cols].sum().reset_index()
svc_totals.columns = ["Service_Type", "Total_Units"]
svc_totals["Short_Label"] = svc_totals["Service_Type"].map(short_labels)
svc_totals = svc_totals.sort_values("Total_Units", ascending=False)

fig_svc = px.bar(
    svc_totals, x="Short_Label", y="Total_Units",
    color="Short_Label",
    color_discrete_sequence=px.colors.qualitative.Safe,
    text="Total_Units",
    labels={"Short_Label": "Service Type", "Total_Units": "Total Units"},
)
fig_svc.update_traces(textposition="outside")
fig_svc.update_layout(showlegend=False, xaxis_title="Service Type", yaxis_title="Total Units")
st.plotly_chart(fig_svc, use_container_width=True)

# Per-requester grouped bar
st.subheader("Service Units per Requester by Type")
melt = (
    df_filtered[["Requester_Name"] + service_cols]
    .melt(id_vars="Requester_Name", value_vars=service_cols,
          var_name="Service_Type", value_name="Units")
)
melt["Service_Type"] = melt["Service_Type"].map(short_labels)
melt = melt[melt["Units"] > 0]

if not melt.empty:
    fig_grp = px.bar(
        melt, x="Requester_Name", y="Units",
        color="Service_Type", barmode="group", text="Units",
        color_discrete_sequence=px.colors.qualitative.Safe,
        labels={"Requester_Name": "Requester", "Units": "Units", "Service_Type": "Service"},
    )
    fig_grp.update_traces(textposition="outside")
    fig_grp.update_layout(xaxis_tickangle=-30)
    st.plotly_chart(fig_grp, use_container_width=True)
else:
    st.info("No service unit data to display for the selected filters.")

# ─────────────────────────────────────────
# Total Service Count by Month
# ─────────────────────────────────────────
st.subheader("📅 Total Service Count by Month")

svc_by_month = (
    df_filtered
    .assign(Total_Units=df_filtered[service_cols].sum(axis=1))
    .groupby(["Month_Year", "Month_Year_Label"], as_index=False)["Total_Units"]
    .sum()
    .sort_values("Month_Year")
)

if not svc_by_month.empty and svc_by_month["Total_Units"].sum() > 0:
    fig_month = px.bar(
        svc_by_month, x="Month_Year_Label", y="Total_Units",
        color="Total_Units", color_continuous_scale="Teal", text="Total_Units",
        labels={"Month_Year_Label": "Month", "Total_Units": "Total Service Units"},
    )
    fig_month.update_traces(textposition="outside")
    fig_month.update_layout(
        xaxis_title="Month / Year", yaxis_title="Total Service Units",
        coloraxis_showscale=False,
    )
    st.plotly_chart(fig_month, use_container_width=True)
else:
    st.info("No service unit data available for the selected filters.")

# Stacked service type by month
st.subheader("Service Type Count by Month")

melt_month = (
    df_filtered[["Month_Year", "Month_Year_Label"] + service_cols]
    .melt(id_vars=["Month_Year", "Month_Year_Label"],
          value_vars=service_cols, var_name="Service_Type", value_name="Units")
)
melt_month["Service_Type"] = melt_month["Service_Type"].map(short_labels)
melt_month = (
    melt_month[melt_month["Units"] > 0]
    .groupby(["Month_Year", "Month_Year_Label", "Service_Type"], as_index=False)["Units"]
    .sum()
    .sort_values("Month_Year")
)

if not melt_month.empty:
    fig_svc_month = px.bar(
        melt_month, x="Month_Year_Label", y="Units",
        color="Service_Type", barmode="stack", text="Units",
        color_discrete_sequence=px.colors.qualitative.Safe,
        labels={"Month_Year_Label": "Month", "Units": "Units", "Service_Type": "Service"},
    )
    fig_svc_month.update_traces(textposition="inside")
    fig_svc_month.update_layout(xaxis_title="Month / Year", yaxis_title="Service Units")
    st.plotly_chart(fig_svc_month, use_container_width=True)
else:
    st.info("No monthly service breakdown available.")

# ─────────────────────────────────────────
# Cost Recovery Over Time
# ─────────────────────────────────────────
st.subheader("💵 Cost Recovery Over Time")

cost_time = (
    df_filtered
    .groupby(["Month_Year", "Month_Year_Label"])["Cost_Recovery"]
    .sum()
    .reset_index()
    .sort_values("Month_Year")
)

if not cost_time.empty:
    fig_time = px.line(
        cost_time, x="Month_Year_Label", y="Cost_Recovery",
        markers=True, color_discrete_sequence=["#3498db"],
        labels={"Month_Year_Label": "Month", "Cost_Recovery": "Cost Recovery ($)"},
    )
    fig_time.update_traces(line_width=2.5, marker_size=8)
    fig_time.update_layout(xaxis_title="Month / Year", yaxis_title="Cost ($)")
    st.plotly_chart(fig_time, use_container_width=True)
else:
    st.info("No date data available.")

# ─────────────────────────────────────────
# Raw Data Table
# ─────────────────────────────────────────
with st.expander("📋 Raw Data Table"):
    display_df = df_filtered[
        ["Requester_Name", "Month_Year_Label", "Status", "Cost_Recovery"] + service_cols
    ].copy()
    display_df = display_df.rename(columns={"Month_Year_Label": "Month / Year", **short_labels})
    st.dataframe(
        display_df.style.format({"Cost_Recovery": "${:,.2f}"}),
        use_container_width=True,
    )