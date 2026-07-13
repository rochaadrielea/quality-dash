"""
app.py — Quality BRM Dashboard
Interactive charts with tooltips explaining every measurement and data source.
"""
from datetime import datetime, date
from io import BytesIO
from pathlib import Path
import sqlite3
import math

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from auth import check_password

DB_FILE = "quality.db"

st.set_page_config(
    page_title="BU Launchers - Quality BRM",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

if not check_password():
    st.stop()

st.markdown("""
    <style>
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; max-width: 1400px; }
    [data-testid="stMetricValue"] { font-size: 1.8rem; }
    [data-testid="stMetricLabel"] { font-size: 0.85rem; }
    .main-header {
        background: linear-gradient(90deg, #1E2761 0%, #F26E21 100%);
        padding: 0.8rem 1.5rem; border-radius: 8px; color: white; margin-bottom: 1rem;
    }
    .main-header h1 { color: white; margin: 0; font-size: 1.5rem; }
    .main-header p { color: #e0e0e0; margin: 0; font-size: 0.9rem; }
    .nav-link {
        display: block; padding: 0.4rem 0.6rem; margin: 0.2rem 0;
        border-radius: 4px; text-decoration: none; color: #1E2761;
        border-left: 3px solid transparent;
    }
    .nav-link:hover { background: #f0f2f6; border-left-color: #F26E21; }
    .filter-active {
        background: #FFF3E0; border: 1px solid #F26E21; border-radius: 6px;
        padding: 0.3rem 0.6rem; font-size: 0.8rem; color: #E65100; margin-top: 0.5rem;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="main-header">
        <h1>BU Launchers Switzerland — Quality</h1>
        <p>Business Review Dashboard</p>
    </div>
""", unsafe_allow_html=True)

if not Path(DB_FILE).exists():
    st.error(f"`{DB_FILE}` not found. Run `python ingest.py` first.")
    st.stop()


def _q(sql, params=None):
    with sqlite3.connect(DB_FILE) as conn:
        return pd.read_sql(sql, conn, params=params or [])


def to_excel_bytes(df, sheet_name="Data"):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return buf.getvalue()


CHART_CONFIG = {
    "displayModeBar": True,
    "toImageButtonOptions": {
        "format": "png",
        "height": 600,
        "width": 1000,
        "scale": 2,
    },
    "displaylogo": False,
}


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### Navigate")
    st.markdown("""
        <a href="#burndown" class="nav-link">Burndown Tracker</a>
        <a href="#monthly" class="nav-link">Monthly Performance</a>
        <a href="#dashboard" class="nav-link">Quality Dashboard</a>
        <a href="#trends" class="nav-link">Trends & Breakdown</a>
        <a href="#data-quality" class="nav-link">Data Quality</a>
    """, unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("### 🔍 Filters")
    month_label = st.text_input("Reporting month", value=datetime.now().strftime("%B %Y"))

    exercise_start = st.date_input(
        "Exercise start date", value=date(2025, 6, 1),
        help="The date when the NC closure exercise was defined. All burndown metrics use this as the baseline. Agreed in the quality review meeting.")
    target_deadline = st.date_input(
        "Closure target deadline", value=date(2026, 9, 30),
        help="Target date to close all backlog NCs. The weekly/monthly targets are calculated from the remaining open NCs divided by weeks until this date.")

    st.markdown("---")
    st.markdown("**Period filter** (dashboard + trends)")
    fc1, fc2 = st.columns(2)
    with fc1:
        date_from = st.date_input("From", value=date(2026, 1, 1))
    with fc2:
        date_to = st.date_input("To", value=date.today())

    nc_type = st.selectbox("NC type", ["All", "Production", "Supplier"])
    nc_status = st.selectbox("Status", ["All", "Open", "Closed"])

    _projects = _q("SELECT DISTINCT COALESCE(project,'(no project)') AS p FROM nc ORDER BY p")["p"].tolist()
    nc_project = st.multiselect("Project", _projects, default=[])

    _owners = _q("SELECT DISTINCT COALESCE(owner,'(no owner)') AS o FROM nc ORDER BY o")["o"].tolist()
    nc_owner = st.multiselect("Owner", _owners, default=[])

    _sources = _q("SELECT DISTINCT source FROM nc WHERE source IS NOT NULL ORDER BY source")["source"].tolist()
    if _sources:
        nc_source = st.selectbox("Data source", ["All"] + _sources)
    else:
        nc_source = "All"

    if st.button("Reset filters"):
        st.session_state.clear()
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# Filter builder
# ══════════════════════════════════════════════════════════════════════════════
def build_where(extra=None, date_col="created_on", use_date=True):
    cl, pr = [], []
    if use_date and date_from:
        cl.append(f"{date_col} >= ?"); pr.append(str(date_from))
    if use_date and date_to:
        cl.append(f"{date_col} <= ?"); pr.append(str(date_to) + " 23:59:59")
    if nc_type == "Production": cl.append("is_supplier_nc = 0")
    elif nc_type == "Supplier": cl.append("is_supplier_nc = 1")
    if nc_status == "Open": cl.append("is_open = 1")
    elif nc_status == "Closed": cl.append("is_open = 0")
    if nc_project:
        cl.append(f"COALESCE(project,'(no project)') IN ({','.join(['?']*len(nc_project))})")
        pr.extend(nc_project)
    if nc_owner:
        cl.append(f"COALESCE(owner,'(no owner)') IN ({','.join(['?']*len(nc_owner))})")
        pr.extend(nc_owner)
    if nc_source != "All":
        cl.append("source = ?"); pr.append(nc_source)
    if extra:
        for s, p in extra: cl.append(s); pr.extend(p)
    return ("WHERE " + " AND ".join(cl)) if cl else "", pr


def _qf(tmpl, extra=None, date_col="created_on", use_date=True):
    w, p = build_where(extra, date_col, use_date)
    return _q(tmpl.replace("{WHERE}", w), p)


# ══════════════════════════════════════════════════════════════════════════════
# BURNDOWN TRACKER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div id="burndown"></div>', unsafe_allow_html=True)
st.subheader(f"NC Burndown Tracker · since {exercise_start.strftime('%d %B %Y')}")

es = str(exercise_start)

backlog_at_start = _q(
    "SELECT COUNT(*) AS n FROM nc WHERE created_on < ? AND (is_open=1 OR closure_date >= ?)", [es, es]
).iloc[0]["n"]

closed_from_backlog = _q(
    "SELECT COUNT(*) AS n FROM nc WHERE created_on < ? AND is_open=0 AND closure_date >= ?", [es, es]
).iloc[0]["n"]

still_open_backlog = _q(
    "SELECT COUNT(*) AS n FROM nc WHERE created_on < ? AND is_open=1", [es]
).iloc[0]["n"]

new_since_start = _q("SELECT COUNT(*) AS n FROM nc WHERE created_on >= ?", [es]).iloc[0]["n"]
new_still_open = _q("SELECT COUNT(*) AS n FROM nc WHERE created_on >= ? AND is_open=1", [es]).iloc[0]["n"]
total_open = _q("SELECT COUNT(*) AS n FROM nc WHERE is_open=1").iloc[0]["n"]

today = date.today()
weeks_left = max(1, (target_deadline - today).days / 7)
weeks_elapsed = max(1, (today - exercise_start).days / 7)
avg_new_wk = round(new_since_start / weeks_elapsed, 1)
weekly_target = math.ceil(total_open / weeks_left + avg_new_wk)
team_size = 11

b1, b2, b3, b4 = st.columns(4)
b1.metric("Backlog at freeze", int(backlog_at_start),
          help=f"NCs that were open on {exercise_start.strftime('%d.%m.%Y')} — the starting point of the closure exercise. Includes NCs created before that date that were either still open or closed after that date. This is the fixed denominator for the burndown progress bar.")
b2.metric("Closed since start", int(closed_from_backlog),
          help=f"NCs from the original backlog that have been closed since {exercise_start.strftime('%d.%m.%Y')}. Only counts NCs created before the exercise start with a closure_date after it. This is the team's measurable progress against the backlog.")
b3.metric("Still open (backlog)", int(still_open_backlog),
          help=f"Remaining NCs from the original backlog (created before {exercise_start.strftime('%d.%m.%Y')}) that are still open. This number should decrease toward zero by the closure deadline.")
b4.metric("Total open now", int(total_open),
          help="All currently open NCs across the entire database — includes the original backlog plus any new NCs opened after the exercise started. This is the real workload the team faces.")

b5, b6, b7, b8 = st.columns(4)
b5.metric("New since start", int(new_since_start),
          help=f"NCs created after {exercise_start.strftime('%d.%m.%Y')}. These are new quality issues that arrived on top of the original backlog — the inflow that prevents the backlog from shrinking even when the team closes NCs.")
b6.metric("New still open", int(new_still_open),
          help="Of the new NCs created since the exercise start, how many remain open. High numbers here mean the team is not keeping up with the incoming rate.")
b7.metric("Avg new / week", f"{avg_new_wk}",
          help=f"{int(new_since_start)} new NCs ÷ {weeks_elapsed:.0f} weeks elapsed = {avg_new_wk} NCs/week average inflow. This rate is factored into the weekly target so the team knows how many to close beyond just the backlog.")
b8.metric("Target / week", f"{weekly_target} NCs",
          help=f"{int(total_open)} open NCs ÷ {weeks_left:.0f} weeks to deadline + {avg_new_wk} avg weekly inflow = {weekly_target} NCs/week.\n\nWith {team_size} team members: ~{max(1, round(weekly_target / team_size, 1))} NC per person per week.\n\nIf we do 1 NC/person/week: {team_size}/week × {weeks_left:.0f} weeks = {int(team_size * weeks_left)} closures possible vs {int(total_open)} needed.")

progress = closed_from_backlog / max(1, backlog_at_start)
st.progress(min(progress, 1.0), text=f"Backlog closure: {int(closed_from_backlog)} / {int(backlog_at_start)} ({progress:.0%})")
st.caption(f"Progress bar = closures from original backlog ÷ backlog at freeze. Does not include new NCs opened after {exercise_start.strftime('%d.%m.%Y')}.")

st.markdown("**Monthly burndown — closures vs target path**")
bd = _q("""
    SELECT substr(closure_date,1,7) AS month, COUNT(*) AS closed
    FROM nc WHERE created_on < ? AND is_open=0 AND closure_date >= ?
    GROUP BY month ORDER BY month
""", [es, es])

if not bd.empty:
    bd["cum"] = bd["closed"].cumsum()
    bd["remaining"] = int(backlog_at_start) - bd["cum"]
    tgt_months = pd.date_range(exercise_start, target_deadline, freq="MS").strftime("%Y-%m").tolist()
    tgt_vals = [int(backlog_at_start) - int(backlog_at_start) * (i+1) / len(tgt_months) for i in range(len(tgt_months))]

    fig = go.Figure()
    fig.add_trace(go.Bar(x=bd["month"], y=bd["closed"], name="Closed (month)", marker_color="#4CAF50",
                         hovertemplate="<b>%{x}</b><br>Closed: %{y} NCs<extra></extra>"))
    fig.add_trace(go.Scatter(x=bd["month"], y=bd["remaining"], name="Remaining",
                             mode="lines+markers", line=dict(color="#F26E21", width=3),
                             hovertemplate="<b>%{x}</b><br>Remaining backlog: %{y} NCs<extra></extra>"))
    fig.add_trace(go.Scatter(x=tgt_months, y=tgt_vals, name="Target path",
                             mode="lines", line=dict(color="#1E2761", width=2, dash="dash"),
                             hovertemplate="<b>%{x}</b><br>Target remaining: %{y:.0f} NCs<extra></extra>"))
    fig.update_layout(height=300, margin=dict(l=0, r=0, t=10, b=0),
                      legend=dict(orientation="h", y=-0.15), xaxis_title="", yaxis_title="NCs")
    st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
    st.caption(f"Green bars = NCs closed per month from the original backlog. Orange line = remaining backlog count. Dashed blue line = linear path from {int(backlog_at_start)} to zero by {target_deadline.strftime('%d.%m.%Y')}. When the orange line is above the dashed line, the team is behind target.")

    dl1, dl2, _ = st.columns([1, 1, 4])
    with dl1:
        st.download_button("📥 Excel", to_excel_bytes(bd, "Burndown"),
                           "burndown_monthly.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("No closures recorded since the exercise start date.")

st.markdown("")


# ══════════════════════════════════════════════════════════════════════════════
# CURRENT MONTH
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div id="monthly"></div>', unsafe_allow_html=True)
cm = datetime.now().strftime("%Y-%m")
cm_label = datetime.now().strftime("%B %Y")
st.subheader(f"Current Month — {cm_label}")

cm_opened = _q("SELECT COUNT(*) AS n FROM nc WHERE substr(created_on,1,7)=?", [cm]).iloc[0]["n"]
cm_closed = _q("SELECT COUNT(*) AS n FROM nc WHERE substr(closure_date,1,7)=?", [cm]).iloc[0]["n"]
cm_day = datetime.now().day
cm_pace = round(cm_closed / max(1, cm_day) * 22)

m1, m2, m3, m4 = st.columns(4)
m1.metric("Opened this month", int(cm_opened),
          help=f"NCs with created_on in {cm_label}. New quality issues entering the queue this month.")
m2.metric("Closed this month", int(cm_closed),
          help=f"NCs with closure_date in {cm_label}. The team's closure output this month.")
m3.metric("Net (closed − opened)", int(cm_closed - cm_opened),
          delta_color="normal" if cm_closed >= cm_opened else "inverse",
          help="Closed minus opened this month. Positive = backlog is shrinking. Negative = backlog is growing faster than the team can close. Target: always positive.")
m4.metric("Projected monthly close", cm_pace,
          help=f"Linear projection: {int(cm_closed)} closures in {cm_day} calendar days × 22 working days ≈ {cm_pace} by month end. This projection is more reliable after the 15th of the month.")

st.markdown("**Monthly opens vs closes (since exercise start)**")
df_mo = _q("""
    WITH o AS (SELECT substr(created_on,1,7) AS month, COUNT(*) AS opened
               FROM nc WHERE created_on >= ? GROUP BY month),
         c AS (SELECT substr(closure_date,1,7) AS month, COUNT(*) AS closed
               FROM nc WHERE closure_date >= ? GROUP BY month)
    SELECT o.month, COALESCE(o.opened,0) AS opened, COALESCE(c.closed,0) AS closed
    FROM o LEFT JOIN c ON o.month=c.month ORDER BY o.month
""", [es, es])

if not df_mo.empty:
    monthly_tgt = round(weekly_target * 4.33)
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df_mo["month"], y=df_mo["opened"], name="Opened", marker_color="#F26E21",
                         hovertemplate="<b>%{x}</b><br>Opened: %{y} NCs<extra></extra>"))
    fig.add_trace(go.Bar(x=df_mo["month"], y=df_mo["closed"], name="Closed", marker_color="#4CAF50",
                         hovertemplate="<b>%{x}</b><br>Closed: %{y} NCs<extra></extra>"))
    fig.add_hline(y=monthly_tgt, line_dash="dash", line_color="#1E2761",
                  annotation_text=f"Target: {monthly_tgt}/mo", annotation_position="top left")
    fig.update_layout(height=280, margin=dict(l=0, r=0, t=10, b=0), barmode="group",
                      legend=dict(orientation="h", y=-0.15), xaxis_title="", yaxis_title="NCs")
    st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
    st.caption(f"Orange = new NCs opened per month. Green = NCs closed per month. Dashed line = monthly closure target ({monthly_tgt}/mo), derived from weekly target ({weekly_target}/wk) × 4.33 weeks/month. Green bars should consistently exceed both orange bars and the target line.")

    dl1, dl2, _ = st.columns([1, 1, 4])
    with dl1:
        st.download_button("📥 Excel", to_excel_bytes(df_mo, "Monthly_Trend"),
                           "monthly_trend.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_monthly")

st.markdown("")


# ══════════════════════════════════════════════════════════════════════════════
# QUALITY DASHBOARD (filtered)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div id="dashboard"></div>', unsafe_allow_html=True)
st.subheader(f"Quality Dashboard · {month_label}")

kpis = _qf("""
    SELECT
        SUM(CASE WHEN is_open=1 THEN 1 ELSE 0 END) AS ncs_wip,
        SUM(CASE WHEN is_open=0 THEN 1 ELSE 0 END) AS ncs_closed,
        SUM(CASE WHEN classification LIKE 'Major%%' AND is_open=1 THEN 1 ELSE 0 END) AS major_open,
        SUM(CASE WHEN is_supplier_nc=1 AND is_open=1 THEN 1 ELSE 0 END) AS supplier_open,
        SUM(CASE WHEN is_supplier_nc=0 AND is_open=1 THEN 1 ELSE 0 END) AS production_open,
        COUNT(*) AS total_filtered
    FROM nc {WHERE}
""").iloc[0]

total_all = _q("SELECT COUNT(*) AS n FROM nc").iloc[0]["n"]
tf = int(kpis["total_filtered"] or 0)
if tf < total_all:
    st.caption(f"Showing **{tf:,}** of {total_all:,} NCs (filtered: {date_from} → {date_to})")

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("NCs Open (WIP)", int(kpis["ncs_wip"] or 0),
          help="All NCs with status OPEN in the filtered period. This is the active workload — each of these needs an owner driving it to closure.")
c2.metric("NCs Closed", int(kpis["ncs_closed"] or 0),
          help="All NCs with status CLOSED in the filtered period. Includes both backlog closures and closures of newly opened NCs.")
c3.metric("Major NCs Open", int(kpis["major_open"] or 0),
          help="Open NCs classified as Major (classification field starts with 'Major'). These require NRB disposition review and carry higher risk. Prioritize for closure.")
c4.metric("Production NCs", int(kpis["production_open"] or 0),
          help="Open NCs originating from internal production processes (notification type Z3). These are issues found during manufacturing, assembly, or testing at the Emmen site.")
c5.metric("Supplier NCs", int(kpis["supplier_open"] or 0),
          help="Open NCs originating from supplier/procurement complaints (notification type Z2). These are incoming material or component issues requiring supplier follow-up.")

st.markdown("")


# ══════════════════════════════════════════════════════════════════════════════
# TRENDS & BREAKDOWN (filtered)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div id="trends"></div>', unsafe_allow_html=True)
st.subheader("Trends & Breakdown")

col_a, col_b = st.columns(2)

with col_a:
    st.markdown("**Top 6 Projects — NCs WIP**")
    df_proj = _qf("SELECT COALESCE(project,'(no project)') AS project, COUNT(*) AS open_ncs FROM nc {WHERE} GROUP BY project ORDER BY open_ncs DESC LIMIT 6",
                   extra=[("is_open=1", [])])
    if not df_proj.empty:
        fig = px.bar(df_proj, x="open_ncs", y="project", orientation="h",
                     color_discrete_sequence=["#1E2761"], text="open_ncs")
        fig.update_layout(height=260, margin=dict(l=0, r=0, t=10, b=0), showlegend=False,
                          yaxis=dict(categoryorder="total ascending"), xaxis_title="", yaxis_title="")
        fig.update_traces(textposition="outside", hovertemplate="<b>%{y}</b><br>Open NCs: %{x}<extra></extra>")
        st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
        st.caption("Projects with the most open NCs. '(no project)' = NCs missing a project assignment — these need data cleanup.")
        st.download_button("📥 Excel", to_excel_bytes(df_proj, "Projects_WIP"),
                           "projects_wip.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_proj_wip")
    else:
        st.info("No open NCs in range.")

with col_b:
    st.markdown("**NCs by Detection Area**")
    df_area = _qf("SELECT COALESCE(detection_area,'BLANK - to clean') AS area, COUNT(*) AS n FROM nc {WHERE} GROUP BY area ORDER BY n DESC LIMIT 12")
    if not df_area.empty:
        colors = ["#E53E3E" if a.startswith("BLANK") else "#F26E21" for a in df_area["area"]]
        fig = px.bar(df_area, x="n", y="area", orientation="h", color=df_area["area"],
                     color_discrete_sequence=colors, text="n")
        fig.update_layout(height=260, margin=dict(l=0, r=0, t=10, b=0), showlegend=False,
                          yaxis=dict(categoryorder="total ascending"), xaxis_title="", yaxis_title="")
        fig.update_traces(textposition="outside", hovertemplate="<b>%{y}</b><br>NCs: %{x}<extra></extra>")
        st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
        st.caption("Where NCs were detected. Red = missing detection area (data quality issue). Source: detection_area field from the NC tracker.")
        st.download_button("📥 Excel", to_excel_bytes(df_area, "Detection_Area"),
                           "detection_area.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_area")
    else:
        st.info("No NCs in range.")

col_c, col_d = st.columns(2)

with col_c:
    st.markdown("**Production vs Supplier**")
    df_ps = _qf("""SELECT CASE WHEN is_supplier_nc=1 THEN 'Supplier' ELSE 'Production' END AS source,
                SUM(CASE WHEN is_open=1 THEN 1 ELSE 0 END) AS open_ncs,
                SUM(CASE WHEN is_open=0 THEN 1 ELSE 0 END) AS closed_ncs
                FROM nc {WHERE} GROUP BY source""")
    if not df_ps.empty:
        fig = px.bar(df_ps, x="source", y=["open_ncs", "closed_ncs"], barmode="group",
                     color_discrete_map={"open_ncs": "#F26E21", "closed_ncs": "#4CAF50"}, text_auto=True)
        fig.update_traces(textposition="outside")
        fig.update_layout(height=280, margin=dict(l=0, r=0, t=10, b=0), legend_title="", xaxis_title="", yaxis_title="")
        st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
        st.caption("Production = internal manufacturing NCs (Z3). Supplier = procurement complaints (Z2). Orange = still open, green = closed. Source: is_supplier_nc flag derived from notification type or supplier_name presence.")
        st.download_button("📥 Excel", to_excel_bytes(df_ps, "Prod_vs_Supplier"),
                           "prod_vs_supplier.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_ps")
    else:
        st.info("No NCs in range.")

with col_d:
    st.markdown("**Open NCs by Owner** (Top 10)")
    df_own = _qf("""SELECT COALESCE(owner,'(no owner)') AS owner,
                SUM(CASE WHEN is_open=1 THEN 1 ELSE 0 END) AS open_ncs,
                MAX(days_open) AS oldest_days
                FROM nc {WHERE} GROUP BY owner HAVING open_ncs>0 ORDER BY open_ncs DESC LIMIT 10""")
    if not df_own.empty:
        st.dataframe(df_own, use_container_width=True, hide_index=True, height=280,
                     column_config={"owner": "Owner",
                                    "open_ncs": st.column_config.NumberColumn("Open"),
                                    "oldest_days": st.column_config.NumberColumn("Oldest (days)")})
        st.caption("Open NCs per owner with the age of their oldest open NC. Owners with high 'Oldest' values may need support or escalation. Source: owner field from the NC tracker, days_open calculated from created_on to today.")
        st.download_button("📥 Excel", to_excel_bytes(df_own, "Owner_Workload"),
                           "owner_workload.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_own")
    else:
        st.info("No open NCs in range.")

st.markdown("**NC Total per Year** (all data)")
df_yr = _q("SELECT substr(created_on,1,4) AS yr, COUNT(*) AS n FROM nc WHERE created_on IS NOT NULL GROUP BY yr ORDER BY yr")
fig = px.bar(df_yr, x="yr", y="n", color_discrete_sequence=["#1E2761"], text="n")
fig.update_layout(height=240, margin=dict(l=0, r=0, t=10, b=0), showlegend=False, xaxis_title="", yaxis_title="")
fig.update_traces(textposition="outside", hovertemplate="<b>%{x}</b><br>Total NCs: %{y}<extra></extra>")
st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
st.caption("Historical NC volume by year. Not affected by sidebar date filters — always shows the full dataset for trend context. Source: created_on year extracted from quality.db.")
dl1, dl2, _ = st.columns([1, 1, 4])
with dl1:
    st.download_button("📥 Excel", to_excel_bytes(df_yr, "Yearly"),
                       "nc_per_year.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="dl_yr")

st.markdown("")


# ══════════════════════════════════════════════════════════════════════════════
# DATA QUALITY
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div id="data-quality"></div>', unsafe_allow_html=True)
st.subheader("Data Quality — rows needing cleaning")

incomplete = _qf("""SELECT nc_id, owner, project, detection_area, classification, status
    FROM nc {WHERE} ORDER BY owner, nc_id""",
    extra=[("(project IS NULL OR detection_area IS NULL OR classification IS NULL OR owner IS NULL)", [])])
st.caption(f"**{len(incomplete)} NCs** with at least one missing field (project, detection area, classification, or owner). Send this list to the responsible owners for cleanup.")
st.dataframe(incomplete, use_container_width=True, hide_index=True, height=300)

dl1, dl2, _ = st.columns([1, 1, 4])
with dl1:
    st.download_button("📥 Excel", to_excel_bytes(incomplete, "Data_Quality"),
                       "data_quality.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="dl_dq")

st.caption(f"Data source: `quality.db` · Last modified: "
           f"{datetime.fromtimestamp(Path(DB_FILE).stat().st_mtime).strftime('%d.%m.%Y %H:%M')}")