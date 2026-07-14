"""
app.py — Quality BRM Dashboard
Interactive charts with tooltips explaining every measurement and data source.
Compatible with both old (tracker-only) and new (merged) quality.db schemas.
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
    /* Highlight the reach-zero calculator section so it draws attention */
    div[data-testid="stVerticalBlockBorderWrapper"]:has(div.reach-zero-anchor) {
        background: #EAF2FB;
        border: 2px solid #2a78d6 !important;
        border-radius: 12px;
        padding: 0.5rem 0.75rem;
    }
    /* KPI cards with visible descriptions */
    .kpi-card {
        padding: 0.4rem 0.2rem 0.9rem 0;
    }
    .kpi-label {
        font-size: 0.82rem; color: #5B6B78; font-weight: 600; margin-bottom: 0.1rem;
    }
    .kpi-value {
        font-size: 1.9rem; color: #1E2761; font-weight: 700; line-height: 1.1;
    }
    .kpi-desc {
        font-size: 0.78rem; color: #7A8894; margin-top: 0.25rem; line-height: 1.35;
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


def _has_column(col_name):
    """Check if a column exists in the nc table."""
    try:
        cols = _q("PRAGMA table_info(nc)")
        return col_name in cols["name"].values
    except Exception:
        return False


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

# Check schema once at startup
HAS_SOURCE = _has_column("source")
HAS_COPQ = _has_column("copq")


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

    # Exercise start / deadline stay OUTSIDE the form — the burndown section has
    # its own live deadline slider, and these feed it directly.
    exercise_start = st.date_input(
        "Exercise start date", value=date(2025, 6, 1),
        help="The date when the NC closure exercise was defined. All burndown metrics use this as the baseline.")

    # ---- Defaults for the general filters (committed values live in session_state) ----
    # From defaults to the earliest NC in the data, so the full story shows by default.
    _min_row = _q("SELECT MIN(created_on) AS m FROM nc WHERE created_on IS NOT NULL")
    _earliest = None
    if not _min_row.empty and _min_row.iloc[0]["m"]:
        try:
            _earliest = pd.to_datetime(_min_row.iloc[0]["m"]).date()
        except Exception:
            _earliest = None
    _default_from = _earliest or date(2015, 1, 1)

    _filter_defaults = {
        "date_from": _default_from,
        "date_to": date.today(),
        "nc_type": "All",
        "nc_status": "All",
        "nc_project": [],
        "nc_owner": [],
        "nc_source": "All",
    }
    for _k, _v in _filter_defaults.items():
        st.session_state.setdefault(_k, _v)

    _projects = _q("SELECT DISTINCT COALESCE(project,'(no project)') AS p FROM nc ORDER BY p")["p"].tolist()
    _owners = _q("SELECT DISTINCT COALESCE(owner,'(no owner)') AS o FROM nc ORDER BY o")["o"].tolist()
    _sources = []
    if HAS_SOURCE:
        _sources = _q("SELECT DISTINCT source FROM nc WHERE source IS NOT NULL ORDER BY source")["source"].tolist()

    st.markdown("---")
    # ---- General period filter: only applies when the button is clicked ----
    with st.form("general_filters", clear_on_submit=False):
        st.markdown("**Period filter** (dashboard + trends)")
        fc1, fc2 = st.columns(2)
        with fc1:
            f_date_from = st.date_input("From", value=st.session_state["date_from"])
        with fc2:
            f_date_to = st.date_input("To", value=st.session_state["date_to"])

        f_nc_type = st.selectbox("NC type", ["All", "Production", "Supplier"],
                                 index=["All", "Production", "Supplier"].index(st.session_state["nc_type"]))
        f_nc_status = st.selectbox("Status", ["All", "Open", "Closed"],
                                   index=["All", "Open", "Closed"].index(st.session_state["nc_status"]))
        f_nc_project = st.multiselect("Project", _projects,
                                      default=[p for p in st.session_state["nc_project"] if p in _projects])
        f_nc_owner = st.multiselect("Owner", _owners,
                                    default=[o for o in st.session_state["nc_owner"] if o in _owners])
        if _sources:
            f_nc_source = st.selectbox("Data source", ["All"] + _sources,
                                       index=(["All"] + _sources).index(st.session_state["nc_source"])
                                       if st.session_state["nc_source"] in (["All"] + _sources) else 0)
        else:
            f_nc_source = "All"

        c_apply, c_reset = st.columns(2)
        applied = c_apply.form_submit_button("✓ Apply filters", use_container_width=True, type="primary")
        reset = c_reset.form_submit_button("Reset", use_container_width=True)

    if applied:
        st.session_state["date_from"] = f_date_from
        st.session_state["date_to"] = f_date_to
        st.session_state["nc_type"] = f_nc_type
        st.session_state["nc_status"] = f_nc_status
        st.session_state["nc_project"] = f_nc_project
        st.session_state["nc_owner"] = f_nc_owner
        st.session_state["nc_source"] = f_nc_source
        st.rerun()

    if reset:
        for _k, _v in _filter_defaults.items():
            st.session_state[_k] = _v
        st.rerun()

    # Read committed filter values (used by build_where downstream)
    date_from = st.session_state["date_from"]
    date_to = st.session_state["date_to"]
    nc_type = st.session_state["nc_type"]
    nc_status = st.session_state["nc_status"]
    nc_project = st.session_state["nc_project"]
    nc_owner = st.session_state["nc_owner"]
    nc_source = st.session_state["nc_source"]

    # Show which filters are currently active
    _active = []
    if (date_from, date_to) != (_filter_defaults["date_from"], _filter_defaults["date_to"]):
        _active.append(f"{date_from} → {date_to}")
    if nc_type != "All": _active.append(nc_type)
    if nc_status != "All": _active.append(nc_status)
    if nc_project: _active.append(f"{len(nc_project)} project(s)")
    if nc_owner: _active.append(f"{len(nc_owner)} owner(s)")
    if nc_source != "All": _active.append(nc_source)
    if _active:
        st.markdown(f'<div class="filter-active">Active: {" · ".join(_active)}</div>', unsafe_allow_html=True)


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
    if HAS_SOURCE and nc_source != "All":
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

# ---- Burndown filter: applies Project/Owner/NC-type/Status/Data-source (NOT the From/To period) ----
def _bd_filter(include_dates=True):
    cl, pr = [], []
    if include_dates and date_from:
        cl.append("created_on >= ?"); pr.append(str(date_from))
    if include_dates and date_to:
        cl.append("created_on <= ?"); pr.append(str(date_to) + " 23:59:59")
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
    if HAS_SOURCE and nc_source != "All":
        cl.append("source = ?"); pr.append(nc_source)
    return cl, pr

# Full filter (with dates) used across the whole dashboard.
_BF_CL, _BF_PR = _bd_filter(include_dates=True)
_filter_sig = f"{date_from}|{date_to}|{nc_type}|{nc_status}|{'-'.join(sorted(nc_project))}|{'-'.join(sorted(nc_owner))}|{nc_source}"
def _bd(where_parts, params):
    """Run a burndown count query with the active filter AND-ed in."""
    parts = list(where_parts) + _BF_CL
    prm = list(params) + _BF_PR
    sql = "SELECT COUNT(*) AS n FROM nc" + (" WHERE " + " AND ".join(parts) if parts else "")
    return _q(sql, prm).iloc[0]["n"]

backlog_at_start = _bd(["created_on < ?", "(is_open=1 OR closure_date >= ?)"], [es, es])
closed_from_backlog = _bd(["created_on < ?", "is_open=0", "closure_date >= ?"], [es, es])
still_open_backlog = _bd(["created_on < ?", "is_open=1"], [es])
new_since_start = _bd(["created_on >= ?"], [es])
new_still_open = _bd(["created_on >= ?", "is_open=1"], [es])
total_open = _bd(["is_open=1"], [])

today = date.today()
weeks_elapsed = max(1, (today - exercise_start).days / 7)
avg_new_wk = round(new_since_start / weeks_elapsed, 1)
team_size = 11

def _kpi_card(label, value, desc):
    return f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-desc">{desc}</div>
    </div>
    """

_es_str = exercise_start.strftime('%d %b %Y')
_kpis_row1 = [
    ("Backlog at freeze", int(backlog_at_start), f"NCs already open on {_es_str} — the starting pile."),
    ("Closed since start", int(closed_from_backlog), f"How many of that original {int(backlog_at_start)} you've since closed."),
    ("Still open (backlog)", int(still_open_backlog), f"Of the original {int(backlog_at_start)}, how many are still open."),
    ("Total open now", int(total_open), "All currently open NCs (old backlog + everything new)."),
]
_kpis_row2 = [
    ("New since start", int(new_since_start), f"NCs created after {_es_str}."),
    ("New still open", int(new_still_open), f"Of those {int(new_since_start)} new ones, how many are still open."),
    ("Avg new / week", f"{avg_new_wk}", "Inflow rate — new NCs arriving per week."),
]

for _row, _ncols in [(_kpis_row1, 4), (_kpis_row2, 3)]:
    _cols = st.columns(_ncols)
    for _c, (_lab, _val, _desc) in zip(_cols, _row):
        _c.markdown(_kpi_card(_lab, _val, _desc), unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────────
# Interactive "how many per week to reach ZERO open" — live deadline slider
# (this control is intentionally live/separate from the sidebar Apply filter)
# ──────────────────────────────────────────────────────────────────────────
st.markdown('<div class="reach-zero-anchor"></div>', unsafe_allow_html=True)
with st.container(border=True):
    st.markdown("##### 🎯 Closure rate to reach zero open")
    sc1, sc2, sc3 = st.columns([2, 1, 1])
    with sc1:
        target_deadline = st.slider(
            "Deadline", min_value=today,
            max_value=date(today.year + 2, 12, 31),
            value=date(2026, 9, 30), format="DD MMM YYYY",
            help="Drag to change the target date. The per-week / per-month numbers and the chart's target line update live.")
    with sc2:
        inflow = st.number_input(
            "New / week (inflow)", min_value=0.0, value=float(avg_new_wk), step=0.1,
            key=f"inflow_{_filter_sig}",
            help=f"Assumed new NCs arriving per week. Default is this selection's average ({avg_new_wk}). Adjust to stress-test.")
    with sc3:
        open_now_input = st.number_input(
            "Open now", min_value=0, value=int(total_open), step=1,
            key=f"opennow_{_filter_sig}",
            help="Current open NC count for this selection. Editable to test scenarios.")

    weeks_left = max(1, (target_deadline - today).days / 7)
    # Option B: clear all open + absorb ongoing inflow  ->  open/weeks + inflow
    required_per_week = open_now_input / weeks_left + inflow
    required_per_month = required_per_week * 4.33
    weekly_target = math.ceil(required_per_week)          # kept for downstream monthly-trend target line
    breakeven = math.ceil(inflow)

    zc1, zc2, zc3 = st.columns(3)
    zc1.metric(f"Close / week → 0 by {target_deadline.strftime('%d %b %Y')}",
               f"{math.ceil(required_per_week)} NCs",
               help=f"{open_now_input} open ÷ {weeks_left:.0f} weeks + {inflow:.1f}/wk inflow. "
                    f"With {team_size} people ≈ {max(1, round(required_per_week/team_size,1))} NC/person/week.")
    zc2.metric("Close / month", f"{math.ceil(required_per_month)} NCs",
               help="Weekly required × 4.33 weeks per month.")
    zc3.metric("Break-even (hold the line)", f"{breakeven} NCs/wk",
               help="Closing only this many just matches inflow — open count stays flat. Below this, it grows.")

    progress = closed_from_backlog / max(1, backlog_at_start)
    st.progress(min(progress, 1.0), text=f"Backlog closure: {int(closed_from_backlog)} / {int(backlog_at_start)} ({progress:.0%})")
    st.caption(f"Progress = closures from original backlog ÷ backlog at freeze. Does not include new NCs opened after {exercise_start.strftime('%d.%m.%Y')}.")

    # ══════════════════════════════════════════════════════════════════════════
    # Build the shared monthly series once (used by all three charts below)
    # ══════════════════════════════════════════════════════════════════════════
    _hist_months = pd.date_range(exercise_start.replace(day=1), today, freq="ME")
    _flt = (" AND " + " AND ".join(_BF_CL)) if _BF_CL else ""
    _rows = []
    for _m in _hist_months:
        _me = _m.strftime("%Y-%m-%d")
        _n = _q("SELECT COUNT(*) AS n FROM nc WHERE created_on <= ? AND (is_open=1 OR closure_date > ?)" + _flt,
                [_me, _me] + _BF_PR).iloc[0]["n"]
        _rows.append({"month": _m.strftime("%Y-%m"), "open": int(_n)})
    actual = pd.DataFrame(_rows)

    # opened & closed per month (in / out flow) — filter-aware
    flow = _q(f"""
        WITH o AS (SELECT substr(created_on,1,7) AS month, COUNT(*) AS opened
                   FROM nc WHERE created_on >= ?{_flt} GROUP BY month),
             c AS (SELECT substr(closure_date,1,7) AS month, COUNT(*) AS closed
                   FROM nc WHERE closure_date >= ?{_flt} GROUP BY month)
        SELECT COALESCE(o.month,c.month) AS month,
               COALESCE(o.opened,0) AS opened, COALESCE(c.closed,0) AS closed
        FROM o LEFT JOIN c ON o.month=c.month ORDER BY month
    """, [es] + _BF_PR + [es] + _BF_PR)

    if actual.empty:
        st.info("No data since the exercise start date.")
    else:
        # anchor last actual point to live open-now
        actual.loc[actual.index[-1], "open"] = int(open_now_input)
        last_month = actual["month"].iloc[-1]
        start_open = int(open_now_input)

        # recent actual close rate (last 12 weeks) for the prediction — filter-aware
        _cut = (today - pd.Timedelta(weeks=12)).isoformat()
        _closed_recent = _q("SELECT COUNT(*) n FROM nc WHERE is_open=0 AND closure_date>=?" + _flt,
                            [_cut] + _BF_PR).iloc[0]["n"]
        close_rate_wk = round(_closed_recent / 12, 1)

        # ──────────────────────────────────────────────────────────────────────
        # CHART 1 — DATA ANALYSIS: Open NCs per month (+ opened vs closed flow)
        # ──────────────────────────────────────────────────────────────────────
        st.markdown("**Open NCs per month** — what actually happened")
        fig1 = go.Figure()
        fig1.add_trace(go.Bar(x=flow["month"], y=flow["opened"], name="Opened", marker_color="#F26E21",
                              hovertemplate="<b>%{x}</b><br>Opened: %{y}<extra></extra>"))
        fig1.add_trace(go.Bar(x=flow["month"], y=flow["closed"], name="Closed", marker_color="#4CAF50",
                              hovertemplate="<b>%{x}</b><br>Closed: %{y}<extra></extra>"))
        fig1.add_trace(go.Scatter(x=actual["month"], y=actual["open"], name="Open at month-end",
                                  mode="lines+markers", line=dict(color="#1E2761", width=3), marker=dict(size=6),
                                  hovertemplate="<b>%{x}</b><br>Open: %{y}<extra></extra>"))
        fig1.update_layout(height=300, margin=dict(l=0, r=0, t=10, b=0), barmode="group",
                           legend=dict(orientation="h", y=-0.15), xaxis_title="", yaxis_title="NCs")
        st.plotly_chart(fig1, use_container_width=True, config=CHART_CONFIG)
        st.caption(f"Orange = NCs opened that month. Green = NCs closed that month. Navy line = total open at month-end. "
                   f"Recent pace: ~{inflow:.1f} opened/week vs ~{close_rate_wk:.1f} closed/week — "
                   f"{'closing faster than opening (backlog shrinks)' if close_rate_wk >= inflow else 'opening faster than closing (backlog grows)'}.")

        dl1, dl2, _ = st.columns([1, 1, 4])
        with dl1:
            st.download_button("📥 Excel", to_excel_bytes(actual, "Burndown"),
                               "burndown_monthly.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown("")


# ══════════════════════════════════════════════════════════════════════════════
# CURRENT MONTH
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div id="monthly"></div>', unsafe_allow_html=True)
cm = datetime.now().strftime("%Y-%m")
cm_label = datetime.now().strftime("%B %Y")
st.subheader(f"Current Month — {cm_label}")

# Current Month respects Project/Owner/type/status/source (dates are inherently the current month here)
_cm_cl, _cm_pr = _bd_filter(include_dates=False)
_cm_flt = (" AND " + " AND ".join(_cm_cl)) if _cm_cl else ""
cm_opened = _q("SELECT COUNT(*) AS n FROM nc WHERE substr(created_on,1,7)=?" + _cm_flt, [cm] + _cm_pr).iloc[0]["n"]
cm_closed = _q("SELECT COUNT(*) AS n FROM nc WHERE substr(closure_date,1,7)=?" + _cm_flt, [cm] + _cm_pr).iloc[0]["n"]

m1, m2, m3 = st.columns(3)
m1.metric("Opened this month", int(cm_opened),
          help=f"NCs with created_on in {cm_label}. New quality issues entering the queue.")
m2.metric("Closed this month", int(cm_closed),
          help=f"NCs with closure_date in {cm_label}. The team's closure output.")
m3.metric("Net (closed − opened)", int(cm_closed - cm_opened),
          delta_color="normal" if cm_closed >= cm_opened else "inverse",
          help="Closed minus opened. Positive = backlog shrinking. Negative = backlog growing. Target: always positive.")

st.markdown("**Monthly opens vs closes**")
_mo_flt = (" AND " + " AND ".join(_cm_cl)) if _cm_cl else ""
df_mo = _q(f"""
    WITH o AS (SELECT substr(created_on,1,7) AS month, COUNT(*) AS opened
               FROM nc WHERE created_on >= ?{_mo_flt} GROUP BY month),
         c AS (SELECT substr(closure_date,1,7) AS month, COUNT(*) AS closed
               FROM nc WHERE closure_date >= ?{_mo_flt} GROUP BY month)
    SELECT o.month, COALESCE(o.opened,0) AS opened, COALESCE(c.closed,0) AS closed
    FROM o LEFT JOIN c ON o.month=c.month ORDER BY o.month
""", [es] + _cm_pr + [es] + _cm_pr)

if not df_mo.empty:
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df_mo["month"], y=df_mo["opened"], name="Opened", marker_color="#F26E21",
                         hovertemplate="<b>%{x}</b><br>Opened: %{y} NCs<extra></extra>"))
    fig.add_trace(go.Bar(x=df_mo["month"], y=df_mo["closed"], name="Closed", marker_color="#4CAF50",
                         hovertemplate="<b>%{x}</b><br>Closed: %{y} NCs<extra></extra>"))
    fig.update_layout(height=280, margin=dict(l=0, r=0, t=10, b=0), barmode="group",
                      legend=dict(orientation="h", y=-0.15), xaxis_title="", yaxis_title="NCs")
    st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
    st.caption("Orange = new NCs opened per month. Green = NCs closed per month. When orange exceeds green, the backlog grows that month.")

    dl1, dl2, _ = st.columns([1, 1, 4])
    with dl1:
        st.download_button("📥 Excel", to_excel_bytes(df_mo, "Monthly_Trend"),
                           "monthly_trend.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_monthly")

st.markdown("**Opened vs Closed — monthly trend**")
df_trend = _q(f"""
    WITH o AS (SELECT substr(created_on,1,7) AS month, COUNT(*) AS opened
               FROM nc WHERE created_on IS NOT NULL{_mo_flt} GROUP BY month),
         c AS (SELECT substr(closure_date,1,7) AS month, COUNT(*) AS closed
               FROM nc WHERE closure_date IS NOT NULL{_mo_flt} GROUP BY month)
    SELECT COALESCE(o.month,c.month) AS month,
           COALESCE(o.opened,0) AS opened, COALESCE(c.closed,0) AS closed
    FROM o LEFT JOIN c ON o.month=c.month
    WHERE COALESCE(o.month,c.month) IS NOT NULL
    ORDER BY month
""", _cm_pr + _cm_pr)
if not df_trend.empty:
    fig_tr = go.Figure()
    fig_tr.add_trace(go.Scatter(x=df_trend["month"], y=df_trend["opened"], name="Opened",
                                mode="lines", line=dict(color="#F26E21", width=2),
                                hovertemplate="<b>%{x}</b><br>Opened: %{y}<extra></extra>"))
    fig_tr.add_trace(go.Scatter(x=df_trend["month"], y=df_trend["closed"], name="Closed",
                                mode="lines", line=dict(color="#4CAF50", width=2),
                                hovertemplate="<b>%{x}</b><br>Closed: %{y}<extra></extra>"))
    fig_tr.update_layout(height=280, margin=dict(l=0, r=0, t=10, b=0),
                         legend=dict(orientation="h", y=-0.15), xaxis_title="", yaxis_title="NCs")
    st.plotly_chart(fig_tr, use_container_width=True, config=CHART_CONFIG)
    st.caption("Opened (orange) and closed (green) NCs per month across the full dataset. "
               "Where the orange line sits above green, more opened than closed that month — backlog grew. Follows the active filters.")
    dlp1, dlp2, _ = st.columns([1, 1, 4])
    with dlp1:
        st.download_button("📥 Excel", to_excel_bytes(df_trend, "Opened_vs_Closed_Trend"),
                           "opened_vs_closed_trend.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_permonth")

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
          help="All NCs with status OPEN in the filtered period. Active workload for the quality team.")
c2.metric("NCs Closed", int(kpis["ncs_closed"] or 0),
          help="All NCs with status CLOSED in the filtered period.")
c3.metric("Major NCs Open", int(kpis["major_open"] or 0),
          help="Open NCs classified as Major. Require NRB disposition review and carry higher risk. Prioritize for closure.")
c4.metric("Production NCs", int(kpis["production_open"] or 0),
          help="Open NCs from internal production (Z3). Issues found during manufacturing, assembly, or testing.")
c5.metric("Supplier NCs", int(kpis["supplier_open"] or 0),
          help="Open NCs from supplier/procurement complaints (Z2). Incoming material or component issues.")

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
        st.caption("Projects with the most open NCs. '(no project)' = missing project assignment — needs data cleanup.")
        st.download_button("📥 Excel", to_excel_bytes(df_proj, "Projects_WIP"),
                           "projects_wip.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_proj_wip")

        # ---- Drill-down: pick a project → explode into its Flight Units ----
        _drill_projects = [p for p in df_proj["project"].tolist() if p != "(no project)"]
        if _drill_projects:
            picked = st.selectbox("🔍 Explode a project into Flight Units",
                                  ["— select —"] + _drill_projects, key="proj_drill")
            if picked and picked != "— select —":
                df_fu = _qf(
                    "SELECT COALESCE(flight_unit,'(not recorded)') AS flight_unit, COUNT(*) AS open_ncs "
                    "FROM nc {WHERE} GROUP BY flight_unit ORDER BY open_ncs DESC",
                    extra=[("is_open=1", []), ("COALESCE(project,'(no project)') = ?", [picked])])
                if not df_fu.empty:
                    recorded = df_fu[df_fu["flight_unit"] != "(not recorded)"]
                    if not recorded.empty:
                        figd = px.bar(df_fu, x="open_ncs", y="flight_unit", orientation="h",
                                      color_discrete_sequence=["#F26E21"], text="open_ncs")
                        figd.update_layout(height=max(180, 40 * len(df_fu)), margin=dict(l=0, r=0, t=10, b=0),
                                           showlegend=False, yaxis=dict(categoryorder="total ascending"),
                                           xaxis_title="", yaxis_title="")
                        figd.update_traces(textposition="outside",
                                           hovertemplate="<b>%{y}</b><br>Open NCs: %{x}<extra></extra>")
                        st.plotly_chart(figd, use_container_width=True, config=CHART_CONFIG,
                                        key="fu_drill_chart")
                        _n_blank = int(df_fu.loc[df_fu["flight_unit"] == "(not recorded)", "open_ncs"].sum()) \
                            if "(not recorded)" in df_fu["flight_unit"].values else 0
                        st.caption(f"Open NCs for **{picked}** by Flight Unit. "
                                   + (f"{_n_blank} NC(s) have no Flight Unit recorded yet." if _n_blank else
                                      "All open NCs have a Flight Unit recorded."))
                    else:
                        st.info(f"No Flight Unit recorded yet for {picked} — the field is blank for all its open NCs.")
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
        st.caption("Where NCs were detected. Red = missing detection area (data quality issue).")
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
    # Always show both categories
    _base = pd.DataFrame({"source": ["Production", "Supplier"]})
    df_ps = _base.merge(df_ps, on="source", how="left").fillna(0)
    df_ps[["open_ncs", "closed_ncs"]] = df_ps[["open_ncs", "closed_ncs"]].astype(int)
    _prod = df_ps[df_ps["source"] == "Production"].iloc[0]
    _supp = df_ps[df_ps["source"] == "Supplier"].iloc[0]

    pcol, scol = st.columns(2)
    pcol.metric("Production — open", int(_prod["open_ncs"]),
                help=f"Internal manufacturing NCs (Z3) currently open. {int(_prod['closed_ncs'])} closed in range.")
    scol.metric("Supplier — open", int(_supp["open_ncs"]),
                help=f"Procurement complaints (Z2) currently open. {int(_supp['closed_ncs'])} closed in range.")

    # Focused chart: OPEN NCs only (the workload) — closed totals are shown in the cards above.
    fig = go.Figure()
    fig.add_trace(go.Bar(x=["Production", "Supplier"],
                         y=[int(_prod["open_ncs"]), int(_supp["open_ncs"])],
                         marker_color=["#1E2761", "#F26E21"],
                         text=[int(_prod["open_ncs"]), int(_supp["open_ncs"])],
                         textposition="outside",
                         hovertemplate="<b>%{x}</b><br>Open: %{y}<extra></extra>"))
    fig.update_layout(height=220, margin=dict(l=0, r=0, t=20, b=0), showlegend=False,
                      xaxis_title="", yaxis_title="Open NCs")
    st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
    st.caption("Open NCs by origin — Production = internal manufacturing (Z3), Supplier = procurement complaints (Z2). "
               "Closed totals are in the cards above. Follows the active filters, including the date range.")
    st.download_button("📥 Excel", to_excel_bytes(df_ps, "Prod_vs_Supplier"),
                       "prod_vs_supplier.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="dl_ps")

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
        st.caption("Open NCs per owner with oldest NC age. High 'Oldest' values may need escalation.")
        st.download_button("📥 Excel", to_excel_bytes(df_own, "Owner_Workload"),
                           "owner_workload.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_own")
    else:
        st.info("No open NCs in range.")

with st.container(border=True):
    st.markdown("📅 **NC Total per Year** — full history *(ignores the From/To date filter)*")
    # Non-date filter only: obeys Project/Owner/type/status/source, NOT From/To
    _yr_cl, _yr_pr = _bd_filter(include_dates=False)
    _yr_where = ("WHERE " + " AND ".join(_yr_cl + ["created_on IS NOT NULL"])) if (_yr_cl or True) else ""
    df_yr = _q(f"SELECT substr(created_on,1,4) AS yr, COUNT(*) AS n FROM nc {_yr_where} GROUP BY yr ORDER BY yr", _yr_pr)
    fig = px.bar(df_yr, x="yr", y="n", color_discrete_sequence=["#1E2761"], text="n")
    fig.update_layout(height=240, margin=dict(l=0, r=0, t=10, b=0), showlegend=False, xaxis_title="", yaxis_title="")
    fig.update_traces(textposition="outside", hovertemplate="<b>%{x}</b><br>Total NCs: %{y}<extra></extra>")
    st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
    st.caption("Full-history NC volume by year. Obeys Project / Owner / NC type / Status / Data source, "
               "but **not** the From/To dates — so the long-term trend context stays visible while you filter.")
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
st.caption(f"**{len(incomplete)} NCs** with at least one missing field. Send this list to owners for cleanup.")
st.dataframe(incomplete, use_container_width=True, hide_index=True, height=300)

dl1, dl2, _ = st.columns([1, 1, 4])
with dl1:
    st.download_button("📥 Excel", to_excel_bytes(incomplete, "Data_Quality"),
                       "data_quality.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="dl_dq")

st.caption(f"Data source: `quality.db` · Last modified: "
           f"{datetime.fromtimestamp(Path(DB_FILE).stat().st_mtime).strftime('%d.%m.%Y %H:%M')}")