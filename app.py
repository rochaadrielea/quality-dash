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
    /* Month picker section box */
    div[data-testid="stVerticalBlockBorderWrapper"]:has(div.month-anchor) {
        background: #FFF6EC;
        border: 2px solid #F26E21 !important;
        border-radius: 12px;
        padding: 0.5rem 0.75rem;
    }
    /* CAPA coverage section box */
    div[data-testid="stVerticalBlockBorderWrapper"]:has(div.coverage-anchor) {
        background: #F3F0FA;
        border: 2px solid #5A63A0 !important;
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


def _has_column(col_name, table="nc"):
    """Check if a column exists in a table."""
    try:
        cols = _q(f"PRAGMA table_info({table})")
        return col_name in cols["name"].values
    except Exception:
        return False


def _filter_context_rows(include_since=False):
    """Context block describing when the report was taken and which filters were active.
    Reads the committed filter values from session_state, so every export self-documents."""
    ss = st.session_state
    rows = [
        ("Report generated", datetime.now().strftime("%d.%m.%Y %H:%M")),
        ("Data source", DB_FILE),
        ("Data last modified",
         datetime.fromtimestamp(Path(DB_FILE).stat().st_mtime).strftime("%d.%m.%Y %H:%M")
         if Path(DB_FILE).exists() else "n/a"),
        ("Period From", str(ss.get("date_from", ""))),
        ("Period To", str(ss.get("date_to", ""))),
    ]
    if include_since:
        rows.append(("Since (burndown anchor)", str(ss.get("since_date", ""))))
    rows += [
        ("NC type", str(ss.get("nc_type", "All"))),
        ("Status", str(ss.get("nc_status", "All"))),
        ("Project", ", ".join(ss.get("nc_project", [])) or "All"),
        ("Owner", ", ".join(ss.get("nc_owner", [])) or "All"),
        ("Data source filter", str(ss.get("nc_source", "All"))),
    ]
    return pd.DataFrame(rows, columns=["Field", "Value"])


def to_excel_bytes(df, sheet_name="Data", include_since=False):
    """Export a dataframe with a filter-context header block above the data."""
    buf = BytesIO()
    ctx = _filter_context_rows(include_since=include_since)
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # Context block at the top, data below it
        ctx.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=len(ctx) + 2)
        ws = writer.sheets[sheet_name]
        for _c in ("A", "B", "C", "D", "E", "F"):
            ws.column_dimensions[_c].width = 24
    return buf.getvalue()


def build_full_report_bytes(datasets, include_since=True):
    """One workbook: a Filters/Context sheet + one sheet per chart dataset."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        _filter_context_rows(include_since=include_since).to_excel(
            writer, sheet_name="Filters", index=False)
        writer.sheets["Filters"].column_dimensions["A"].width = 26
        writer.sheets["Filters"].column_dimensions["B"].width = 40
        for name, d in datasets.items():
            if d is None or (hasattr(d, "empty") and d.empty):
                continue
            sheet = name[:31]
            d.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            for _c in ("A", "B", "C", "D", "E", "F"):
                ws.column_dimensions[_c].width = 22
    return buf.getvalue()


CHART_CONFIG = {
    # 'hover' keeps the zoom/pan toolbar out of the plot area until you point at
    # the chart — as 'True' it renders on top of the tallest bars and hides them.
    "displayModeBar": "hover",
    "toImageButtonOptions": {
        "format": "png",
        "height": 600,
        "width": 1000,
        "scale": 2,
    },
    "displaylogo": False,
}

# Top margin reserved for the Plotly modebar so it never overlaps the top bar.
MODEBAR_T = 40

# Check schema once at startup
HAS_SOURCE = _has_column("source")


def _has_table(name):
    try:
        t = _q("SELECT name FROM sqlite_master WHERE type='table' AND name=?", [name])
        return not t.empty
    except Exception:
        return False


HAS_CAPA = _has_table("capa")
HAS_COPQ = _has_column("copq")
# capa_type is what makes RCA / CA / PA reporting possible. Older DBs built by
# the RCA-only ingest have no such column — coverage hides itself in that case.
HAS_CAPA_TYPE = HAS_CAPA and _has_column("capa_type", "capa")


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
        <a href="#capa-coverage" class="nav-link">CAPA Coverage</a>
        <a href="#root-cause" class="nav-link">Root Cause</a>
        <a href="#data-quality" class="nav-link">Data Quality</a>
    """, unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("### 🔍 Filters")
    month_label = st.text_input("Reporting month", value=datetime.now().strftime("%B %Y"))

    # Earliest NC in the data — used as the lower bound of the date pickers.
    _min_row = _q("SELECT MIN(created_on) AS m FROM nc WHERE created_on IS NOT NULL")
    _earliest = None
    if not _min_row.empty and _min_row.iloc[0]["m"]:
        try:
            _earliest = pd.to_datetime(_min_row.iloc[0]["m"]).date()
        except Exception:
            _earliest = None
    _pick_min = _earliest or date(2015, 1, 1)
    _pick_max = date(date.today().year + 2, 12, 31)
    # Default view starts in 2023 — older data exists and is still selectable,
    # but loading a decade of history makes the charts noisy.
    _default_from = max(_pick_min, date(2023, 1, 1))

    # ---- Defaults for the general filters (committed values live in session_state) ----
    _filter_defaults = {
        "date_from": _default_from,
        "date_to": date.today(),
        "since_date": date(2025, 6, 1),
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
        # Explicit bounds (_pick_min/_pick_max set above): without these, Streamlit only
        # offers ~10 years around the current value, which hid 2025/2026.
        fc1, fc2 = st.columns(2)
        with fc1:
            f_date_from = st.date_input("From", value=st.session_state["date_from"],
                                        min_value=_pick_min, max_value=_pick_max,
                                        key="w_date_from")
        with fc2:
            f_date_to = st.date_input("To", value=st.session_state["date_to"],
                                      min_value=_pick_min, max_value=_pick_max,
                                      key="w_date_to")

        f_since = st.date_input(
            "Since (burndown anchor)", value=st.session_state["since_date"],
            min_value=_pick_min, max_value=_pick_max,
            key="w_since",
            help="The date the burndown measures FROM. It affects only the burndown KPI numbers "
                 "(Backlog at freeze, Closed since start, Still open, New since start, New still open). "
                 "All charts follow the From/To window instead. Must sit inside From/To.")

        f_nc_type = st.selectbox("NC type", ["All", "Production", "Supplier"],
                                 index=["All", "Production", "Supplier"].index(st.session_state["nc_type"]),
                                 key="w_nc_type")
        f_nc_status = st.selectbox("Status", ["All", "Open", "Closed"],
                                   index=["All", "Open", "Closed"].index(st.session_state["nc_status"]),
                                   key="w_nc_status")
        f_nc_project = st.multiselect("Project", _projects,
                                      default=[p for p in st.session_state["nc_project"] if p in _projects],
                                      key="w_nc_project")
        f_nc_owner = st.multiselect("Owner", _owners,
                                    default=[o for o in st.session_state["nc_owner"] if o in _owners],
                                    key="w_nc_owner")
        if _sources:
            f_nc_source = st.selectbox("Data source", ["All"] + _sources,
                                       index=(["All"] + _sources).index(st.session_state["nc_source"])
                                       if st.session_state["nc_source"] in (["All"] + _sources) else 0,
                                       key="w_nc_source")
        else:
            f_nc_source = "All"

        c_apply, c_reset = st.columns(2)
        applied = c_apply.form_submit_button("✓ Apply filters", use_container_width=True, type="primary")
        reset = c_reset.form_submit_button("Reset", use_container_width=True)

    if applied:
        # Read the SUBMITTED widget values from their keys (authoritative on submit)
        f_date_from = st.session_state["w_date_from"]
        f_date_to = st.session_state["w_date_to"]
        f_since = st.session_state["w_since"]
        f_nc_type = st.session_state["w_nc_type"]
        f_nc_status = st.session_state["w_nc_status"]
        f_nc_project = st.session_state["w_nc_project"]
        f_nc_owner = st.session_state["w_nc_owner"]
        f_nc_source = st.session_state.get("w_nc_source", "All")

        if f_date_from > f_date_to:
            st.error(f"⚠️ 'From' ({f_date_from}) is after 'To' ({f_date_to}). "
                     "Pick a From date on or before the To date — filters not applied.")
        else:
            # Keep 'Since' inside the From/To window — pull it along rather than
            # rejecting the whole apply (Since must satisfy From <= Since <= To).
            _since_adj = f_since
            _since_moved = False
            if _since_adj < f_date_from:
                _since_adj = f_date_from
                _since_moved = True
            elif _since_adj > f_date_to:
                _since_adj = f_date_to
                _since_moved = True

            st.session_state["date_from"] = f_date_from
            st.session_state["date_to"] = f_date_to
            st.session_state["since_date"] = _since_adj
            st.session_state["nc_type"] = f_nc_type
            st.session_state["nc_status"] = f_nc_status
            st.session_state["nc_project"] = f_nc_project
            st.session_state["nc_owner"] = f_nc_owner
            st.session_state["nc_source"] = f_nc_source
            # Let the Since widget re-initialise from the (possibly adjusted) value
            if _since_moved:
                st.session_state.pop("w_since", None)
                st.session_state["_since_notice"] = (
                    f"'Since' moved to {_since_adj} to stay inside the From/To window.")
            else:
                st.session_state.pop("_since_notice", None)
            st.rerun()

    if reset:
        for _k, _v in _filter_defaults.items():
            st.session_state[_k] = _v
        # Clear the form widget keys so they re-init from defaults
        for _wk in ["w_date_from", "w_date_to", "w_since", "w_nc_type",
                    "w_nc_status", "w_nc_project", "w_nc_owner", "w_nc_source"]:
            st.session_state.pop(_wk, None)
        st.rerun()

    # Read committed filter values (used by build_where downstream)
    if st.session_state.get("_since_notice"):
        st.info("⚓ " + st.session_state["_since_notice"])
    date_from = st.session_state["date_from"]
    date_to = st.session_state["date_to"]
    # Safety: never let an inverted range reach the queries
    if date_from > date_to:
        date_from, date_to = date_to, date_from
    # 'Since' anchors the burndown; keep it inside the From/To window
    exercise_start = st.session_state["since_date"]
    exercise_start = max(date_from, min(exercise_start, date_to))
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
st.caption(f"⚓ **Since {exercise_start.strftime('%d %b %Y')}** = the date the burndown measures from. "
           f"It drives **only the KPI numbers below**. The three *backlog* figures (Backlog at freeze, "
           f"Closed since start, Still open) describe NCs created **before** that date, so they are not "
           f"clipped by the From date — but they do follow the Project / Owner / NC-type filters. "
           f"Every chart follows the **From/To** window (**{date_from}** → **{date_to}**).")

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
# Non-date filter — for the BACKLOG KPIs, which describe NCs created BEFORE the
# 'Since' anchor. Applying "created_on >= From" to them is self-contradictory
# (it asks for NCs both before Since and after From) and always yields 0.
_NB_CL, _NB_PR = _bd_filter(include_dates=False)
_filter_sig = f"{date_from}|{date_to}|{nc_type}|{nc_status}|{'-'.join(sorted(nc_project))}|{'-'.join(sorted(nc_owner))}|{nc_source}"

def _bd(where_parts, params, use_dates=True):
    """Run a burndown count query with the active filter AND-ed in.
    use_dates=False for backlog metrics that look before the From date."""
    _cl, _pr = (_BF_CL, _BF_PR) if use_dates else (_NB_CL, _NB_PR)
    parts = list(where_parts) + _cl
    prm = list(params) + _pr
    sql = "SELECT COUNT(*) AS n FROM nc" + (" WHERE " + " AND ".join(parts) if parts else "")
    return _q(sql, prm).iloc[0]["n"]

# Backlog metrics look BEFORE 'Since', so they must not be clipped by the From date.
backlog_at_start = _bd(["created_on < ?", "(is_open=1 OR closure_date >= ?)"], [es, es], use_dates=False)
closed_from_backlog = _bd(["created_on < ?", "is_open=0", "closure_date >= ?"], [es, es], use_dates=False)
still_open_backlog = _bd(["created_on < ?", "is_open=1"], [es], use_dates=False)
# These live inside the window, so they follow the full filter.
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
    st.markdown("##### 🎯 Closure rate to reach target")
    sc1, sc2, sc3, sc4 = st.columns([2, 1, 1, 1])
    with sc1:
        target_deadline = st.slider(
            "Deadline", min_value=today,
            max_value=date(today.year + 2, 12, 31),
            value=date(2026, 9, 30), format="DD MMM YYYY",
            help="Drag to change the target date. The per-week / per-month numbers update live.")
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
    with sc4:
        target_open = st.number_input(
            "Target open", min_value=0, value=65, step=1,
            key="target_open",
            help="The acceptable number of open NCs to get down to by the deadline (not necessarily zero). "
                 "Change it here if the goal changes.")

    weeks_left = max(1, (target_deadline - today).days / 7)
    # Close the gap down to the target AND absorb ongoing inflow:
    #   (open - target) / weeks + inflow
    _gap = open_now_input - target_open
    _at_target = _gap <= 0
    required_per_week = (max(0, _gap) / weeks_left) + inflow
    required_per_month = required_per_week * 4.33
    weekly_target = math.ceil(required_per_week)          # kept for downstream use
    breakeven = math.ceil(inflow)

    zc1, zc2, zc3 = st.columns(3)
    if _at_target:
        zc1.metric(f"Close / week → {target_open} by {target_deadline.strftime('%d %b %Y')}",
                   f"{breakeven} NCs",
                   help=f"Already at or below the target ({open_now_input} open ≤ {target_open}). "
                        f"Closing {breakeven}/week just matches inflow and holds the line.")
    else:
        zc1.metric(f"Close / week → {target_open} by {target_deadline.strftime('%d %b %Y')}",
                   f"{math.ceil(required_per_week)} NCs",
                   help=f"({open_now_input} open − {target_open} target) ÷ {weeks_left:.0f} weeks "
                        f"+ {inflow:.1f}/wk inflow. With {team_size} people ≈ "
                        f"{max(1, round(required_per_week/team_size,1))} NC/person/week.")
    zc2.metric("Close / month", f"{math.ceil(required_per_month)} NCs",
               help="Weekly required × 4.33 weeks per month.")
    zc3.metric("Break-even (hold the line)", f"{breakeven} NCs/wk",
               help="Closing only this many just matches inflow — open count stays flat. Below this, it grows.")

    if _at_target:
        st.success(f"✅ Already at target — {open_now_input} open is at or below the target of {target_open}. "
                   f"Keep closing ~{breakeven}/week to hold it.")
    else:
        st.caption(f"Need to clear **{_gap}** NCs ({open_now_input} open → {target_open} target) "
                   f"in **{weeks_left:.0f}** weeks, while ~{inflow:.1f} new NCs arrive each week.")

    progress = closed_from_backlog / max(1, backlog_at_start)
    st.progress(min(progress, 1.0), text=f"Backlog closure: {int(closed_from_backlog)} / {int(backlog_at_start)} ({progress:.0%})")
    st.caption(f"Progress = closures from original backlog ÷ backlog at freeze. Does not include new NCs opened after {exercise_start.strftime('%d.%m.%Y')}.")

    # ══════════════════════════════════════════════════════════════════════════
    # Build the shared monthly series once (used by all three charts below)
    # ══════════════════════════════════════════════════════════════════════════
    _hist_months = pd.date_range(date_from.replace(day=1), min(date_to, today), freq="ME")
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
                   FROM nc WHERE created_on IS NOT NULL{_flt} GROUP BY month),
             c AS (SELECT substr(closure_date,1,7) AS month, COUNT(*) AS closed
                   FROM nc WHERE closure_date IS NOT NULL{_flt} GROUP BY month)
        SELECT COALESCE(o.month,c.month) AS month,
               COALESCE(o.opened,0) AS opened, COALESCE(c.closed,0) AS closed
        FROM o LEFT JOIN c ON o.month=c.month
        WHERE COALESCE(o.month,c.month) IS NOT NULL
        ORDER BY month
    """, _BF_PR + _BF_PR)

    if actual.empty:
        st.info("No data in the selected From/To range.")
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
        st.caption(f"📅 Showing **{date_from}** → **{date_to}**")
        fig1 = go.Figure()
        fig1.add_trace(go.Bar(x=flow["month"], y=flow["opened"], name="Opened", marker_color="#F26E21",
                              hovertemplate="<b>%{x}</b><br>Opened: %{y}<extra></extra>"))
        fig1.add_trace(go.Bar(x=flow["month"], y=flow["closed"], name="Closed", marker_color="#4CAF50",
                              hovertemplate="<b>%{x}</b><br>Closed: %{y}<extra></extra>"))
        fig1.add_trace(go.Scatter(x=actual["month"], y=actual["open"], name="Open at month-end",
                                  mode="lines+markers", line=dict(color="#1E2761", width=3), marker=dict(size=6),
                                  hovertemplate="<b>%{x}</b><br>Open: %{y}<extra></extra>"))
        fig1.update_layout(height=300, margin=dict(l=0, r=0, t=MODEBAR_T, b=0), barmode="group",
                           legend=dict(orientation="h", y=-0.15), xaxis_title="", yaxis_title="NCs")
        st.plotly_chart(fig1, use_container_width=True, config=CHART_CONFIG)
        st.caption(f"Orange = NCs opened that month. Green = NCs closed that month. Navy line = total open at month-end. "
                   f"Recent pace: ~{inflow:.1f} opened/week vs ~{close_rate_wk:.1f} closed/week — "
                   f"{'closing faster than opening (backlog shrinks)' if close_rate_wk >= inflow else 'opening faster than closing (backlog grows)'}.")

        dl1, dl2, _ = st.columns([1, 1, 4])
        with dl1:
            st.download_button("📥 Excel", to_excel_bytes(actual, "Burndown", include_since=True),
                               "burndown_monthly.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown("")


# ══════════════════════════════════════════════════════════════════════════════
# CURRENT MONTH
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div id="monthly"></div>', unsafe_allow_html=True)

# Current Month obeys the full filter (dates + project/owner/type/status/source)
_cm_cl, _cm_pr = _bd_filter(include_dates=True)
_cm_flt = (" AND " + " AND ".join(_cm_cl)) if _cm_cl else ""

st.markdown('<div class="month-anchor"></div>', unsafe_allow_html=True)
with st.container(border=True):
    # Months available inside the active From/To window (newest first)
    _months_df = _q(
        "SELECT DISTINCT substr(created_on,1,7) AS m FROM nc WHERE created_on IS NOT NULL"
        + _cm_flt + " ORDER BY m DESC", _cm_pr)
    _month_opts = [m for m in _months_df["m"].tolist() if m]
    _this_month = datetime.now().strftime("%Y-%m")
    if _this_month not in _month_opts:
        _month_opts = [_this_month] + _month_opts
    _default_idx = _month_opts.index(_this_month) if _this_month in _month_opts else 0

    hcol, scol = st.columns([2, 1])
    with scol:
        cm = st.selectbox("Month", _month_opts, index=_default_idx, key="month_pick",
                          help="Pick any month inside the From/To window. Defaults to the current month.")
    cm_label = datetime.strptime(cm, "%Y-%m").strftime("%B %Y")
    with hcol:
        _is_current = (cm == _this_month)
        st.subheader(("Current Month — " if _is_current else "Month — ") + cm_label)

    cm_opened = _q("SELECT COUNT(*) AS n FROM nc WHERE substr(created_on,1,7)=?" + _cm_flt, [cm] + _cm_pr).iloc[0]["n"]
    cm_closed = _q("SELECT COUNT(*) AS n FROM nc WHERE substr(closure_date,1,7)=?" + _cm_flt, [cm] + _cm_pr).iloc[0]["n"]

    m1, m2, m3 = st.columns(3)
    m1.metric("Opened", int(cm_opened),
              help=f"NCs created in {cm_label}. New quality issues entering the queue.")
    m2.metric("Closed", int(cm_closed),
              help=f"NCs closed in {cm_label}. The team's closure output.")
    m3.metric("Net (closed − opened)", int(cm_closed - cm_opened),
              delta_color="normal" if cm_closed >= cm_opened else "inverse",
              help="Closed minus opened. Positive = backlog shrinking. Negative = backlog growing.")

st.markdown("**Monthly opens vs closes**")
st.caption(f"📅 Showing **{date_from}** → **{date_to}**")
_mo_flt = (" AND " + " AND ".join(_cm_cl)) if _cm_cl else ""
df_mo = _q(f"""
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

if not df_mo.empty:
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df_mo["month"], y=df_mo["opened"], name="Opened", marker_color="#F26E21",
                         hovertemplate="<b>%{x}</b><br>Opened: %{y} NCs<extra></extra>"))
    fig.add_trace(go.Bar(x=df_mo["month"], y=df_mo["closed"], name="Closed", marker_color="#4CAF50",
                         hovertemplate="<b>%{x}</b><br>Closed: %{y} NCs<extra></extra>"))
    fig.update_layout(height=280, margin=dict(l=0, r=0, t=MODEBAR_T, b=0), barmode="group",
                      legend=dict(orientation="h", y=-0.15), xaxis_title="", yaxis_title="NCs")
    st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
    st.caption("Orange = new NCs opened per month. Green = NCs closed per month. When orange exceeds green, the backlog grows that month.")

    dl1, dl2, _ = st.columns([1, 1, 4])
    with dl1:
        st.download_button("📥 Excel", to_excel_bytes(df_mo, "Monthly_Trend"),
                           "monthly_trend.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_monthly")

st.markdown("**Opened vs Closed — monthly trend**")
st.caption(f"📅 Showing **{date_from}** → **{date_to}**")
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
    fig_tr.update_layout(height=280, margin=dict(l=0, r=0, t=MODEBAR_T, b=0),
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
          help="OPEN NCs in the filtered period. Active workload. Note: open NCs are mostly recent, "
               "so widening the date range adds few — old NCs are usually already closed.")
c2.metric("NCs Closed", int(kpis["ncs_closed"] or 0),
          help="CLOSED NCs in the filtered period. This grows a lot when you widen the date range.")
c3.metric("Major NCs Open", int(kpis["major_open"] or 0),
          help="OPEN NCs classified as Major. Require NRB disposition review and carry higher risk.")
c4.metric("Production NCs — open", int(kpis["production_open"] or 0),
          help="OPEN NCs from internal production (Z3). Issues found during manufacturing, assembly, or testing.")
c5.metric("Supplier NCs — open", int(kpis["supplier_open"] or 0),
          help="OPEN NCs from supplier/procurement complaints (Z2). Incoming material or component issues.")
st.caption("ℹ️ **Open** = currently unresolved (WIP). **Closed** = resolved. Open counts change little "
           "when you widen the dates because old NCs are mostly closed; closed counts change a lot.")

st.markdown("")


# ══════════════════════════════════════════════════════════════════════════════
# TRENDS & BREAKDOWN (filtered)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div id="trends"></div>', unsafe_allow_html=True)
st.subheader("Trends & Breakdown")

col_a, col_b = st.columns(2)

with col_a:
    st.markdown("**Top 6 Projects — NCs WIP** · :grey[open only]")
    st.caption(f"📅 {date_from} → {date_to}")
    df_proj = _qf("SELECT COALESCE(project,'(no project)') AS project, COUNT(*) AS open_ncs FROM nc {WHERE} GROUP BY project ORDER BY open_ncs DESC LIMIT 6",
                   extra=[("is_open=1", [])])
    if not df_proj.empty:
        fig = px.bar(df_proj, x="open_ncs", y="project", orientation="h",
                     color_discrete_sequence=["#1E2761"], text="open_ncs")
        fig.update_layout(height=260, margin=dict(l=0, r=0, t=MODEBAR_T, b=0), showlegend=False,
                          yaxis=dict(categoryorder="total ascending"), xaxis_title="", yaxis_title="")
        fig.update_traces(textposition="outside", hovertemplate="<b>%{y}</b><br>Open NCs: %{x}<extra></extra>")
        st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
        st.caption("**Counts OPEN NCs only** (WIP = current workload). Closed NCs are excluded, so widening "
                   "the date range adds little here — old NCs are mostly closed already. "
                   "'(no project)' = missing project assignment, needs data cleanup.")
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
                        figd.update_layout(height=max(180, 40 * len(df_fu)), margin=dict(l=0, r=0, t=MODEBAR_T, b=0),
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
    st.markdown("**NCs by Detection Area** · :grey[all NCs: open + closed]")
    st.caption(f"📅 {date_from} → {date_to}")
    df_area = _qf("SELECT COALESCE(detection_area,'BLANK - to clean') AS area, COUNT(*) AS n FROM nc {WHERE} GROUP BY area ORDER BY n DESC LIMIT 12")
    if not df_area.empty:
        colors = ["#E53E3E" if a.startswith("BLANK") else "#F26E21" for a in df_area["area"]]
        fig = px.bar(df_area, x="n", y="area", orientation="h", color=df_area["area"],
                     color_discrete_sequence=colors, text="n")
        fig.update_layout(height=260, margin=dict(l=0, r=0, t=MODEBAR_T, b=0), showlegend=False,
                          yaxis=dict(categoryorder="total ascending"), xaxis_title="", yaxis_title="")
        fig.update_traces(textposition="outside", hovertemplate="<b>%{y}</b><br>NCs: %{x}<extra></extra>")
        st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
        st.caption("**Counts ALL NCs — open and closed** (where each NC was first detected). "
                   "Because it includes closed NCs, this chart moves a lot when you widen the date range. "
                   "Red = missing detection area (data quality issue).")
        st.download_button("📥 Excel", to_excel_bytes(df_area, "Detection_Area"),
                           "detection_area.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_area")
    else:
        st.info("No NCs in range.")

col_c, col_d = st.columns(2)

with col_c:
    st.markdown("**Production vs Supplier** · :grey[open only (closed in tooltips)]")
    st.caption(f"📅 {date_from} → {date_to}")
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
    fig.update_layout(height=220, margin=dict(l=0, r=0, t=MODEBAR_T, b=0), showlegend=False,
                      xaxis_title="", yaxis_title="Open NCs")
    st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
    st.caption("Open NCs by origin — Production = internal manufacturing (Z3), Supplier = procurement complaints (Z2). "
               "Closed totals are in the cards above. Follows the active filters, including the date range.")
    st.download_button("📥 Excel", to_excel_bytes(df_ps, "Prod_vs_Supplier"),
                       "prod_vs_supplier.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="dl_ps")

with col_d:
    st.markdown("**Open NCs by Owner** (Top 10) · :grey[open only]")
    st.caption(f"📅 {date_from} → {date_to}")
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
    fig.update_layout(height=240, margin=dict(l=0, r=0, t=MODEBAR_T, b=0), showlegend=False, xaxis_title="", yaxis_title="")
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
# CAPA COVERAGE — does every NC have an RCA, a CA and a PA?
# ══════════════════════════════════════════════════════════════════════════════
df_cov = None
df_combo = None
if HAS_CAPA_TYPE:
    st.markdown('<div id="capa-coverage"></div>', unsafe_allow_html=True)
    st.subheader("CAPA Coverage")

    # Denominator = EVERY NC matching the active filter, not just the ones that
    # already appear in the CAPA tracker. An NC with no CAPA row at all is the
    # gap this section exists to surface, so it has to be inside the total.
    _cov_flt = (" AND " + " AND ".join(_BF_CL)) if _BF_CL else ""

    _cov_sql = f"""
        WITH scope AS (
            SELECT nc_id, is_open FROM nc WHERE 1=1{_cov_flt}
        )
        SELECT
            COUNT(*) AS n_total,
            SUM(CASE WHEN is_open=1 THEN 1 ELSE 0 END) AS n_open,
            SUM(CASE WHEN is_open=0 THEN 1 ELSE 0 END) AS n_closed,
            SUM(CASE WHEN rca THEN 1 ELSE 0 END) AS rca_all,
            SUM(CASE WHEN rca AND is_open=1 THEN 1 ELSE 0 END) AS rca_open,
            SUM(CASE WHEN rca AND is_open=0 THEN 1 ELSE 0 END) AS rca_closed,
            SUM(CASE WHEN ca THEN 1 ELSE 0 END) AS ca_all,
            SUM(CASE WHEN ca AND is_open=1 THEN 1 ELSE 0 END) AS ca_open,
            SUM(CASE WHEN ca AND is_open=0 THEN 1 ELSE 0 END) AS ca_closed,
            SUM(CASE WHEN pa THEN 1 ELSE 0 END) AS pa_all,
            SUM(CASE WHEN pa AND is_open=1 THEN 1 ELSE 0 END) AS pa_open,
            SUM(CASE WHEN pa AND is_open=0 THEN 1 ELSE 0 END) AS pa_closed,
            SUM(CASE WHEN NOT rca AND NOT ca AND NOT pa THEN 1 ELSE 0 END) AS none_all,
            SUM(CASE WHEN NOT rca AND NOT ca AND NOT pa AND is_open=1 THEN 1 ELSE 0 END) AS none_open,
            SUM(CASE WHEN NOT rca AND NOT ca AND NOT pa AND is_open=0 THEN 1 ELSE 0 END) AS none_closed
        FROM (
            SELECT s.is_open,
                EXISTS(SELECT 1 FROM capa c WHERE c.nc_id=s.nc_id AND c.capa_type='RCA') AS rca,
                EXISTS(SELECT 1 FROM capa c WHERE c.nc_id=s.nc_id AND c.capa_type='CA')  AS ca,
                EXISTS(SELECT 1 FROM capa c WHERE c.nc_id=s.nc_id AND c.capa_type='PA')  AS pa
            FROM scope s
        )
    """
    _cov = _q(_cov_sql, _BF_PR).iloc[0]

    _n_total = int(_cov["n_total"] or 0)
    _n_open = int(_cov["n_open"] or 0)
    _n_closed = int(_cov["n_closed"] or 0)

    if _n_total == 0:
        st.info("No NCs in the selected range.")
    else:
        st.markdown('<div class="coverage-anchor"></div>', unsafe_allow_html=True)
        with st.container(border=True):
            view = st.radio(
                "View", ["All NCs", "Open only", "Closed only", "Side by side"],
                horizontal=True, key="cov_view",
                help="The denominator is always every NC matching the sidebar filters — an NC with no "
                     "CAPA row at all is counted as not covered. Open/Closed split the same total; they "
                     "do not change it.")

            _suffix = {"All NCs": "all", "Open only": "open", "Closed only": "closed"}
            _denoms = {"all": _n_total, "open": _n_open, "closed": _n_closed}

            def _cov_row(key):
                """(label, count, denominator) for each CAPA type in one view."""
                d = _denoms[key]
                return [
                    ("RCA", int(_cov[f"rca_{key}"] or 0), d),
                    ("CA", int(_cov[f"ca_{key}"] or 0), d),
                    ("PA", int(_cov[f"pa_{key}"] or 0), d),
                    ("No CAPA at all", int(_cov[f"none_{key}"] or 0), d),
                ]

            _descs = {
                "RCA": "NCs with a root cause analysis recorded in the CAPA tracker.",
                "CA": "NCs with a corrective action recorded.",
                # Deliberately neutral: whether every NC *needs* a PA is still an
                # open question on the process side, so this doesn't imply a target.
                "PA": "NCs with a preventive action recorded.",
                "No CAPA at all": "NCs with no RCA, CA or PA row — nothing recorded.",
            }

            if view == "Side by side":
                st.caption(f"**{_n_open:,} open** · **{_n_closed:,} closed** · {_n_total:,} total NCs in range")
                _rows = []
                for _lab in ["RCA", "CA", "PA", "No CAPA at all"]:
                    _o = int(_cov[f"{_lab.split()[0].lower() if _lab != 'No CAPA at all' else 'none'}_open"] or 0)
                    _c = int(_cov[f"{_lab.split()[0].lower() if _lab != 'No CAPA at all' else 'none'}_closed"] or 0)
                    _rows.append({
                        "CAPA type": _lab,
                        "Open — with": _o,
                        "Open — %": round(100 * _o / _n_open, 1) if _n_open else 0.0,
                        "Closed — with": _c,
                        "Closed — %": round(100 * _c / _n_closed, 1) if _n_closed else 0.0,
                    })
                df_cov = pd.DataFrame(_rows)

                figc = go.Figure()
                figc.add_trace(go.Bar(
                    x=df_cov["CAPA type"], y=df_cov["Open — %"], name=f"Open (n={_n_open:,})",
                    marker_color="#F26E21", text=[f"{v}%" for v in df_cov["Open — %"]],
                    textposition="outside",
                    hovertemplate="<b>%{x}</b><br>Open with: %{customdata}<br>%{y}% of open<extra></extra>",
                    customdata=df_cov["Open — with"]))
                figc.add_trace(go.Bar(
                    x=df_cov["CAPA type"], y=df_cov["Closed — %"], name=f"Closed (n={_n_closed:,})",
                    marker_color="#1E2761", text=[f"{v}%" for v in df_cov["Closed — %"]],
                    textposition="outside",
                    hovertemplate="<b>%{x}</b><br>Closed with: %{customdata}<br>%{y}% of closed<extra></extra>",
                    customdata=df_cov["Closed — with"]))
                figc.update_layout(height=320, margin=dict(l=0, r=0, t=MODEBAR_T, b=0), barmode="group",
                                   legend=dict(orientation="h", y=-0.15),
                                   xaxis_title="", yaxis_title="% of NCs",
                                   yaxis=dict(range=[0, 105]))
                st.plotly_chart(figc, use_container_width=True, config=CHART_CONFIG, key="cov_side")
                st.caption("Each bar is a % of its **own** group (open, or closed) — the two groups have "
                           "different denominators, so compare the shape, not the height. A newly opened NC "
                           "has had no time to get an RCA yet, which is why open sits lower than closed.")
                st.dataframe(df_cov, use_container_width=True, hide_index=True)

            else:
                _key = _suffix[view]
                _d = _denoms[_key]
                _scope_txt = {"all": "every NC", "open": "open NCs", "closed": "closed NCs"}[_key]
                st.caption(f"Denominator: **{_d:,}** {_scope_txt} in range "
                           f"({_n_open:,} open · {_n_closed:,} closed · {_n_total:,} total).")

                if _d == 0:
                    st.info(f"No {_scope_txt} in the selected range.")
                    df_cov = pd.DataFrame()
                else:
                    _cards = _cov_row(_key)
                    _cols = st.columns(4)
                    for _c_, (_lab, _n, _dd) in zip(_cols, _cards):
                        _pct = (100 * _n / _dd) if _dd else 0
                        _c_.markdown(_kpi_card(
                            f"NCs with {_lab}" if _lab != "No CAPA at all" else _lab,
                            f"{_pct:.0f}%",
                            f"{_n:,} of {_dd:,} — {_descs[_lab]}"), unsafe_allow_html=True)

                    df_cov = pd.DataFrame([
                        {"CAPA type": _lab, "NCs with": _n, "Of total": _dd,
                         "%": round(100 * _n / _dd, 1) if _dd else 0.0}
                        for _lab, _n, _dd in _cards])

                    figc = px.bar(df_cov, x="CAPA type", y="%",
                                  color_discrete_sequence=["#5A63A0"],
                                  text=[f"{v}%" for v in df_cov["%"]])
                    figc.update_layout(height=300, margin=dict(l=0, r=0, t=MODEBAR_T, b=0),
                                       showlegend=False, xaxis_title="", yaxis_title="% of NCs",
                                       yaxis=dict(range=[0, 105]))
                    figc.update_traces(textposition="outside",
                                       hovertemplate="<b>%{x}</b><br>%{y}% of NCs<extra></extra>")
                    st.plotly_chart(figc, use_container_width=True, config=CHART_CONFIG, key="cov_single")
                    st.caption(f"Share of the {_d:,} {_scope_txt} in range that carry each CAPA type. "
                               "An NC missing from the CAPA tracker entirely counts as not covered.")

                st.download_button("📥 Excel", to_excel_bytes(df_cov, "CAPA_Coverage"),
                                   "capa_coverage.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key="dl_cov")

        # ---- Which combination does each NC have? ----
        st.markdown("**What each NC actually has**")
        _combo_sql = f"""
            WITH scope AS (SELECT nc_id, is_open FROM nc WHERE 1=1{_cov_flt}),
            flags AS (
                SELECT s.is_open,
                    EXISTS(SELECT 1 FROM capa c WHERE c.nc_id=s.nc_id AND c.capa_type='RCA') AS rca,
                    EXISTS(SELECT 1 FROM capa c WHERE c.nc_id=s.nc_id AND c.capa_type='CA')  AS ca,
                    EXISTS(SELECT 1 FROM capa c WHERE c.nc_id=s.nc_id AND c.capa_type='PA')  AS pa
                FROM scope s
            )
            SELECT
                CASE
                    WHEN rca AND ca AND pa THEN 'RCA + CA + PA'
                    WHEN rca AND ca THEN 'RCA + CA'
                    WHEN rca AND pa THEN 'RCA + PA'
                    WHEN ca AND pa THEN 'CA + PA'
                    WHEN rca THEN 'RCA only'
                    WHEN ca THEN 'CA only'
                    WHEN pa THEN 'PA only'
                    ELSE 'Nothing recorded'
                END AS combo,
                COUNT(*) AS n,
                SUM(CASE WHEN is_open=1 THEN 1 ELSE 0 END) AS n_open,
                SUM(CASE WHEN is_open=0 THEN 1 ELSE 0 END) AS n_closed
            FROM flags GROUP BY combo ORDER BY n DESC
        """
        df_combo = _q(_combo_sql, _BF_PR)
        if not df_combo.empty:
            _cmb = df_combo.copy()
            if view == "Open only":
                _cmb = _cmb[["combo", "n_open"]].rename(columns={"n_open": "n"})
                _cmb = _cmb[_cmb["n"] > 0]
            elif view == "Closed only":
                _cmb = _cmb[["combo", "n_closed"]].rename(columns={"n_closed": "n"})
                _cmb = _cmb[_cmb["n"] > 0]

            if _cmb.empty:
                st.info("Nothing to show for this view.")
            else:
                _cmb_colors = ["#E53E3E" if c == "Nothing recorded" else "#5A63A0"
                               for c in _cmb["combo"]]
                figk = px.bar(_cmb, x="n", y="combo", orientation="h",
                              color=_cmb["combo"], color_discrete_sequence=_cmb_colors, text="n")
                figk.update_layout(height=max(220, 40 * len(_cmb)),
                                   margin=dict(l=0, r=0, t=MODEBAR_T, b=0), showlegend=False,
                                   yaxis=dict(categoryorder="total ascending"),
                                   xaxis_title="NCs", yaxis_title="")
                figk.update_traces(textposition="outside",
                                   hovertemplate="<b>%{y}</b><br>NCs: %{x}<extra></extra>")
                st.plotly_chart(figk, use_container_width=True, config=CHART_CONFIG, key="cov_combo")
                st.caption("Every NC in range falls into exactly one bar. Red = nothing recorded at all. "
                           "This is the same population as the cards above, cut a different way.")
                st.download_button("📥 Excel", to_excel_bytes(df_combo, "CAPA_Combinations"),
                                   "capa_combinations.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key="dl_combo")

    st.markdown("")


# ══════════════════════════════════════════════════════════════════════════════
# ROOT CAUSE (CAPA / RCA tracker)
# ══════════════════════════════════════════════════════════════════════════════
if HAS_CAPA:
    st.markdown('<div id="root-cause"></div>', unsafe_allow_html=True)
    st.subheader("Root Cause Analysis")
    st.caption(f"📅 {date_from} → {date_to} · From the CAPA/RCA tracker, joined to NCs on NC number. "
               "**RCA rows only** — Origin Area and RC Category are recorded on the RCA row; CA and PA "
               "rows carry the action, not the cause. Blank / N/A / 0 are excluded, so these charts show "
               "only NCs where a root cause was actually written down.")

    # CAPA joined to nc so the sidebar filters apply
    _rc_flt = (" AND " + " AND ".join(_BF_CL)) if _BF_CL else ""
    # Restrict to RCA rows — with all three types now in the table, an unqualified
    # join would count each NC up to 3 times.
    _rca_only = " AND c.capa_type='RCA'" if HAS_CAPA_TYPE else ""

    # Denominator: NCs that have an RCA row at all. The "of every NC" version of
    # this question is the CAPA Coverage section above; here the useful question
    # is "of the RCAs we did write, how many recorded a cause".
    _rc_total = _q(f"SELECT COUNT(*) n FROM capa c JOIN nc USING(nc_id) "
                   f"WHERE 1=1{_rca_only}{_rc_flt}", _BF_PR).iloc[0]["n"]

    rc1, rc2 = st.columns(2)

    # ---- Real Origin Area L1 (+ L2 drill-down) ----
    with rc1:
        st.markdown("**NCs by Real Origin Area (L1)**")
        df_o1 = _q(f"""SELECT c.origin_area_l1 AS area, COUNT(*) AS n
                       FROM capa c JOIN nc USING(nc_id)
                       WHERE c.origin_area_l1 IS NOT NULL{_rca_only}{_rc_flt}
                       GROUP BY area ORDER BY n DESC""", _BF_PR)
        if not df_o1.empty:
            fig = px.bar(df_o1, x="n", y="area", orientation="h",
                         color_discrete_sequence=["#1C7293"], text="n")
            fig.update_layout(height=280, margin=dict(l=0, r=0, t=MODEBAR_T, b=0), showlegend=False,
                              yaxis=dict(categoryorder="total ascending"), xaxis_title="", yaxis_title="")
            fig.update_traces(textposition="outside", hovertemplate="<b>%{y}</b><br>NCs: %{x}<extra></extra>")
            st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
            st.caption(f"Where the problem actually originated. **{int(df_o1['n'].sum())}** of "
                       f"{int(_rc_total)} NCs **with an RCA** have an Origin Area recorded. "
                       f"(For coverage against every NC, see CAPA Coverage above.)")
            st.download_button("📥 Excel", to_excel_bytes(df_o1, "Origin_Area_L1"),
                               "origin_area_l1.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="dl_o1")

            _o_pick = st.selectbox("🔍 Explode an Origin Area into L2",
                                   ["— select —"] + df_o1["area"].tolist(), key="o1_drill")
            if _o_pick != "— select —":
                df_o2 = _q(f"""SELECT COALESCE(c.origin_area_l2,'(L2 not recorded)') AS sub, COUNT(*) AS n
                               FROM capa c JOIN nc USING(nc_id)
                               WHERE c.origin_area_l1 = ?{_rca_only}{_rc_flt}
                               GROUP BY sub ORDER BY n DESC""", [_o_pick] + _BF_PR)
                if not df_o2.empty:
                    figd = px.bar(df_o2, x="n", y="sub", orientation="h",
                                  color_discrete_sequence=["#65A6C0"], text="n")
                    figd.update_layout(height=max(160, 42 * len(df_o2)), margin=dict(l=0, r=0, t=MODEBAR_T, b=0),
                                       showlegend=False, yaxis=dict(categoryorder="total ascending"),
                                       xaxis_title="", yaxis_title="")
                    figd.update_traces(textposition="outside",
                                       hovertemplate="<b>%{y}</b><br>NCs: %{x}<extra></extra>")
                    st.plotly_chart(figd, use_container_width=True, config=CHART_CONFIG, key="o2_chart")
                    st.caption(f"L2 breakdown for **{_o_pick}**.")
        else:
            st.info("No Origin Area recorded for the current selection.")

    # ---- RC Category L1 (+ L2 drill-down) ----
    with rc2:
        st.markdown("**NCs by RC Category (L1)**")
        df_r1 = _q(f"""SELECT c.rc_category_l1 AS cat, COUNT(*) AS n
                       FROM capa c JOIN nc USING(nc_id)
                       WHERE c.rc_category_l1 IS NOT NULL{_rca_only}{_rc_flt}
                       GROUP BY cat ORDER BY n DESC""", _BF_PR)
        if not df_r1.empty:
            fig = px.bar(df_r1, x="n", y="cat", orientation="h",
                         color_discrete_sequence=["#21295C"], text="n")
            fig.update_layout(height=280, margin=dict(l=0, r=0, t=MODEBAR_T, b=0), showlegend=False,
                              yaxis=dict(categoryorder="total ascending"), xaxis_title="", yaxis_title="")
            fig.update_traces(textposition="outside", hovertemplate="<b>%{y}</b><br>NCs: %{x}<extra></extra>")
            st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
            st.caption(f"What kind of cause it was. **{int(df_r1['n'].sum())}** of {int(_rc_total)} NCs "
                       f"**with an RCA** have an RC Category recorded.")
            st.download_button("📥 Excel", to_excel_bytes(df_r1, "RC_Category_L1"),
                               "rc_category_l1.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="dl_r1")

            _r_pick = st.selectbox("🔍 Explode an RC Category into L2",
                                   ["— select —"] + df_r1["cat"].tolist(), key="r1_drill")
            if _r_pick != "— select —":
                df_r2 = _q(f"""SELECT COALESCE(c.rc_category_l2,'(L2 not recorded)') AS sub, COUNT(*) AS n
                               FROM capa c JOIN nc USING(nc_id)
                               WHERE c.rc_category_l1 = ?{_rca_only}{_rc_flt}
                               GROUP BY sub ORDER BY n DESC""", [_r_pick] + _BF_PR)
                if not df_r2.empty:
                    figd = px.bar(df_r2, x="n", y="sub", orientation="h",
                                  color_discrete_sequence=["#5A63A0"], text="n")
                    figd.update_layout(height=max(160, 42 * len(df_r2)), margin=dict(l=0, r=0, t=MODEBAR_T, b=0),
                                       showlegend=False, yaxis=dict(categoryorder="total ascending"),
                                       xaxis_title="", yaxis_title="")
                    figd.update_traces(textposition="outside",
                                       hovertemplate="<b>%{y}</b><br>NCs: %{x}<extra></extra>")
                    st.plotly_chart(figd, use_container_width=True, config=CHART_CONFIG, key="r2_chart")
                    st.caption(f"L2 breakdown for **{_r_pick}**.")
        else:
            st.info("No RC Category recorded for the current selection.")

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

# ---- RCA rows that exist but never recorded a cause: the actionable cleanup list ----
rca_blank = None
if HAS_CAPA_TYPE:
    rca_blank = _q(f"""
        SELECT c.nc_id, n.project, n.owner, c.responsible, n.status
        FROM capa c JOIN nc n USING(nc_id)
        WHERE c.capa_type='RCA'
          AND (c.origin_area_l1 IS NULL OR c.rc_category_l1 IS NULL)
          {_BF_CL and (' AND ' + ' AND '.join(_BF_CL)) or ''}
        ORDER BY n.project, c.nc_id
    """, _BF_PR)
    st.caption(f"**{len(rca_blank)} NCs have an RCA row but no Origin Area and/or no RC Category.** "
               "These are the ones where the analysis was opened but the cause was never written down — "
               "the fastest fill-rate win.")
    st.dataframe(rca_blank, use_container_width=True, hide_index=True, height=240)
    st.download_button("📥 Excel", to_excel_bytes(rca_blank, "RCA_Missing_Cause"),
                       "rca_missing_cause.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="dl_rca_blank")

st.caption(f"Data source: `quality.db` · Last modified: "
           f"{datetime.fromtimestamp(Path(DB_FILE).stat().st_mtime).strftime('%d.%m.%Y %H:%M')}")


# ══════════════════════════════════════════════════════════════════════════════
# COMPLETE REPORT — one workbook with every chart's data + the filter context
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
with st.container(border=True):
    st.markdown("### 📦 Complete report")
    st.caption("One Excel workbook containing every chart's underlying data, plus a **Filters** "
               "sheet recording when the report was taken and which filters were active "
               "(From / To / Since / NC type / Status / Project / Owner / Data source). "
               "All sheets reflect the filters currently applied.")

    _report_sets = {}
    for _name, _var in [
        ("Burndown_OpenPerMonth", "actual"),
        ("Monthly_Opens_vs_Closes", "df_mo"),
        ("Opened_vs_Closed_Trend", "df_trend"),
        ("Top_Projects_WIP", "df_proj"),
        ("Detection_Area", "df_area"),
        ("Production_vs_Supplier", "df_ps"),
        ("Owner_Workload", "df_own"),
        ("NC_Total_per_Year", "df_yr"),
        ("CAPA_Coverage", "df_cov"),
        ("CAPA_Combinations", "df_combo"),
        ("Origin_Area_L1", "df_o1"),
        ("RC_Category_L1", "df_r1"),
        ("RCA_Missing_Cause", "rca_blank"),
        ("Data_Quality", "incomplete"),
    ]:
        _d = globals().get(_var)
        if _d is not None:
            _report_sets[_name] = _d

    st.download_button(
        "📥 Download complete report (.xlsx)",
        build_full_report_bytes(_report_sets, include_since=True),
        f"BRM_Quality_Report_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_full_report", type="primary")
    st.caption(f"Includes {len(_report_sets)} data sheets + Filters context sheet.")