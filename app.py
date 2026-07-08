"""
app.py
------
Streamlit dashboard for the Quality BRM.

Shows all the data that goes into slides 1-4 and gives two download buttons:
    - BRM PowerPoint (4 slides, Simone's format)
    - BRM Excel workbook (multi-sheet, all the raw + summary data)

Run:
    streamlit run app.py

Then open http://localhost:8501 in the browser.
"""

from datetime import datetime
from pathlib import Path
import sqlite3
import subprocess
import sys

import pandas as pd
import plotly.express as px
import streamlit as st

from auth import check_password
from make_brm import build_brm_pptx
from make_excel import build_brm_excel

DB_FILE = "quality.db"

st.set_page_config(
    page_title="BU Launchers - Quality BRM",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---- PASSWORD GATE ----
if not check_password():
    st.stop()


# --------- STYLES ---------
st.markdown("""
    <style>
    /* Make the main content area a bit narrower for readability */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 1400px;
    }

    /* Compact metric cards */
    [data-testid="stMetricValue"] {
        font-size: 1.8rem;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.85rem;
    }

    /* Header band */
    .main-header {
        background: linear-gradient(90deg, #1E2761 0%, #F26E21 100%);
        padding: 0.8rem 1.5rem;
        border-radius: 8px;
        color: white;
        margin-bottom: 1rem;
    }
    .main-header h1 { color: white; margin: 0; font-size: 1.5rem; }
    .main-header p { color: #e0e0e0; margin: 0; font-size: 0.9rem; }

    /* Section anchor buttons in sidebar */
    .nav-link {
        display: block;
        padding: 0.4rem 0.6rem;
        margin: 0.2rem 0;
        border-radius: 4px;
        text-decoration: none;
        color: #1E2761;
        border-left: 3px solid transparent;
    }
    .nav-link:hover {
        background: #f0f2f6;
        border-left-color: #F26E21;
    }
    </style>
""", unsafe_allow_html=True)


# --------- HEADER ---------
st.markdown("""
    <div class="main-header">
        <h1>BU Launchers Switzerland — Quality</h1>
        <p>Business Review Dashboard · Data from NC Tracker</p>
    </div>
""", unsafe_allow_html=True)


# --------- CHECK DB EXISTS ---------
if not Path(DB_FILE).exists():
    st.error(f"❌ `{DB_FILE}` not found. Run `python ingest.py` first.")
    st.stop()


def _q(sql):
    with sqlite3.connect(DB_FILE) as conn:
        return pd.read_sql(sql, conn)


# Chart config — fully static, NO zoom/pan possible
# staticPlot=True disables ALL interactions - no more zoom traps
CHART_CONFIG = {
    "displayModeBar": False,
    "staticPlot": True,          # <-- this line kills ALL zoom/pan
    "responsive": True,
}


# --------- SIDEBAR ---------
with st.sidebar:
    st.markdown("### 📍 Navigate")
    st.markdown("""
        <a href="#slide-2" class="nav-link">📊 Slide 2 — Dashboard</a>
        <a href="#slide-3" class="nav-link">📈 Slide 3 — Trends 1/2</a>
        <a href="#slide-4" class="nav-link">📉 Slide 4 — Trends 2/2</a>
        <a href="#data-quality" class="nav-link">🔍 Data Quality</a>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### ⚙️ Report settings")
    month_label = st.text_input("Reporting month", value=f"{datetime.now().strftime('%B %Y')}")

    st.markdown("---")
    st.markdown("### 📥 Downloads")

    if st.button("🎯 Build PowerPoint", use_container_width=True):
        with st.spinner("Generating..."):
            st.session_state["pptx"] = build_brm_pptx(month=month_label)
            st.session_state["pptx_name"] = f"BRM_Quality_{month_label.replace(' ', '_')}.pptx"

    if "pptx" in st.session_state:
        st.download_button(
            "⬇️ Download .pptx",
            data=st.session_state["pptx"],
            file_name=st.session_state["pptx_name"],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

    st.markdown("")

    if st.button("📊 Build Excel", use_container_width=True):
        with st.spinner("Generating..."):
            st.session_state["xlsx"] = build_brm_excel()
            st.session_state["xlsx_name"] = f"BRM_Quality_{month_label.replace(' ', '_')}.xlsx"

    if "xlsx" in st.session_state:
        st.download_button(
            "⬇️ Download .xlsx",
            data=st.session_state["xlsx"],
            file_name=st.session_state["xlsx_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.markdown("---")
    if st.button("🔄 Re-run ingest.py", use_container_width=True):
        with st.spinner("Reloading from Excel..."):
            result = subprocess.run([sys.executable, "ingest.py"], capture_output=True, text=True)
            if result.returncode == 0:
                st.success("✓ Refreshed")
            else:
                st.error("Failed")
                st.code(result.stderr)


# --------- SLIDE 2 — DASHBOARD ---------
st.markdown('<div id="slide-2"></div>', unsafe_allow_html=True)
st.subheader(f"📊 Slide 2 — Quality Dashboard · {month_label}")

kpis = _q("""
    SELECT
        SUM(CASE WHEN is_open=1 THEN 1 ELSE 0 END) AS ncs_wip,
        SUM(CASE WHEN is_open=0 THEN 1 ELSE 0 END) AS ncs_closed,
        SUM(CASE WHEN classification LIKE 'Major%' AND is_open=1 THEN 1 ELSE 0 END) AS major_open,
        SUM(CASE WHEN is_supplier_nc=1 AND is_open=1 THEN 1 ELSE 0 END) AS supplier_open,
        SUM(CASE WHEN is_supplier_nc=0 AND is_open=1 THEN 1 ELSE 0 END) AS production_open,
        SUM(CASE WHEN project IS NULL THEN 1 ELSE 0 END) AS blank_project,
        SUM(CASE WHEN detection_area IS NULL THEN 1 ELSE 0 END) AS blank_detection
    FROM nc
""").iloc[0]

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("NCs Open (WIP)", int(kpis["ncs_wip"]))
c2.metric("NCs Closed", int(kpis["ncs_closed"]))
c3.metric("Major NCs Open", int(kpis["major_open"]))
c4.metric("Production NCs", int(kpis["production_open"]))
c5.metric("Supplier NCs", int(kpis["supplier_open"]))

if kpis["blank_detection"] > 20 or kpis["blank_project"] > 10:
    st.warning(
        f"⚠️ Data quality: {int(kpis['blank_detection'])} blank Detection, "
        f"{int(kpis['blank_project'])} blank Project. Charts affected."
    )

st.markdown("")


# --------- SLIDE 3 — TRENDS 1/2 ---------
st.markdown('<div id="slide-3"></div>', unsafe_allow_html=True)
st.subheader("📈 Slide 3 — Quality Performance & Trends (1/2)")

col_a, col_b = st.columns(2)

with col_a:
    st.markdown("**Top 6 Projects — NCs WIP**")
    df = _q("""
        SELECT COALESCE(project, '(no project)') AS project, COUNT(*) AS open_ncs
        FROM nc WHERE is_open = 1
        GROUP BY project ORDER BY open_ncs DESC LIMIT 6
    """)
    fig = px.bar(df, x="open_ncs", y="project", orientation="h",
                 color_discrete_sequence=["#1E2761"], text="open_ncs")
    fig.update_layout(height=260, margin=dict(l=0, r=0, t=10, b=0), showlegend=False,
                      yaxis=dict(categoryorder="total ascending"), xaxis_title="", yaxis_title="")
    fig.update_traces(textposition="outside")
    st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)

with col_b:
    st.markdown("**NCs by Detection Area (YTD 2026)**")
    df = _q("""
        SELECT COALESCE(detection_area, 'BLANK - to clean') AS area, COUNT(*) AS n
        FROM nc WHERE substr(created_on, 1, 4) = '2026'
        GROUP BY area ORDER BY n DESC LIMIT 12
    """)
    colors = ["#E53E3E" if a.startswith("BLANK") else "#F26E21" for a in df["area"]]
    fig = px.bar(df, x="n", y="area", orientation="h",
                 color=df["area"], color_discrete_sequence=colors, text="n")
    fig.update_layout(height=260, margin=dict(l=0, r=0, t=10, b=0), showlegend=False,
                      yaxis=dict(categoryorder="total ascending"), xaxis_title="", yaxis_title="")
    fig.update_traces(textposition="outside")
    st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)

col_c, col_d = st.columns(2)

with col_c:
    st.markdown("**Top 6 Projects — Opened YTD 2026**")
    df = _q("""
        SELECT COALESCE(project, '(no project)') AS project, COUNT(*) AS n
        FROM nc WHERE substr(created_on, 1, 4) = '2026'
        GROUP BY project ORDER BY n DESC LIMIT 6
    """)
    fig = px.bar(df, x="n", y="project", orientation="h",
                 color_discrete_sequence=["#F26E21"], text="n")
    fig.update_layout(height=240, margin=dict(l=0, r=0, t=10, b=0), showlegend=False,
                      yaxis=dict(categoryorder="total ascending"), xaxis_title="", yaxis_title="")
    fig.update_traces(textposition="outside")
    st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)

with col_d:
    st.markdown("**NC Total per Year**")
    df = _q("""
        SELECT substr(created_on, 1, 4) AS yr, COUNT(*) AS n
        FROM nc WHERE created_on IS NOT NULL GROUP BY yr ORDER BY yr
    """)
    fig = px.bar(df, x="yr", y="n", color_discrete_sequence=["#1E2761"], text="n")
    fig.update_layout(height=240, margin=dict(l=0, r=0, t=10, b=0), showlegend=False,
                      xaxis_title="", yaxis_title="")
    fig.update_traces(textposition="outside")
    st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)


st.markdown("")


# --------- SLIDE 4 — TRENDS 2/2 ---------
st.markdown('<div id="slide-4"></div>', unsafe_allow_html=True)
st.subheader("📉 Slide 4 — Quality Performance & Trends (2/2)")

st.markdown("**Monthly Trend — Opens vs Closes (2026)**")
df = _q("""
    WITH opens AS (
        SELECT substr(created_on, 1, 7) AS month, COUNT(*) AS opened
        FROM nc WHERE created_on IS NOT NULL GROUP BY month
    ),
    closes AS (
        SELECT substr(closure_date, 1, 7) AS month, COUNT(*) AS closed
        FROM nc WHERE closure_date IS NOT NULL GROUP BY month
    )
    SELECT o.month, o.opened, COALESCE(c.closed, 0) AS closed
    FROM opens o LEFT JOIN closes c USING (month)
    WHERE o.month >= '2026-01' ORDER BY o.month
""")
fig = px.line(df, x="month", y=["opened", "closed"], markers=True,
              color_discrete_map={"opened": "#F26E21", "closed": "#4CAF50"})
fig.update_layout(height=260, margin=dict(l=0, r=0, t=10, b=0),
                  legend_title="", xaxis_title="", yaxis_title="NCs")
st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)

col_e, col_f = st.columns(2)

with col_e:
    st.markdown("**Production vs Supplier**")
    df = _q("""
        SELECT
            CASE WHEN is_supplier_nc = 1 THEN 'Supplier' ELSE 'Production' END AS source,
            SUM(CASE WHEN is_open = 1 THEN 1 ELSE 0 END) AS open_ncs,
            SUM(CASE WHEN is_open = 0 THEN 1 ELSE 0 END) AS closed_ncs
        FROM nc GROUP BY source
    """)
    fig = px.bar(df, x="source", y=["open_ncs", "closed_ncs"], barmode="group",
                 color_discrete_map={"open_ncs": "#F26E21", "closed_ncs": "#4CAF50"}, text_auto=True)
    fig.update_traces(textposition="outside")
    fig.update_layout(height=280, margin=dict(l=0, r=0, t=10, b=0),
                      legend_title="", xaxis_title="", yaxis_title="")
    st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)

with col_f:
    st.markdown("**Open NCs by Owner** (Top 10)")
    df = _q("""
        SELECT COALESCE(owner, '(no owner)') AS owner,
               SUM(CASE WHEN is_open = 1 THEN 1 ELSE 0 END) AS open_ncs,
               MAX(days_open) AS oldest_days
        FROM nc
        GROUP BY owner HAVING open_ncs > 0
        ORDER BY open_ncs DESC LIMIT 10
    """)
    st.dataframe(df, use_container_width=True, hide_index=True, height=280,
                 column_config={
                     "owner": "Owner",
                     "open_ncs": st.column_config.NumberColumn("Open"),
                     "oldest_days": st.column_config.NumberColumn("Oldest (days)"),
                 })


st.markdown("")


# --------- DATA QUALITY ---------
st.markdown('<div id="data-quality"></div>', unsafe_allow_html=True)
st.subheader("🔍 Data Quality — rows needing cleaning")

incomplete = _q("""
    SELECT nc_id, owner, project, detection_area, classification, status
    FROM nc
    WHERE project IS NULL OR detection_area IS NULL
       OR classification IS NULL OR owner IS NULL
    ORDER BY owner, nc_id
""")
st.caption(f"**{len(incomplete)} NCs** with missing data — send this list to owners for CW26 cleanup")
st.dataframe(incomplete, use_container_width=True, hide_index=True, height=300)

# --------- FOOTER ---------
st.caption(
    f"Data source: `quality.db` · Last modified: "
    f"{datetime.fromtimestamp(Path(DB_FILE).stat().st_mtime).strftime('%d.%m.%Y %H:%M')}"
)