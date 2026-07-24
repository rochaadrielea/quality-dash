"""
capa_view.py — the CAPA tab, rendered inside app.py via `capa_view.render()`.

A zoomable "pizza" (Plotly sunburst) that goes from the big picture down to the
detail: Launcher class -> Project -> Classification -> CAPA open/done. Click a
wedge to zoom in, click the middle to zoom out, and the report table below
always shows exactly what you're looking at. A burnout split by launcher class
(LLV / MLV / SLV) sits on top, and an RCA-by-department pizza on the side.

RULE (Adriele, current): EVERY NC owes a CAPA. So an NC with no CAPA record on
file = CAPA OPEN (outstanding), whether the NC is closed or not. An NC that has
a CAPA record (RCA / CA / PA / Ext-8D) = CAPA DONE.

  * "Done" now includes Ext-8D (external supplier 8D). It is a corrective-action
    record like the others, so an NC whose only action is an Ext-8D counts as
    covered. (Before the ingest fix, Ext-8D rows were dropped and 131 covered
    NCs wrongly showed as open.)

Launcher classes — confirmed against the CAPA tracker's own "Affected Project"
column, which literally prefixes LLV_ / MLV_ / SLV_:
    LLV = Ariane (A6) + Relativity (RS) + MHI_H3 (H3)   [+ Atlas, Vulcan share the LLV prefix]
    MLV = Vega
    SLV = Flexline + SAS
    (Vulcan is kept as its own bucket; no-project rows show as '(no project)')

This module exposes render(); app.py calls it from inside the "CAPA" tab.
Page config and the password gate live in app.py, not here.
"""
import sqlite3
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st

DB_FILE = "quality.db"
FEEDBACK_DB = "dashboard_feedback.db"   # persistent — ingest.py never touches it

LAUNCHER = {"Ariane": "LLV", "Relativity": "LLV", "MHI_H3": "LLV",
            "Vega": "MLV", "Flexline": "SLV", "SAS": "SLV", "Vulcan": "Vulcan"}
OPEN, DONE = "CAPA open", "CAPA done"
STATUS_COLORS = {OPEN: "#E53E3E", DONE: "#4CAF50"}
# fixed colours per launcher class so the burnout bars read consistently
CLASS_COLORS = {"LLV": "#1E2761", "MLV": "#F26E21", "SLV": "#5A63A0",
                "Vulcan": "#4CAF50", "(no project)": "#B0B0B0"}


# =========================================================================
# feedback button (top of every tool — Adriele's standing rule)
# =========================================================================
def feedback_widget():
    with st.expander("💡  Improve this tab — send the team your idea",
                     expanded=False):
        with st.form("capa_feedback", clear_on_submit=True):
            c1, c2 = st.columns([1, 1])
            name = c1.text_input("Your name (optional)")
            cat = c2.selectbox("Type", ["Idea / improvement", "Something's wrong",
                                        "Question", "Other"])
            msg = st.text_area("What would make this better?")
            if st.form_submit_button("Send") and msg.strip():
                try:
                    con = sqlite3.connect(FEEDBACK_DB)
                    con.execute("CREATE TABLE IF NOT EXISTS feedback("
                                "ts TEXT, page TEXT, name TEXT, category TEXT, "
                                "message TEXT)")
                    con.execute("INSERT INTO feedback VALUES (?,?,?,?,?)",
                                (datetime.now().isoformat(timespec="seconds"),
                                 "CAPA", name, cat, msg.strip()))
                    con.commit()
                    con.close()
                    st.success("Thank you! Sent to the team.")
                except Exception as e:
                    st.error(f"Could not save: {e}")


# =========================================================================
# data
# =========================================================================
@st.cache_data(ttl=300)
def load_nc():
    con = sqlite3.connect(DB_FILE)
    nc = pd.read_sql(
        "SELECT nc_id, project, classification, owner, detection_area, "
        "created_on, closure_date, is_open, status_state FROM nc", con)
    capa_ncs = set(pd.read_sql("SELECT DISTINCT nc_id FROM capa", con)["nc_id"])
    con.close()
    nc["Launcher"] = nc["project"].map(LAUNCHER).fillna("(no project)")
    nc["Project"] = nc["project"].fillna("(no project)")

    def _cls(c):
        s = str(c)
        if s.startswith("Major"):
            return "Major"
        if "Minor" in s:
            return "Minor"
        if c is None or s in ("0", "None", "(no class)", "nan"):
            return "(no class)"
        return s
    nc["Class"] = nc["classification"].apply(_cls)
    nc["CAPA"] = nc["nc_id"].apply(lambda x: DONE if x in capa_ncs else OPEN)
    return nc


@st.cache_data(ttl=300)
def load_rca_departments():
    con = sqlite3.connect(DB_FILE)
    df = pd.read_sql(
        "SELECT COALESCE(origin_area_l1,'(not recorded)') AS dept, "
        "COALESCE(rc_category_l1,'(not recorded)') AS cause "
        "FROM capa WHERE capa_type='RCA'", con)
    con.close()
    return df


# =========================================================================
# render — called by app.py from inside the CAPA tab
# =========================================================================
def render():
    feedback_widget()
    st.title("CAPA — Root Cause & Actions")
    st.caption("Rule: every NC owes a CAPA. No CAPA on record = **CAPA open** "
               "(even if the NC is closed). 'Done' counts RCA, CA, PA and "
               "Ext-8D (external supplier 8D).")

    nc = load_nc()

    # ---- burnout split by launcher class (LLV / MLV / SLV) --------------
    st.subheader("CAPA burnout by launcher class")
    st.caption("Each launcher class: how many of its NCs still owe a CAPA "
               "(red) vs how many are covered (green). LLV = Ariane + RS + H3, "
               "MLV = Vega, SLV = Flexline + SAS.")
    burn = (nc.groupby(["Launcher", "CAPA"]).size()
              .reset_index(name="n"))
    # keep a stable class order
    _order = ["LLV", "MLV", "SLV", "Vulcan", "(no project)"]
    burn["Launcher"] = pd.Categorical(burn["Launcher"], categories=_order,
                                      ordered=True)
    burn = burn.sort_values("Launcher")
    figb = px.bar(burn, x="Launcher", y="n", color="CAPA",
                  color_discrete_map=STATUS_COLORS, text="n",
                  category_orders={"Launcher": _order})
    figb.update_layout(barmode="stack", height=320,
                       margin=dict(t=10, l=0, r=0, b=0),
                       xaxis_title="", yaxis_title="NCs",
                       legend_title_text="")
    figb.update_traces(textposition="inside")
    st.plotly_chart(figb, width='stretch')

    # a small coverage table under the burnout
    piv = (nc.pivot_table(index="Launcher", columns="CAPA", values="nc_id",
                          aggfunc="count", fill_value=0)
             .reindex(_order).dropna(how="all"))
    for _col in (OPEN, DONE):
        if _col not in piv.columns:
            piv[_col] = 0
    piv["Total"] = piv[OPEN] + piv[DONE]
    piv["Coverage"] = (100 * piv[DONE] / piv["Total"].replace(0, 1)).round(0)
    piv = piv.rename(columns={OPEN: "Open", DONE: "Done"})
    st.dataframe(piv[["Open", "Done", "Total", "Coverage"]]
                 .style.format({"Coverage": "{:.0f}%"}),
                 width='stretch')

    st.divider()

    # ---- filters (also driven by the sunburst clicks) ------------------
    if "capa_path" not in st.session_state:
        st.session_state.capa_path = []          # e.g. ['LLV','Ariane']

    fc = st.columns([1, 1, 1, 2])
    major_only = fc[0].checkbox("Major only", value=False)
    show_by = fc[1].radio("Colour by", ["CAPA status", "Classification"],
                          horizontal=True)
    if fc[2].button("⬆ Zoom out", disabled=not st.session_state.capa_path):
        st.session_state.capa_path.pop()
    if fc[3].button("Reset to top"):
        st.session_state.capa_path = []

    df = nc.copy()
    if major_only:
        df = df[df["Class"] == "Major"]

    # apply the current drill path (Launcher -> Project -> Class)
    path = st.session_state.capa_path
    levels = ["Launcher", "Project", "Class"]
    for i, val in enumerate(path):
        if i < len(levels):
            df = df[df[levels[i]] == val]

    crumbs = " › ".join(["All"] + path)
    st.markdown(f"**Viewing:** {crumbs}")

    # ---- KPI row -------------------------------------------------------
    tot = len(df)
    opn = int((df["CAPA"] == OPEN).sum())
    done = int((df["CAPA"] == DONE).sum())
    k = st.columns(4)
    k[0].metric("NCs in view", f"{tot:,}")
    k[1].metric("CAPA open", f"{opn:,}", help="NCs with no CAPA on record")
    k[2].metric("CAPA done", f"{done:,}")
    k[3].metric("Coverage", f"{(100*done/tot):.0f}%" if tot else "—")

    # ---- the zoomable pizza (sunburst) + RCA-by-department -------------
    left, right = st.columns([3, 2])
    with left:
        st.subheader("Click a slice to zoom in · centre to zoom out")
        remaining = levels[len(path):] + ["CAPA"]
        color = "CAPA" if show_by == "CAPA status" else "Class"
        cmap = STATUS_COLORS if show_by == "CAPA status" else None
        if df.empty:
            st.info("No NCs in this view.")
        else:
            fig = px.sunburst(
                df, path=remaining, color=color,
                color_discrete_map=cmap or {},
                maxdepth=3)
            fig.update_traces(
                insidetextorientation="radial",
                hovertemplate="<b>%{label}</b><br>%{value} NCs<extra></extra>")
            fig.update_layout(margin=dict(t=10, l=0, r=0, b=0), height=460)
            sel = st.plotly_chart(fig, width='stretch',
                                  on_select="rerun", key="sunburst")
            # a click on the first ring drills one level down
            try:
                pts = sel["selection"]["points"] if sel else []
                if pts:
                    label = pts[0].get("label")
                    nxt = levels[len(path)] if len(path) < len(levels) else None
                    if label and nxt and label in set(df[nxt]):
                        st.session_state.capa_path.append(label)
                        st.rerun()
            except Exception:
                pass

    with right:
        st.subheader("RCA by department")
        st.caption("Where the root cause was located, from the RCA rows in the "
                   "CAPA tracker (Origin Area).")
        rca = load_rca_departments()
        dep = rca["dept"].value_counts().reset_index()
        dep.columns = ["dept", "n"]
        figd = px.pie(dep, names="dept", values="n", hole=0.45)
        figd.update_traces(textposition="inside", textinfo="percent+label")
        figd.update_layout(margin=dict(t=10, l=0, r=0, b=0), height=250,
                           showlegend=False)
        st.plotly_chart(figd, width='stretch')

        st.subheader("Root-cause category")
        cau = rca["cause"].value_counts().reset_index()
        cau.columns = ["cause", "n"]
        figc = px.pie(cau, names="cause", values="n", hole=0.45)
        figc.update_traces(textposition="inside", textinfo="percent+label")
        figc.update_layout(margin=dict(t=10, l=0, r=0, b=0), height=250,
                           showlegend=False)
        st.plotly_chart(figc, width='stretch')

    # ---- the live report of what you're viewing ------------------------
    st.subheader(f"Report — {crumbs}  ({tot:,} NCs)")
    report = df[["nc_id", "Launcher", "Project", "Class", "CAPA",
                 "status_state", "owner", "detection_area", "created_on"]]\
        .sort_values(["CAPA", "Project", "nc_id"])
    st.dataframe(report, width='stretch', height=340)

    buf = report.to_csv(index=False).encode("utf-8")
    st.download_button("⬇ Download this view (CSV)", buf,
                       file_name=f"CAPA_{'_'.join(['ALL']+path)}.csv",
                       mime="text/csv")
