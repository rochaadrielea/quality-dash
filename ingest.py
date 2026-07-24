"""
ingest.py
---------
Reads NCR_Cutover_Tracker.xlsx AND NC_s_Overview (SAP export),
cleans, merges (deduplicates overlapping NCs), loads to SQLite.

Two files in data/:
  - NCR_Cutover_Tracker*.xlsx  → active tracker (owner/disposition detail)
  - NC_s_Overview*.xlsx        → SAP full history (CoPQ/leadtime/defect)
  - *CAPA*.xls[xm]             → CAPA/RCA tracker (optional)

Merge logic:
  - NCs in both files    → enriched: tracker fields + SAP fields, source='both'
  - Tracker-only rows    → source='tracker'
  - SAP-only rows        → source='sap'

CAPA table:
  - Keyed on (nc_id, capa_type). NC/SCAR Number is the foreign key to nc.nc_id.
  - Keeps ALL three types (RCA / CA / PA) — one row per NC per type.
  - Origin Area / RC Category L1+L2 are only populated on RCA rows by design;
    CA and PA rows exist to record that the action exists, not a root cause.
  - Where the same (nc_id, capa_type) appears more than once, the row carrying
    the most content wins (an all-N/A duplicate never beats a filled row).

Run daily/weekly to refresh. Idempotent (safe to re-run).

Usage:
    python ingest.py
"""

import re
import sqlite3
from pathlib import Path
from datetime import datetime

import pandas as pd

# ---- CONFIG ----
DATA_DIR = Path("data")
DB_FILE = "quality.db"
TRACKER_SHEET = "NC_Tracker_Black_Out"
CAPA_GLOB = "*CAPA*.xls*"      # matches .xlsx AND .xlsm (the real file is macro-enabled)
CAPA_SHEET = "Requestor"
# RCA / CA / PA are the internal actions. EXT-8D is the external supplier 8D —
# it IS a corrective-action record, so an NC that only carries an Ext-8D still
# "has a CAPA". Dropping it (as this list used to) made 131 covered NCs read as
# 'CAPA open' on the burnout. It is kept as its own type so the RCA-only views
# (pizza, root-cause) are unaffected — they filter capa_type='RCA' explicitly.
CAPA_TYPES = ["RCA", "CA", "PA", "EXT-8D"]

# Column maps live at module level so validate.py can check the real mapping
# rather than a hand-kept copy of it. Two lists that can disagree WILL disagree:
# an earlier version of validate.py had its own list, and warned that 'NC Type',
# 'Failure' and 'Material' were unrecognised columns while load_tracker was
# reading all three quite happily.
TRACKER_COL_MAP = {
    "System": "system",
    "ID-Blackout": "nc_id",            # was "Title"
    "Title": "nc_id",                  # fallback for old snapshots
    "NC Type": "nc_type",
    "TC ID": "tc_id",
    "Migrated to EZ1": "migrated_to_ez1",
    "Project": "project",              # was "Project + Flight Unit"
    "Project + Flight Unit": "project",# fallback
    "Flight Unit": "flight_unit",      # new dedicated column
    "Detection": "detection_area",
    "Title and Problem Description": "description",  # was "Description"
    "Description": "description",      # fallback
    "Failure": "failure",
    "Material": "material",
    "Batch": "batch",
    "Issue Owner (QA/PA)": "owner",
    "Issue Owner": "owner",            # fallback
    "Created On": "created_on",
    "NRB disposition": "nrb_disposition",
    "Disposition Implemented Date": "disposition_date",
    "Classification": "classification",
    "PSP ref.": "psp_ref",
    "NC WBS (EzyOne)": "nc_wbs",
    "Status": "status",                # was "Notific. Status"
    "Notific. Status": "status",       # fallback
    "Closure date": "closure_date",
    "Supplier name": "supplier_name",
}

# Columns whose absence actually breaks something downstream. The rest of
# TRACKER_COL_MAP is either optional or has a fallback spelling.
TRACKER_REQUIRED_COLS = [
    "System", "ID-Blackout", "Issue Owner (QA/PA)", "Created On", "Status",
]


# ---- FILE FINDERS ----
def find_file(pattern):
    """Find the newest file matching pattern in data/.
    Skips Excel lock/temp files (~$...) which appear while a workbook is open."""
    DATA_DIR.mkdir(exist_ok=True)
    candidates = [p for p in DATA_DIR.glob(pattern) if not p.name.startswith("~$")]
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


# ---- CLEANERS ----
def clean_status(val):
    if pd.isna(val) or not str(val).strip() or str(val).strip() == '\xa0':
        return None
    return str(val).strip().upper()


def clean_classification(val):
    if pd.isna(val) or not str(val).strip() or str(val).strip() == '\xa0':
        return None
    v = str(val).strip()
    v = re.sub(r"\s*-\s*", " ", v)
    v = re.sub(r"\s+", " ", v)
    v = v.title()
    return v


def clean_owner(val):
    """Normalise 'Issue Owner (QA/PA)' to one canonical spelling per person.

    The tracker has no naming standard: the same person is entered as a bare
    first name on some rows and a full name on others, which splits their NC
    count across two owner values and undercounts everyone on the dashboard
    (verified against NCR_Cutover_Tracker_14_07_2026: Vitaly 34 + Vitaly Meshin
    11, Rikard 26 + Rikard Bjon 9, Tiziano 33 + Tiziano Casarico 3, and so on).

    Only unambiguous cases are mapped — a bare first name that matches exactly
    one full name elsewhere in the file. Deliberately NOT mapped, by Adriele's
    decision, and left exactly as typed:
      - 'Güven', 'Nikolaos'  — no full name appears anywhere in the tracker
      - 'E.Martins/ R. Bjon', 'Vitaly /Salvatore', 'Vitaly/ Domingos'
                             — two people on one NC; the data model has one
                               owner per NC, so these stay as their own value
                               rather than silently crediting one person
    """
    if pd.isna(val) or not str(val).strip() or str(val).strip() == '\xa0':
        return None
    v = str(val).strip()
    aliases = {
        # bare first name -> the single full name it maps to
        "Vitaly": "Vitaly Meshin",
        "Rikard": "Rikard Bjon",
        "Tiziano": "Tiziano Casarico",
        "Noel": "Noel Orwa",
        "Joaquin": "Joaquin Vera Rubio",
        "Domingos": "Domingos Moreira",
        # long form -> short canonical form
        "Domingos Manuel Ferreira Moreira": "Domingos Moreira",
        "Noel Kenoreny Orwa": "Noel Orwa",
        "S. Scampini": "Simone Scampini",
    }
    return aliases.get(v, v)


def clean_project(val):
    """Normalize raw project text to the canonical validation set.

    Canonical: Ariane, Vega, MHI_H3, Relativity, SAS, Vulcan, Flexline.
    Anything that doesn't map cleanly -> None (shown as '(no project)').
    """
    if pd.isna(val):
        return None
    v = str(val).strip()
    if not v or v == '\xa0':
        return None

    up = v.upper()

    # Explicit non-projects / placeholders -> no project
    _blanks = ("BLANK", "(BLANK)", "NOT ASSIGNED", "TBD", "??", "?", "ALL",
               "RETURN TO FLY", "N/A", "-")
    if up in _blanks:
        return None
    if up.startswith("ADMIN") or up.startswith("PU-L") or up.startswith("TESTING") \
       or up.startswith("SPACE ONE"):
        return None

    # Canonical keyword mapping (order matters: check specific before generic)
    if "FLEXLINE" in up:
        return "Flexline"
    if up.startswith("MHI") or up == "H3" or "H3" in up.split():
        return "MHI_H3"
    if "RELATIVITY" in up or up.startswith("RS "):
        return "Relativity"
    if "VULCAN" in up or up.startswith("VCN"):
        return "Vulcan"
    if "SAS" in up or "KDS" in up:
        return "SAS"
    if "VEGA" in up:
        return "Vega"
    if "ARIANE" in up or up.startswith("A6"):
        return "Ariane"

    # Unrecognized -> no project (kept out of the clean charts)
    return None


def clean_text(val):
    if pd.isna(val):
        return None
    v = str(val).strip()
    if v in ('', '\xa0', '-'):
        return None
    return v


def parse_date(val):
    if pd.isna(val) or not str(val).strip() or str(val).strip() == '\xa0':
        return None
    try:
        return pd.to_datetime(val, dayfirst=True).date().isoformat()
    except Exception:
        return None


def clean_level(val):
    """Clean an Origin/RC level value. '0', 'N/A', blank -> None (not recorded)."""
    if pd.isna(val):
        return None
    v = str(val).strip()
    if not v or v == '\xa0':
        return None
    if v.upper() in ("N/A", "NA", "N.A.", "-", "#N/A", "NONE"):
        return None
    if v == "0":          # '0' is the placeholder people type for "nothing"
        return None
    return v


def clean_capa_type(val):
    """Normalise the CAPA Type column to RCA / CA / PA. Anything else -> None."""
    if pd.isna(val):
        return None
    v = str(val).strip().upper()
    if v in ("", "NAN", "N/A", "NA", "NONE", "-"):
        return None
    return v if v in CAPA_TYPES else None


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1: Load the CAPA / RCA tracker
# ══════════════════════════════════════════════════════════════════════════════
def load_capa():
    """Load the CAPA / RCA tracker. OPTIONAL — returns None if the file isn't in data/.

    Keyed on (nc_id, capa_type): every NC can carry an RCA row, a CA row and a
    PA row. All three are kept — the presence of a CA/PA row is itself the fact
    being reported, even though the root-cause columns on those rows are N/A by
    design.

    Header row is auto-detected so a banner row above it won't break the load.
    """
    path = find_file(CAPA_GLOB)
    if path is None:
        print(f"  No {CAPA_GLOB} in {DATA_DIR}/ — skipping CAPA (charts will hide).")
        return None
    print(f"  Reading {path.name}...")

    try:
        xl = pd.ExcelFile(path)
        sheet = CAPA_SHEET if CAPA_SHEET in xl.sheet_names else xl.sheet_names[0]
        if sheet != CAPA_SHEET:
            print(f"  ! Sheet '{CAPA_SHEET}' not found, using '{sheet}'. Tabs: {xl.sheet_names}")

        # Auto-detect the header row: the row containing 'NC/SCAR'
        probe = pd.read_excel(path, sheet_name=sheet, header=None, nrows=12, dtype=str)
        header_row = 0
        for i in range(len(probe)):
            row_txt = " ".join(str(x) for x in probe.iloc[i].tolist())
            if "NC/SCAR" in row_txt:
                header_row = i
                break
        df = pd.read_excel(path, sheet_name=sheet, header=header_row, dtype=str)
    except Exception as e:
        print(f"  ! Could not read CAPA file: {e} — skipping.")
        return None

    df.columns = [str(c).strip() for c in df.columns]

    def pick(*names):
        for n in names:
            if n in df.columns:
                return n
        return None

    c_nc = pick("NC/SCAR Number", "NC/SCAR", "NC Number")
    c_type = pick("CAPA Type")
    if c_nc is None:
        print(f"  ! No 'NC/SCAR Number' column found. Columns: {list(df.columns)[:12]} — skipping.")
        return None
    if c_type is None:
        print(f"  ! No 'CAPA Type' column found. Columns: {list(df.columns)[:12]} — skipping.")
        return None

    def _clean_id(x):
        # Match the SAP loader's nc_id format exactly, or the join silently fails
        if pd.isna(x):
            return None
        if isinstance(x, (int, float)):
            try:
                return str(int(x))
            except Exception:
                return clean_text(x)
        s = clean_text(x)
        if s is None:
            return None
        try:                      # '213952.0' -> '213952'
            return str(int(float(s)))
        except Exception:
            return s

    out = pd.DataFrame()
    out["nc_id"] = df[c_nc].apply(_clean_id)
    out["capa_type"] = df[c_type].apply(clean_capa_type)

    # Root-cause levels — populated on RCA rows only, by design.
    _level_cols = []
    for tgt, src in [
        ("origin_area_l1", pick("(Real) Origin Area L1", "Origin Area L1")),
        ("origin_area_l2", pick("(Real) Origin Area L2", "Origin Area L2")),
        ("rc_category_l1", pick("RC Category L1")),
        ("rc_category_l2", pick("RC Category L2")),
    ]:
        out[tgt] = df[src].apply(clean_level) if src else None
        _level_cols.append(tgt)

    # Context columns. 'responsible' differs per capa_type on the same NC —
    # the person who owns the analysis is not always the one who owns the action.
    _ctx_cols = []
    for tgt, src in [
        ("capa_project", pick("Affected Project")),
        ("requestor", pick("Requestor")),
        ("responsible", pick("Responsible")),
        ("capa_created_on", pick("Creation date requestor")),
        ("problem_description", pick("Problem Description")),
    ]:
        out[tgt] = df[src].apply(clean_text) if src else None
        _ctx_cols.append(tgt)

    n_raw = len(out)
    out = out[out["nc_id"].notna()].copy()
    n_no_id = n_raw - len(out)

    n_before_type = len(out)
    out = out[out["capa_type"].notna()].copy()
    n_bad_type = n_before_type - len(out)

    # ---- Dedup on (nc_id, capa_type), keeping the row with the most content ----
    # The same NC+type can appear twice: once filled, once all-N/A. Sorting by
    # content score and keeping the last means a filled row always wins over a
    # blank one, regardless of the order they sit in the sheet.
    _score_cols = _level_cols + _ctx_cols
    out["_score"] = out[_score_cols].notna().sum(axis=1)
    out = (out.sort_values(["nc_id", "capa_type", "_score"], kind="stable")
              .drop_duplicates(subset=["nc_id", "capa_type"], keep="last")
              .drop(columns=["_score"])
              .reset_index(drop=True))
    n_dupes = (n_before_type - n_bad_type) - len(out)

    # ---- Reporting ----
    _by_type = out["capa_type"].value_counts().to_dict()
    n_ncs = out["nc_id"].nunique()
    n_o1 = out["origin_area_l1"].notna().sum()
    n_r1 = out["rc_category_l1"].notna().sum()
    n_rca = _by_type.get("RCA", 0)

    print(f"  {len(out)} CAPA rows across {n_ncs} NCs "
          f"(RCA: {_by_type.get('RCA', 0)} | CA: {_by_type.get('CA', 0)} | PA: {_by_type.get('PA', 0)})")
    if n_no_id or n_bad_type or n_dupes:
        print(f"    dropped: {n_no_id} no NC number | {n_bad_type} unrecognised CAPA type "
              f"| {n_dupes} duplicate (nc_id, capa_type)")
    print(f"    Origin L1 filled: {n_o1}/{n_rca} RCA rows | "
          f"RC Category L1 filled: {n_r1}/{n_rca} RCA rows")
    return out


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2: Load the NCR Cutover Tracker
# ══════════════════════════════════════════════════════════════════════════════
def load_tracker():
    """Load and clean the NCR Cutover Tracker."""
    f = find_file("NCR_Cutover_Tracker*.xlsx")
    if not f:
        print("  ⚠ No NCR_Cutover_Tracker*.xlsx found in data/ — skipping tracker.")
        return pd.DataFrame()

    print(f"  Reading {f.name} sheet '{TRACKER_SHEET}'...")
    df = pd.read_excel(f, sheet_name=TRACKER_SHEET)

    # Drop blank/example rows
    df = df[df["System"].notna()].copy()
    df = df[df["System"].astype(str).str.strip() != '\xa0'].copy()
    # NC_EXAMPLE marker may live in the description column of the new schema
    _title_col = "Title and Problem Description" if "Title and Problem Description" in df.columns else (
        "Title" if "Title" in df.columns else None)
    if _title_col:
        df = df[df[_title_col].astype(str) != "NC_EXAMPLE"].copy()

    # Column mapping — see TRACKER_COL_MAP at module level. Kept there so
    # validate.py checks the same map the ingest actually uses.
    col_map = TRACKER_COL_MAP
    df = df.rename(columns=col_map)
    # Rename can create duplicate target names (old+new schema both mapped) —
    # keep the first non-empty occurrence of each.
    df = df.loc[:, ~df.columns.duplicated()].copy()
    # Keep only mapped columns (ignore extras), unique + order-preserving
    seen = set()
    keep = [c for c in col_map.values() if c in df.columns and not (c in seen or seen.add(c))]
    df = df[keep].copy()

    # Clean
    df["status"] = df["status"].apply(clean_status)
    df["classification"] = df["classification"].apply(clean_classification)
    df["owner"] = df["owner"].apply(clean_owner)
    df["project"] = df["project"].apply(clean_project)
    df["created_on"] = df["created_on"].apply(parse_date)
    df["closure_date"] = df["closure_date"].apply(parse_date)
    if "disposition_date" in df.columns:
        df["disposition_date"] = df["disposition_date"].apply(parse_date)

    for col in df.columns:
        if col not in ("created_on", "closure_date", "disposition_date"):
            df[col] = df[col].apply(clean_text)

    # Derived
    # Three states, not two. A blank Status is NOT 'closed' — treating it as
    # closed (which `is_open = status == 'OPEN'` silently did) removed those NCs
    # from every open-NC view, and is how owners' counts came out short. There
    # are 19 such rows in the 14/07 tracker across 8 owners. They are now
    # is_open = NULL: neither open nor closed, and visible as '(no status)'.
    def _state(s):
        # clean_status returns None, but pandas stores that as NaN (a float),
        # so test with isna() rather than `is None`.
        if pd.isna(s):
            return "(no status)"
        s = str(s)
        return "Open" if s == "OPEN" else ("Closed" if s == "CLOSED" else s.title())

    df["is_open"] = df["status"].map({"OPEN": 1, "CLOSED": 0})
    df["status_state"] = df["status"].apply(_state)
    df["is_supplier_nc"] = df["supplier_name"].notna().astype(int)
    df["days_open"] = (
        pd.to_datetime("today").normalize() - pd.to_datetime(df["created_on"], errors="coerce")
    ).dt.days

    # For SAP-system rows, nc_id is the SAP notification number (for matching)
    df["sap_notif"] = None
    sap_mask = df["system"] == "SAP"
    df.loc[sap_mask, "sap_notif"] = df.loc[sap_mask, "nc_id"].apply(
        lambda x: str(int(float(x))) if x and str(x).replace('.', '').isdigit() else None
    )

    df["source"] = "tracker"
    print(f"  {len(df)} tracker rows loaded.")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# STEP 3: Load the SAP NC Overview
# ══════════════════════════════════════════════════════════════════════════════
def load_sap_overview():
    """Load and clean the SAP NC Overview export."""
    f = find_file("NC_s_Overview*.xlsx")
    if not f:
        print("  ⚠ No NC_s_Overview*.xlsx found in data/ — skipping SAP overview.")
        return pd.DataFrame()

    print(f"  Reading {f.name}...")
    df = pd.read_excel(f)

    # Skip the totals row (row index 0 after header, where Plant='Totals')
    df = df[df.iloc[:, 0] != 'Totals'].copy()

    col_map = {
        "Plant": "plant_id",
        "Profit Center (WBS Element)": "profit_center",
        "Business Unit": "business_unit",
        "Project Text  (Notification)": "project",
        "Notification Type": "notification_type",
        "Notif. Year": "notif_year",
        "Notification": "nc_id",
        "Notification Date": "created_on",
        "Closing date": "closure_date",
        "Leadtime": "leadtime",
        "CoPQ (NC)": "copq",
        "Defect Class": "defect_class",
        "Defect Code TEXT": "defect_code_text",
        "Disposition Action TEXT": "nrb_disposition",
        "Notification Cause C TEXT": "detection_area",
        "Notification TEXT": "description",
        "Material Key": "material_key",
        "Material": "material",
        "Batch": "batch",
        "[Vendor NC]": "supplier_name",
        "System Status": "system_status",
    }

    # Only map columns that exist
    actual_map = {k: v for k, v in col_map.items() if k in df.columns}
    df = df.rename(columns=actual_map)
    keep = [c for c in actual_map.values() if c in df.columns]
    df = df[keep].copy()

    # Also check the second "Notif. Status" column (index 31 in original)
    # The first one is at index 8, the second at index 31
    # We need the one that says Open/Closed
    if "Notif. Status" in df.columns:
        df = df.rename(columns={"Notif. Status": "status"})
    else:
        # Try to get status from original dataframe
        orig = pd.read_excel(f)
        orig = orig[orig.iloc[:, 0] != 'Totals'].copy()
        if orig.shape[1] > 31:
            df["status"] = orig.iloc[:, 31].apply(clean_text)

    # Clean
    df["nc_id"] = df["nc_id"].apply(lambda x: str(int(x)) if isinstance(x, (int, float)) else clean_text(x))
    df["status"] = df["status"].apply(clean_status) if "status" in df.columns else None
    df["project"] = df["project"].apply(clean_project)
    df["detection_area"] = df["detection_area"].apply(clean_text)
    df["description"] = df["description"].apply(clean_text)
    df["nrb_disposition"] = df["nrb_disposition"].apply(clean_text)
    df["supplier_name"] = df["supplier_name"].apply(clean_text)
    df["material"] = df["material"].apply(clean_text)

    # Parse dates — SAP dates might be datetime objects or day-counts
    def parse_sap_date(val):
        if pd.isna(val) or val is None:
            return None
        if isinstance(val, datetime):
            return val.date().isoformat()
        if isinstance(val, (int, float)) and val > 30000:
            # Excel serial date
            try:
                return (datetime(1899, 12, 30) + pd.Timedelta(days=int(val))).date().isoformat()
            except Exception:
                return None
        return parse_date(val)

    df["created_on"] = df["created_on"].apply(parse_sap_date)
    df["closure_date"] = df["closure_date"].apply(parse_sap_date)

    # Derived — three states, same as the tracker (see load_tracker).
    if "status" in df.columns:
        def _state(s):
            if pd.isna(s):
                return "(no status)"
            s = str(s)
            return "Open" if s == "OPEN" else ("Closed" if s == "CLOSED" else s.title())
        df["is_open"] = df["status"].map({"OPEN": 1, "CLOSED": 0})
        df["status_state"] = df["status"].apply(_state)
    else:
        df["is_open"] = None
        df["status_state"] = "(no status)"
    df["is_supplier_nc"] = df.get("notification_type", pd.Series()).astype(str).str.contains("Z2", na=False).astype(int)

    # Classification from defect class
    def classify_defect(val):
        if pd.isna(val) or str(val).strip() in ('', '-', '\xa0'):
            return None
        v = str(val).strip().upper()
        if v in ('MA',):
            return 'Major'
        if v in ('MI',):
            return 'Minor'
        if v in ('1',):
            return 'Minor Level 1'
        if v in ('2',):
            return 'Minor Level 2'
        if v in ('3', '4', '5'):
            return f'Class {v}'
        return v
    df["classification"] = df.get("defect_class", pd.Series()).apply(classify_defect)

    df["days_open"] = (
        pd.to_datetime("today").normalize() - pd.to_datetime(df["created_on"], errors="coerce")
    ).dt.days

    # Leadtime and CoPQ as numeric
    if "leadtime" in df.columns:
        df["leadtime"] = pd.to_numeric(df["leadtime"], errors="coerce")
    if "copq" in df.columns:
        df["copq"] = pd.to_numeric(df["copq"], errors="coerce")

    df["sap_notif"] = df["nc_id"]
    df["system"] = "SAP"
    df["source"] = "sap"

    # Columns the tracker has but SAP doesn't
    for col in ["tc_id", "migrated_to_ez1", "failure", "psp_ref", "nc_wbs",
                "purchasing_doc", "po_item", "vendor_code", "disposition_date", "owner"]:
        if col not in df.columns:
            df[col] = None

    print(f"  {len(df)} SAP rows loaded.")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# STEP 4: Merge
# ══════════════════════════════════════════════════════════════════════════════
def merge_data(tracker_df, sap_df):
    """Merge tracker + SAP, deduplicating on SAP notification number."""
    if tracker_df.empty and sap_df.empty:
        raise RuntimeError("No data loaded from either source.")
    if tracker_df.empty:
        return sap_df
    if sap_df.empty:
        return tracker_df

    # Find overlapping notification numbers
    tracker_sap_notifs = set(tracker_df["sap_notif"].dropna())
    sap_notifs = set(sap_df["sap_notif"].dropna())
    overlap = tracker_sap_notifs & sap_notifs

    print(f"  Overlap: {len(overlap)} NCs in both files")
    print(f"  Tracker-only: {len(tracker_df) - len(tracker_df[tracker_df['sap_notif'].isin(overlap)])} rows")
    print(f"  SAP-only: {len(sap_df) - len(sap_df[sap_df['sap_notif'].isin(overlap)])} rows")

    # For overlapping NCs: start with tracker row, enrich with SAP columns
    tracker_overlap = tracker_df[tracker_df["sap_notif"].isin(overlap)].copy()
    sap_overlap = sap_df[sap_df["sap_notif"].isin(overlap)].set_index("sap_notif")

    # Enrich tracker rows with SAP-only fields
    sap_only_cols = ["notification_type", "defect_class", "copq", "leadtime",
                     "notif_year", "profit_center", "business_unit",
                     "material_key", "system_status"]
    for col in sap_only_cols:
        if col in sap_overlap.columns:
            tracker_overlap[col] = tracker_overlap["sap_notif"].map(
                sap_overlap[col].to_dict()
            )

    # Fill missing tracker fields from SAP where tracker is blank
    fill_cols = ["detection_area", "project", "classification", "closure_date"]
    for col in fill_cols:
        if col in sap_overlap.columns:
            mask = tracker_overlap[col].isna()
            tracker_overlap.loc[mask, col] = tracker_overlap.loc[mask, "sap_notif"].map(
                sap_overlap[col].to_dict()
            )

    tracker_overlap["source"] = "both"

    # Tracker-only rows (EZ1, Blackout, TC, or SAP NCs not in the overview)
    tracker_only = tracker_df[~tracker_df["sap_notif"].isin(overlap) | tracker_df["sap_notif"].isna()].copy()

    # SAP-only rows (not in tracker)
    sap_only = sap_df[~sap_df["sap_notif"].isin(overlap)].copy()

    # Combine all three
    merged = pd.concat([tracker_overlap, tracker_only, sap_only], ignore_index=True)

    # Ensure consistent columns
    for col in ["notification_type", "defect_class", "copq", "leadtime",
                "notif_year", "profit_center", "business_unit",
                "material_key", "system_status"]:
        if col not in merged.columns:
            merged[col] = None

    # Recalculate is_open for all rows (in case SAP data filled closure_date).
    # Rows whose status is still blank keep is_open = NULL — they are neither
    # open nor closed, and must not be silently bucketed into either.
    merged.loc[merged["status"] == "CLOSED", "is_open"] = 0
    merged.loc[merged["status"] == "OPEN", "is_open"] = 1
    if "status_state" not in merged.columns:
        merged["status_state"] = merged["status"].apply(
            lambda s: "(no status)" if pd.isna(s) or not str(s).strip()
            else ("Open" if s == "OPEN" else ("Closed" if s == "CLOSED" else str(s).title())))
    else:
        merged["status_state"] = merged["status_state"].fillna("(no status)")

    print(f"  Merged total: {len(merged)} rows")
    return merged


# ══════════════════════════════════════════════════════════════════════════════
# STEP 5: Write to SQLite
# ══════════════════════════════════════════════════════════════════════════════
def write_db(df, capa=None):
    """Write merged DataFrame to SQLite. Optionally write the CAPA table too."""
    conn = sqlite3.connect(DB_FILE)
    df.to_sql("nc", conn, if_exists="replace", index=False)

    cur = conn.cursor()
    # CAPA table is optional — only created when the file is present
    cur.execute("DROP TABLE IF EXISTS capa")
    if capa is not None and not capa.empty:
        capa.to_sql("capa", conn, if_exists="replace", index=False)
        # (nc_id, capa_type) is the natural key — index it as such
        cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_capa_key ON capa(nc_id, capa_type)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_capa_nc ON capa(nc_id)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_capa_type ON capa(capa_type)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_capa_o1 ON capa(origin_area_l1)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_capa_rc1 ON capa(rc_category_l1)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_nc_project ON nc(project)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_nc_status ON nc(status)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_nc_owner ON nc(owner)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_nc_detection ON nc(detection_area)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_nc_source ON nc(source)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_nc_created ON nc(created_on)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_nc_sap_notif ON nc(sap_notif)")
    conn.commit()
    return conn


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
def copy_warnings_to_db(conn, warn_db="warnings.db"):
    """Mirror the warning log into quality.db as a read-only copy.

    warnings.db stays the master — quality.db is dropped and rebuilt on every
    ingest, so the log cannot live here or it would be wiped each run. This
    copy exists purely so the dashboard can query warnings alongside NCs
    without opening a second database.
    """
    if not Path(warn_db).exists():
        return 0
    try:
        with sqlite3.connect(warn_db) as wc:
            df = pd.read_sql("SELECT * FROM warnings", wc)
    except Exception as e:
        print(f"  ! Could not read {warn_db}: {e}")
        return 0
    if df.empty:
        return 0
    df.to_sql("warnings", conn, if_exists="replace", index=False)
    cur = conn.cursor()
    cur.execute("CREATE INDEX IF NOT EXISTS idx_warn_check ON warnings(check_name)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_warn_subject ON warnings(subject)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_warn_sev ON warnings(severity)")
    conn.commit()
    return len(df)


def main():
    print("=" * 60)
    print("Quality DB Ingest — Merge Tracker + SAP Overview")
    print("=" * 60)

    print("\n[1/6] Validating source files...")
    try:
        from validate import run_checks
        _t = find_file("NCR_Cutover_Tracker*.xlsx")
        _s = find_file("NC_s_Overview*.xlsx")
        if _t:
            run_checks(_t, _s, clean_owner, TRACKER_SHEET)
        else:
            print("  No tracker found — skipping validation.")
    except ImportError:
        print("  validate.py not found — skipping checks.")
    except Exception as e:
        # Validation must never stop the ingest.
        print(f"  ! Validation error (ingest continues): {e}")

    print("\n[2/6] Loading NCR Cutover Tracker...")
    tracker = load_tracker()

    print("\n[3/6] Loading SAP NC Overview...")
    sap = load_sap_overview()

    print("\n[4/6] Merging (deduplicating overlaps)...")
    merged = merge_data(tracker, sap)

    print("\n[5/6] Loading CAPA / RCA tracker (optional)...")
    capa = load_capa()

    print("\n[6/6] Writing to SQLite...")
    conn = write_db(merged, capa)

    # Mirror the warning log into quality.db so the dashboard can read it.
    # Runs after write_db because that call replaces the whole file.
    # The CAPA orphan check has to run here too: it compares the capa table to
    # the nc table, so it needs the database that was just written — it cannot
    # be answered from the source files alone.
    try:
        from validate import _connect, check_capa_orphans
        _wc = _connect()
        _no = check_capa_orphans(_wc, DB_FILE)
        _wc.commit(); _wc.close()
        if _no:
            print(f"  ! {_no} CAPA NC number(s) match no NC — logged, left unmatched")
    except Exception as e:
        print(f"  ! CAPA orphan check skipped: {e}")

    _nw = copy_warnings_to_db(conn)
    if _nw:
        print(f"  {_nw} warning(s) copied into quality.db (master stays warnings.db)")

    # Summary. is_open is now three-state (1 / 0 / NULL), so the counts must use
    # explicit IS NULL rather than 1 - is_open, which returns NULL for blanks.
    stats = pd.read_sql("""
        SELECT
            COUNT(*) AS total,
            SUM(CASE WHEN is_open = 1 THEN 1 ELSE 0 END) AS open_count,
            SUM(CASE WHEN is_open = 0 THEN 1 ELSE 0 END) AS closed_count,
            SUM(CASE WHEN is_open IS NULL THEN 1 ELSE 0 END) AS no_status_count,
            SUM(CASE WHEN source = 'tracker' THEN 1 ELSE 0 END) AS tracker_only,
            SUM(CASE WHEN source = 'sap' THEN 1 ELSE 0 END) AS sap_only,
            SUM(CASE WHEN source = 'both' THEN 1 ELSE 0 END) AS in_both,
            SUM(CASE WHEN project IS NULL THEN 1 ELSE 0 END) AS blank_project,
            SUM(CASE WHEN detection_area IS NULL THEN 1 ELSE 0 END) AS blank_detection,
            SUM(CASE WHEN owner IS NULL THEN 1 ELSE 0 END) AS blank_owner
        FROM nc
    """, conn)

    print("\n" + "=" * 60)
    print("Summary:")
    print("=" * 60)
    print(stats.T.to_string(header=False))

    # Owner counts — the number people actually check. Printed every run so a
    # split or a drop is visible immediately rather than reported by an owner.
    own = pd.read_sql("""
        SELECT COALESCE(owner,'(no owner)') AS owner,
               COUNT(*) AS total,
               SUM(CASE WHEN is_open=1 THEN 1 ELSE 0 END) AS open,
               SUM(CASE WHEN is_open IS NULL THEN 1 ELSE 0 END) AS no_status
        FROM nc WHERE source IN ('tracker','both')
        GROUP BY owner ORDER BY total DESC
    """, conn)
    print("\n" + "=" * 60)
    print("NCs per owner (from the tracker — SAP has no owner field):")
    print("=" * 60)
    print(own.to_string(index=False))

    # CAPA coverage — how many NCs in the DB actually carry each action type.
    # This is the number the dashboard's coverage KPI reports on.
    if capa is not None and not capa.empty:
        cov = pd.read_sql("""
            SELECT
                (SELECT COUNT(*) FROM nc) AS nc_total,
                COUNT(DISTINCT c.nc_id) AS ncs_with_any_capa,
                COUNT(DISTINCT CASE WHEN c.capa_type='RCA' THEN c.nc_id END) AS ncs_with_rca,
                COUNT(DISTINCT CASE WHEN c.capa_type='CA'  THEN c.nc_id END) AS ncs_with_ca,
                COUNT(DISTINCT CASE WHEN c.capa_type='PA'  THEN c.nc_id END) AS ncs_with_pa
            FROM capa c JOIN nc n ON n.nc_id = c.nc_id
        """, conn)
        print("\n" + "=" * 60)
        print("CAPA coverage (joined to nc):")
        print("=" * 60)
        print(cov.T.to_string(header=False))

        # Orphans: CAPA rows whose NC number doesn't exist in nc. If this is
        # large the foreign key isn't matching and the join is silently lossy.
        orph = pd.read_sql("""
            SELECT COUNT(DISTINCT c.nc_id) AS n
            FROM capa c LEFT JOIN nc n ON n.nc_id = c.nc_id
            WHERE n.nc_id IS NULL
        """, conn).iloc[0]["n"]
        if orph:
            print(f"\n  ⚠ {orph} NC number(s) in the CAPA tracker have no match in `nc`.")
            ex = pd.read_sql("""
                SELECT DISTINCT c.nc_id
                FROM capa c LEFT JOIN nc n ON n.nc_id = c.nc_id
                WHERE n.nc_id IS NULL LIMIT 10
            """, conn)["nc_id"].tolist()
            print(f"    e.g. {ex}")
            print("    Check the nc_id format matches on both sides before trusting coverage.")

    conn.close()
    print(f"\n✓ Database written to: {DB_FILE}")
    print(f"  Tracker: {len(tracker)} rows | SAP: {len(sap)} rows | Merged: {len(merged)} rows")
    if capa is not None and not capa.empty:
        print(f"  CAPA: {len(capa)} rows across {capa['nc_id'].nunique()} NCs")


if __name__ == "__main__":
    main()