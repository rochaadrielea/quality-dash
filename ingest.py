"""
ingest.py
---------
Reads NCR_Cutover_Tracker.xlsx AND NC_s_Overview (SAP export),
cleans, merges (deduplicates 173 overlapping NCs), loads to SQLite.

Two files in data/:
  - NCR_Cutover_Tracker*.xlsx  → active tracker (251 NCs, owner/disposition detail)
  - NC_s_Overview*.xlsx        → SAP full history (4498 NCs, CoPQ/leadtime/defect)

Merge logic:
  - 173 NCs exist in both → enriched: tracker fields + SAP fields, source='both'
  - 65+ tracker-only (EZ1, Blackout, TC) → source='tracker'
  - 4325 SAP-only (historical) → source='sap'

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
    if pd.isna(val) or not str(val).strip() or str(val).strip() == '\xa0':
        return None
    v = str(val).strip()
    aliases = {
        "Domingos": "Domingos Moreira",
        "Noel Kenoreny Orwa": "Noel Orwa",
        "S. Scampini": "Simone Scampini",
        "Vitaly /Salvatore": "Vitaly Meshin",
        "Vitaly/ Domingos": "Vitaly Meshin",
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


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1: Load the NCR Cutover Tracker
# ══════════════════════════════════════════════════════════════════════════════
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


def load_capa():
    """Load the CAPA / RCA tracker. OPTIONAL — returns None if the file isn't in data/.

    Keeps RCA rows only (one row per NC; CA/PA rows carry no root cause).
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
    out["capa_type"] = df[c_type].astype(str).str.strip() if c_type else None
    for tgt, src in [
        ("origin_area_l1", pick("(Real) Origin Area L1", "Origin Area L1")),
        ("origin_area_l2", pick("(Real) Origin Area L2", "Origin Area L2")),
        ("rc_category_l1", pick("RC Category L1")),
        ("rc_category_l2", pick("RC Category L2")),
    ]:
        out[tgt] = df[src].apply(clean_level) if src else None
    for tgt, src in [
        ("capa_project", pick("Affected Project")),
        ("requestor", pick("Requestor")),
        ("responsible", pick("Responsible")),
        ("capa_created_on", pick("Creation date requestor")),
    ]:
        out[tgt] = df[src].apply(clean_text) if src else None

    out = out[out["nc_id"].notna()].copy()

    # RCA rows only -> one row per NC, and the row that carries the root cause
    if c_type:
        rca = out[out["capa_type"].str.upper() == "RCA"].copy()
        if rca.empty:
            print("  ! No RCA rows found; keeping all CAPA rows instead.")
        else:
            out = rca
    out = out.drop_duplicates(subset=["nc_id"], keep="last")

    n_o1 = out["origin_area_l1"].notna().sum()
    n_r1 = out["rc_category_l1"].notna().sum()
    print(f"  {len(out)} CAPA rows (RCA, one per NC). "
          f"Origin L1 filled: {n_o1} | RC Category L1 filled: {n_r1}")
    return out


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

    # Column mapping — updated for 2026-07 tracker schema
    # (old schema names kept as fallbacks so older snapshots still load)
    col_map = {
        "System": "system",
        "ID-Blackout": "nc_id",            # was "Title"
        "Title": "nc_id",                  # fallback for old snapshots
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
    df["is_open"] = df["status"].isin(["OPEN"]).astype(int)
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
# STEP 2: Load the SAP NC Overview
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

    # Derived
    df["is_open"] = df["status"].isin(["OPEN"]).astype(int) if "status" in df.columns else 0
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
# STEP 3: Merge
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

    # Recalculate is_open for all rows (in case SAP data filled closure_date)
    merged.loc[merged["status"] == "CLOSED", "is_open"] = 0
    merged.loc[merged["status"] == "OPEN", "is_open"] = 1

    print(f"  Merged total: {len(merged)} rows")
    return merged


# ══════════════════════════════════════════════════════════════════════════════
# STEP 4: Write to SQLite
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
        cur.execute("CREATE INDEX IF NOT EXISTS idx_capa_nc ON capa(nc_id)")
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
def main():
    print("=" * 60)
    print("Quality DB Ingest — Merge Tracker + SAP Overview")
    print("=" * 60)

    print("\n[1/5] Loading NCR Cutover Tracker...")
    tracker = load_tracker()

    print("\n[2/5] Loading SAP NC Overview...")
    sap = load_sap_overview()

    print("\n[3/5] Merging (deduplicating overlaps)...")
    merged = merge_data(tracker, sap)

    print("\n[4/5] Loading CAPA / RCA tracker (optional)...")
    capa = load_capa()

    print("\n[5/5] Writing to SQLite...")
    conn = write_db(merged, capa)

    # Summary
    stats = pd.read_sql("""
        SELECT
            COUNT(*) AS total,
            SUM(is_open) AS open_count,
            SUM(1 - is_open) AS closed_count,
            SUM(CASE WHEN source = 'tracker' THEN 1 ELSE 0 END) AS tracker_only,
            SUM(CASE WHEN source = 'sap' THEN 1 ELSE 0 END) AS sap_only,
            SUM(CASE WHEN source = 'both' THEN 1 ELSE 0 END) AS in_both,
            SUM(CASE WHEN project IS NULL THEN 1 ELSE 0 END) AS blank_project,
            SUM(CASE WHEN detection_area IS NULL THEN 1 ELSE 0 END) AS blank_detection,
            SUM(CASE WHEN owner IS NULL THEN 1 ELSE 0 END) AS blank_owner,
            0 AS has_copq
            
        FROM nc
    """, conn)

    print("\n" + "=" * 60)
    print("Summary:")
    print("=" * 60)
    print(stats.T.to_string(header=False))

    conn.close()
    print(f"\n✓ Database written to: {DB_FILE}")
    print(f"  Tracker: {len(tracker)} rows | SAP: {len(sap)} rows | Merged: {len(merged)} rows")


if __name__ == "__main__":
    main()