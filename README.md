# Quality BRM Dashboard

Local web app for generating the monthly BRM (Business Review Meeting) reports.
Runs on your machine. No SharePoint. No cloud. Full backend control.

---

## What you get

- 📊 **Web dashboard** at `http://localhost:8501` with all charts for slides 2-4
- 📥 **Download button** → `.pptx` (4-slide BRM in Simone's style)
- 📥 **Download button** → `.xlsx` (multi-sheet workbook with all data)
- 🔍 **Data quality view** → shows which NCs need cleaning (for CW26)

---

## Folder structure

```
dash/
├── app.py                  ← Streamlit dashboard (main app)
├── ingest.py               ← Excel → SQLite loader
├── make_brm.py             ← PowerPoint generator
├── make_excel.py           ← Excel workbook generator
├── brm_queries.py          ← All SQL queries (used by app.py + terminal)
├── quality.db              ← auto-created SQLite database
├── requirements.txt        ← Python dependencies
└── data/
    ├── NCR_Cutover_Tracker_2026-07-08.xlsx   ← today's snapshot
    ├── NCR_Cutover_Tracker_2026-07-15.xlsx   ← next week's snapshot
    └── ...
```

---

## First-time setup (10 min, once)

In your `(quality)` venv:

```bash
cd C:\Users\BL3914\OneDrive - Beyond Gravity Services AG\Documents\Development\dash

# 1. Install dependencies (only pandas + openpyxl are new to you)
pip install -r requirements.txt

# 2. Create the data folder + drop in your first Excel snapshot
mkdir data
# → download NCR_Cutover_Tracker.xlsx from SharePoint
# → save as data\NCR_Cutover_Tracker_2026-07-08.xlsx

# 3. Load it into SQLite
python ingest.py

# 4. Launch the dashboard
streamlit run app.py
```

Then open `http://localhost:8501` in your browser.

---

## Weekly workflow (5 min)

**Before each weekly review / BRM:**

1. Download the latest Excel from SharePoint
2. Save it in `data/` as `NCR_Cutover_Tracker_YYYY-MM-DD.xlsx` (today's date)
3. In the dashboard sidebar → click **"Re-run ingest.py"**
   - Or from terminal: `python ingest.py`
4. Click **"Build PowerPoint"** → **"Download .pptx"**
5. Click **"Build Excel Workbook"** → **"Download .xlsx"**

Done.

---

## What each file does

| File | Purpose |
|------|---------|
| **ingest.py** | Reads newest `data/NCR_Cutover_Tracker*.xlsx`, cleans the messy fields (unifies "OPEN"/"Open"/"open", fixes owner name duplicates, parses dates), writes to `quality.db`. Idempotent — safe to re-run. |
| **make_brm.py** | Generates a 4-slide `.pptx` from the DB: Cover / Dashboard / Trends 1/2 / Trends 2/2. |
| **make_excel.py** | Generates a multi-sheet `.xlsx` from the DB: Summary / Top6 / By Area / Prod vs Supplier / Monthly Trend / Owner Workload / Data Quality / Full Data. |
| **brm_queries.py** | All the SQL queries as reusable Python functions. Run standalone (`python brm_queries.py`) to see the numbers in the terminal. |
| **app.py** | Streamlit web dashboard — glues everything together with a UI. |

---

## Things to check before sending to Simone

The generated `.pptx` is **auto-filled with real numbers** but still needs manual review:

- [ ] **Slide 2 – Delivery / Cost / People** sections say `[manual]` — Simone owns these numbers, add them by hand.
- [ ] **Slide 2 – Highlights / Lowlights / Outlook** — comment boxes are empty. Add commentary.
- [ ] **Slide 3 – "By Area" chart** — currently dominated by `BLANK - to clean` because 186 NCs are missing Detection. Fix during CW26 cleanup or hide the chart until then.
- [ ] Compare against Simone's May '26 deck — flag anything visually different you want to match.

---

## Data quality alerts

The app shows a yellow warning banner when:
- More than **20** NCs have blank Detection Area, OR
- More than **10** NCs have blank Project

Watch these numbers drop as CW26 cleanup progresses.

The **"Data Quality"** expander at the bottom of the dashboard lists the exact rows needing fixes — send this table to the owners.

---

## Adding more data sources later

Right now, only the `NC_Tracker_Black_Out` sheet is loaded. To add more later (e.g., audit data, CoPQ numbers, CAPA status):

1. Extend `ingest.py` — add a new function that reads a different sheet or file into a new SQLite table
2. Extend `brm_queries.py` — write the SQL you need
3. Extend `app.py` and `make_brm.py` to display / include it

The database schema is deliberately simple so you can grow it.

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `No NCR_Cutover_Tracker*.xlsx in data/` | Download from SharePoint and save into `data/` folder |
| `quality.db not found` | Run `python ingest.py` first |
| `ModuleNotFoundError` | `pip install -r requirements.txt` in your `(quality)` venv |
| Streamlit won't start | Check port 8501 isn't already in use, or run `streamlit run app.py --server.port=8502` |
| Charts empty | Check `python brm_queries.py` output — the DB may be empty |

---

## Later: connecting to SAP directly (skip SharePoint download)

Once you have SAP ODBC access, change the top of `ingest.py`:

```python
# Instead of reading Excel:
# df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

# Read directly from SAP:
import pyodbc
conn = pyodbc.connect("DSN=SAP_HANA;UID=youruser;PWD=yourpass")
df = pd.read_sql("SELECT * FROM SAP_NC_VIEW", conn)
```

Everything downstream keeps working exactly the same.
