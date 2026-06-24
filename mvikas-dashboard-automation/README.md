# MVIKAS Dashboard Automation Backend

This project automates the daily update of your MVIKAS dashboard `script.js` from the Excel/Google Sheet format you shared.

**Python 3.14 compatible:** the backend uses `openpyxl` directly and does **not** require `pandas`, so it avoids Windows pandas build errors on newer Python versions.

It supports two workflows:

1. **Manual update**: upload an Excel file in the backend UI and it regenerates `script.js` immediately.
2. **Daily automatic update**: the backend or GitHub Actions downloads the Google Sheet every day, parses all required sheets, regenerates `script.js`, and optionally commits it to the GitHub Pages repository.

## What the backend reads

The processor understands these sheets from your sample Google Sheet:

- `Shipment Booked Yesterday` → booked-yesterday KPI and chart
- `Daily Tonnage` → daily tonnage KPI and chart/table values
- `Total_Tonnage_this_month` → month cumulative achieved tonnage
- `Tonnage_of_June_Month` → KAM/person, customer list, monthly targets, active days
- `Order due Tommorow` / `Order due Tomorrow` → due-tomorrow KPI and chart
- `Order EDD Crossed` → delayed/EDD-crossed KPI and chart
- `Open Shipment` → open shipment KPI and classification table

The generated JS updates all dashboard conditions currently present in your `index.html`:

- header date and elapsed-days badge
- all KPI cards
- open shipment table with delayed %, due count, and risk labels
- shipment classification charts
- tonnage-by-client progress bars
- target vs achieved charts
- daily average/completion table
- days-needed chart
- KAM performance table and donut chart

## Folder structure

```text
mvikas-dashboard-automation/
├── app.py                         # FastAPI backend + manual upload UI + scheduler
├── requirements.txt
├── .env.example
├── mvikas_backend/
│   ├── processor.py               # Excel/Google Sheet parser
│   ├── renderer.py                # script.js generator
│   ├── service.py                 # update orchestration + optional git commit
│   └── cli.py                     # command-line generator
├── static/
│   ├── index.html                 # cleaned dashboard HTML
│   ├── style.css
│   ├── script.js                  # generated output
│   └── latest_data.json           # generated parsed data for debugging
└── .github/workflows/update-dashboard.yml
```

## Local setup

```bash
cd mvikas-dashboard-automation
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
cp .env.example .env
```

Edit `.env` and set:

```env
MVIKAS_SOURCE_URL=https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit?usp=sharing
MVIKAS_STATIC_DIR=static
MVIKAS_UPDATE_TIME=09:30
MVIKAS_TIMEZONE=Asia/Kolkata
```

> The Google Sheet must be accessible to the machine running the backend. For the simplest setup, share/publish it so the XLSX export URL can be downloaded.

## Run backend

```bash
uvicorn app:app --reload --host 0.0.0.0 --port 8000
```

Open:

- backend/admin UI: <http://localhost:8000/>
- local dashboard preview: <http://localhost:8000/dashboard/>
- health check: <http://localhost:8000/health>

## Manual update from Excel upload

1. Open <http://localhost:8000/>
2. Upload the daily `.xlsx` file
3. Click **Generate dashboard JS**
4. New files are written to `static/script.js` and `static/latest_data.json`

## Manual update from Google Sheet URL

```bash
python -m mvikas_backend.cli \
  --source-url "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit?usp=sharing" \
  --static-dir static
```

Optional report-date override:

```bash
python -m mvikas_backend.cli --excel daily.xlsx --static-dir static --report-date 2026-06-18
```

## Deploy with your GitHub Pages dashboard

Your current dashboard is hosted from a static GitHub Pages repo, so a Python backend cannot run directly inside GitHub Pages. You have two good deployment options:

### Option A — GitHub Actions automation (recommended for GitHub Pages)

Copy these backend files into the root of your dashboard repo:

- `requirements.txt`
- `mvikas_backend/`
- `.github/workflows/update-dashboard.yml`

Then in GitHub repo settings:

1. Go to **Settings → Secrets and variables → Actions → New repository secret**
2. Add `MVIKAS_SOURCE_URL` with your Google Sheet URL
3. The workflow runs daily at **09:30 IST** and can also be run manually from the **Actions** tab
4. It writes `script.js` and `latest_data.json` at repo root and commits/pushes if anything changed

### Option B — Run backend on a server

Deploy this FastAPI app on a VPS/Render/Railway/etc. Set:

```env
MVIKAS_COMMIT_TO_GIT=true
MVIKAS_GIT_REPO_DIR=/path/to/cloned/MVIKAS-DASHBOARD
MVIKAS_GIT_BRANCH=main
```

The scheduler will generate and push updated `script.js` daily.

## Important dashboard note

Your uploaded `index.html` had a legacy inline script after `<script src="script.js"></script>` that hard-coded an older open-shipment table. I removed that from the cleaned `static/index.html` because the generated `script.js` now controls the complete dashboard dynamically.

If you use this with your existing GitHub repo, either remove that old inline block or keep only:

```html
<script src="script.js"></script>
</body>
</html>
```

## Customer name matching

The parser normalizes common naming differences so counts line up across sheets, for example:

- `Carrier - CTD` ↔ `Carrier CTD`
- `Oneric Appliances` ↔ `Oneiric Appliances Pvt Ltd`
- `Sukuga` ↔ `Sukuga Technologies Pvt Ltd`
- `MITRAS` ↔ `Mitras Technocrafts Pvt Ltd-HR`
- `Cosmos Pumps Pvt. Ltd.` ↔ `Cosmos Pumps Pvt Ltd`

Add more mappings in `mvikas_backend/processor.py` under `NAME_ALIASES` if a new customer name variant appears.
