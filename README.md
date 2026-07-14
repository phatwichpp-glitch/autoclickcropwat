# CropWat Auto-runner

Automates running [CropWat 8.0](https://www.fao.org/land-water/databases-and-software/cropwat/en/)
(a Windows desktop application) through many years × many candidate planting
dates, and merges the results into a single Excel workbook — replacing what
would otherwise be thousands of manual clicks.

See `cropwat_autorunner_spec.md` for the full design spec and background.

## Why this has to run locally

CropWat is a desktop application. The backend controls it directly via
Windows UI Automation (`pywinauto`) — clicking menus, typing into fields,
reading dialogs. This only works when the backend runs **on the same
Windows machine where CropWat is open**. There is no way to host this as a
remote/cloud service; every user runs their own local copy against their own
CropWat installation and their own data.

## How it works

1. **Automation engine** (`backend/automation/`) drives CropWat: opens the
   climate/rain files for a year, sets a candidate planting date, runs
   `Crop Water Requirements` → `Irrigation Scheduling`, and prints the result
   to a `.txt` file. One year can involve dozens of candidate planting dates
   (each a separate "what if I planted on this day" run).
2. **File engine** (`backend/file_engine/`) scans the input data folder
   (handles the real-world messy folder structure — decade-range
   sub-folders, inconsistent nesting — by parsing year/month straight out of
   filenames), resolves the correct climate/rain file per the shift-year
   rule, parses the `.txt` output, and builds/rebuilds an Excel "Result"
   sheet from whatever `.txt` files exist.
3. **Dashboard** (`frontend/`) — a local web UI for configuring the input/
   output folders, picking which planting dates to test via a calendar, and
   watching progress in real time over WebSocket.

These two phases are decoupled on purpose: Phase 1 (run CropWat → produce
`.txt` files) and Phase 2 (`.txt` → Excel) can be re-run independently, so a
partial or failed Phase 1 run never blocks rebuilding the Excel from
whatever data already exists.

## Project structure

```
backend/
  app.py                    FastAPI: REST API + WebSocket + serves frontend
  runner.py                 Background-thread orchestrator (year × candidate loop)
  state.py                  In-memory run state + pub/sub for WebSocket
  config.py                 Settings model (single input/output folder, planting calendar)
  cropwat_controls.py       All CropWat control identifiers (see note below)
  automation/
    cropwat_engine.py       pywinauto steps: open files, set date, calculate, print, screenshot
  file_engine/
    paths.py                Folder scanning + shift-year file resolution
    txt_parser.py           Parses CropWat's printed .txt output
    excel_writer.py         Builds the Result sheet from parsed .txt files
frontend/
  index.html / app.js / style.css   Dashboard + Setup UI
inspect_cropwat.py / inspect_menu.py   Scripts used to reverse-engineer CropWat's UI structure
```

## Running it (dev)

```bash
cd backend
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\uvicorn app:app --host 127.0.0.1 --port 8000
```

Open `http://127.0.0.1:8000`. Before starting a run:
1. **Open CropWat 8.0 and load your Crop and Soil files yourself first.**
   This is required, not optional — `File → New → Crop → Dry crop` (the only
   menu path to a "new" crop) creates a blank, empty crop definition that
   needs manual data entry, not a shortcut to opening an existing `.CRO`
   file. So there's no reliable way to automate this from a cold start; the
   engine fails fast with a clear message if Crop/Soil aren't already open
   when a run starts.
2. In the **ตั้งค่า** (Setup) tab, point "โฟลเดอร์ข้อมูลต้นทาง" at the folder
   containing your `Clim_*`/`Rain_*` station folders and crop/soil files, and
   scan
3. In the **Dashboard** tab, adjust the planting-date calendar and year range,
   then start the run

**Don't use `--reload`** — the automation runs in a background thread tied to
the process; a reload orphans it.

## Known unverified pieces

- `cropwat_controls.py` was built by inspecting one real CropWat 8.0
  installation. It should work identically for anyone running the same
  installation/build, but hasn't been tested against other versions.
- The `ERROR_DIALOG` message text can't be read from a dedicated control
  (CropWat draws it directly on the dialog) — errors are detected and
  dismissed, but the specific message text isn't always captured.
