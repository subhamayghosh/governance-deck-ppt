# LPL QE Governance Deck Generator

A Flask web app that ingests the monthly **LPL QE Governance Excel workbook** and
the **PPT template** and produces a fully populated, release-ready Governance Deck
in seconds — no manual transcription.

## What it automates

Per-slide updates from the Excel data (deterministic, format-preserving):

| Slides | Source sheet | What's updated |
|---|---|---|
| 6–9 | `ProjectDetails` | Pentagon sub-domain labels + SVT/Non-SVT bar charts. Static feature descriptions are preserved. |
| 10 | `MonthlySheet` | Scripts distribution pie chart (Feature/Develop/Main, DashBoard>0, Java); UI/API & Python script totals |
| 11 | `DailyExecution` | Critical/Core/All + P0–P4 charts; "Testcases configured in nightly batch"; Pass rate. Filtered to Frequency ∈ {Daily, Weekly} |
| 12 | `InSprintData` | Automation snapshot |
| 13–14 | `DefectData` | Defect views |
| 16–17 | `RADEnabled` | RAD snapshot |
| 18–19 | `ReleaseDayTestCaseSheet` | Primary release automation |
| 20–21 | `TCsDetailsSheet` | Regression automation trends |

## Requirements

- Python 3.10+
- A filled-in monthly Excel workbook (`.xlsm` / `.xlsx`)
- The PPT template (`.pptx`) for the deck

> The Excel and PPT files are **not** included in this repo (they contain LPL
> internal data). Obtain them from the team SharePoint / one-drive.

## Setup

```bash
# 1. Clone
git clone https://github.com/subhamayghosh/governance-deck-ppt.git
cd governance-deck-ppt

# 2. (Optional) create a virtual environment
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt
```

## Run

```bash
python webapp/app.py
```

Then open <http://localhost:5000/> in your browser.

## Usage

1. Click **Upload Excel** and select the latest monthly governance workbook.
2. Click **Upload PPT** and select the deck template.
3. Click **Generate PPT** — the populated `.pptx` will be downloaded.

The active uploads persist across server restarts (`webapp/uploads/.active_files.json`),
so you only need to re-upload when the source files change.

## Project layout

```
.
├── README.md
├── requirements.txt
├── .gitignore
└── webapp/
    ├── app.py                  # Flask routes, upload handling, Excel extraction
    ├── slide_updaters.py       # Per-slide update logic (charts + text)
    ├── templates/index.html    # Upload + generate UI
    ├── static/css, static/js
    └── uploads/                # Holds the active Excel / PPT (gitignored)
```

## Architecture (one-paragraph)

`app.py` is a thin Flask layer that handles uploads, persists which Excel/PPT are
"active" in `uploads/.active_files.json`, and reads each Excel tab into Python
dicts via pandas + openpyxl. On **Generate**, it opens the PPT template with
`python-pptx`, dispatches each slide through a registry in `slide_updaters.py`
(`SLIDE_UPDATERS = {6: update_slide_6, 7: …}`), and streams the resulting deck
back as an in-memory `.pptx` download. All updates are deterministic, preserve
fonts/colors/layout, and skip non-automatable slides.

## Adding a new slide

1. Add `def update_slide_NN(slide, data): ...` to `webapp/slide_updaters.py`.
2. Register it: `SLIDE_UPDATERS[NN] = update_slide_NN`.
3. Use the shared helpers — `update_chart_data`, `_set_text_preserve_format`,
   `_iter_shapes` (recursive group traversal), `_num` — to keep formatting safe.

## Notes

- The app **never** modifies the uploaded Excel or PPT files; all generation
  happens against in-memory copies.
- Generated PPTs are streamed to the browser; no intermediate `.pptx` files are
  written to disk.
- LPL data files (`.xlsm`, `.xlsx`, `.pptx`) and the local active-files state
  are gitignored.
