# Sprint Report Generator

A Streamlit app that converts a Jira CSV export into formatted sprint reports.
It generates Excel, PDF, and PNG summary outputs with sprint KPIs, status counts,
and an Epic -> Story/Task/Bug -> Sub-task hierarchy.

## Features

- Project-first sprint detail entry with local user-level persistence.
- Current-date defaults for sprint date fields.
- Jira CSV upload with required-column validation.
- Excel, landscape PDF, and PNG summary downloads.
- Multi-select download control that downloads selected files separately.
- Light Streamlit theme configured in `.streamlit/config.toml`.

## Local Persistence

Sprint details are saved locally to `data/sprint_details.local.json` on the machine
running Streamlit. This file is ignored by Git and should not be committed.

The app can still read the older `data/sprint_details.json` file as a migration
fallback, but new saves go to `data/sprint_details.local.json`.

This is suitable for local/user-level use. For shared team persistence after
deployment, replace this with a backend such as Google Sheets, a database, or an
object store.

## Jira CSV Requirements

Required columns:

| Column | Description |
| --- | --- |
| `Issue key` | Unique Jira ticket key |
| `Issue Type` | Epic / Story / Task / Bug / Sub-task |
| `Summary` | Ticket title |
| `Status` | Current Jira status |

Optional columns are handled safely when missing:

- `Priority`
- `Assignee`
- `Parent key`
- `Due date`
- `Custom field (Target start)`
- `Custom field (Target end)`
- `Comment`, `Comment.1`, etc.

## Run Locally

```powershell
git clone https://github.com/ashumishra2104/sprint_sheet_generator.git
cd sprint_sheet_generator
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

Open the URL Streamlit prints, usually `http://localhost:8501`.

## Deploy

For Streamlit Cloud:

1. Push only source files to GitHub.
2. Do not push `.venv/`, `Lib/`, `Scripts/`, `Include/`, `__pycache__/`, or `data/*.json`.
3. Create a new Streamlit Cloud app with `app.py` as the entry point.
4. Add a shared backend before deployment if sprint details must persist across users or redeploys.

## Contribution Checklist

Before opening a pull request:

```powershell
python -m py_compile app.py modules\parser.py modules\excel_generator.py modules\pdf_generator.py modules\image_generator.py
```

Also test a minimal Jira CSV with only the required columns and verify Excel, PDF,
and PNG downloads.
