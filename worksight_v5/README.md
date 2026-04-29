# WorkSight Pro React + Node

This version replaces the Streamlit UI with a React frontend and a Node/Express backend while keeping the existing Excel upload analytics workflows.

## Run locally

```bash
npm install
npm run dev
```

Frontend: <http://127.0.0.1:5173>

Backend API: <http://127.0.0.1:3001>

## Features

- Home navigation for the two original workflows.
- Work efficiency dashboard:
  - Upload ISC task data, iAMS attendance data, iWMS volume data, and punch data together.
  - Auto-classify uploaded workbooks.
  - KPI cards, efficiency donut, morning/evening Gantt timelines, group filters.
  - First action analysis, no-operation analysis, and indirect time timeline.
- Weekly quantity analysis:
  - Upload volume, ISC attendance, and optional picking result data.
  - Daily order/unit/UPPH summary.
  - Warehouse target UPPH rules.
  - Excel export.
  - Per-person picking efficiency table with delete marking.

## Notes

The backend uses `xlsx` to preserve broad Excel compatibility with the original app. `npm audit` reports known `xlsx` advisories with no fixed version currently available.
