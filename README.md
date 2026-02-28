# Broadway Portfolio Command Center

Interactive Streamlit dashboard for Account Managers to monitor brand performance, content metrics, LIVE streaming, and action items using Broadway Tool data exports.

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Data Source

Place your Broadway Tool export (`.xlsm`) in the `data/` folder as `broadway_data.xlsm`, or use the file uploader in the sidebar.

### Required Sheets
- **Raw** — Weekly brand-level metrics (SPS, ratings, content, compliance)
- **Partner Raw** — GMV, revenue, and traffic data (joined on Brand + report date)
- **AM Seller Mapping** — Account Manager to brand assignments

## Features

| Tab | What it shows |
|-----|--------------|
| **Overview** | GMV breakdown, SPS sub-scores, performance trends, portfolio distribution |
| **Shop Health** | Brand health table, SPS radar, compliance flags, heatmap over time |
| **Content & Creators** | Creator funnel, collaboration/sampling rates, GPM trends, video scatter |
| **LIVE** | Stream metrics, LIVE GMV, duration vs GPM, engagement treemap |
| **Action Items** | Auto-flagged critical issues and warnings with severity breakdown |
| **Data** | Filtered data table with CSV export |

## Filters

- **Account Manager** — filter to your portfolio
- **Brand** — drill into a single brand (dynamically filtered by AM)
- **Report Week** — select any available week
- **Account Type / Priority / Category** — multi-select filters

## Tech Stack

- [Streamlit](https://streamlit.io) — app framework
- [Plotly](https://plotly.com/python/) — interactive charts
- [pandas](https://pandas.pydata.org) — data processing
- [openpyxl](https://openpyxl.readthedocs.io) — Excel file parsing
