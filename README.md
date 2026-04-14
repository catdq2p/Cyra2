# TPCRA Risk Assessment Dashboard

Streamlit dashboard built specifically for the **TPCRA Questionnaire** Excel format.

## Repo structure

```
├── app.py
├── requirements.txt
├── .streamlit/
│   └── config.toml
└── README.md
```

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Community Cloud

1. Push this repo to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**.
3. Select repo, branch `main`, main file `app.py`.
4. Click **Deploy**.

## Expected Excel format

| Column | Description |
|--------|-------------|
| Col A  | Question key (`A`, `A.1`, `B.2.1`, …) |
| Col B  | Question text / section heading |
| Col C  | Response (`Yes` / `No` / `Partial` / `N/A`) |

Section headers are single letters (`A`, `B`, … `L`).  
The app auto-detects vendor name, rep name, and email from contact rows.

## Dashboard tabs

| Tab | Contents |
|-----|----------|
| **Overview** | Stacked bar by section + donut chart + score bars |
| **By section** | Per-section drill-down with individual question cards |
| **Gap analysis** | All `No` and `Partial` responses flagged, exportable |
| **All responses** | Full filterable table, exportable as CSV or Excel |
