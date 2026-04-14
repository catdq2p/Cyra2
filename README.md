# Third-Party Risk Assessment Dashboard

A Streamlit app that accepts an Excel questionnaire upload and renders a full interactive risk dashboard.

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

Opens at `http://localhost:8501`.

## Deploy to Streamlit Community Cloud

1. Push this repo to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**.
3. Select your repo, branch `main`, main file `app.py`.
4. Click **Deploy**.

## Excel format

Your file needs these columns (names are flexible — the app auto-detects common variations):

| Column     | Accepted names                                      | Required |
|------------|-----------------------------------------------------|----------|
| Category   | `Category`, `Domain`, `Area`, `Control Area`        | Yes      |
| Risk       | `Risk`, `Description`, `Finding`, `Issue`           | Yes      |
| Severity   | `Severity`, `Level`, `Risk Level`, `Rating`         | Yes      |
| Vendor     | `Vendor`, `Company`, `Supplier`                     | Optional |
| Questionnaire | `Questionnaire`, `Assessment`, `Framework`       | Optional |

Severity values must be **High**, **Medium**, or **Low** (case-insensitive).

Download the sample file from the sidebar inside the app for a ready-to-use template.

## Features

- Auto-detects column names — no renaming needed in most cases
- Posture score calculated from weighted severity counts
- Stacked bar chart by category (Plotly)
- Donut chart for severity distribution
- Filterable risk table by severity and category
- Export filtered results as CSV or Excel
