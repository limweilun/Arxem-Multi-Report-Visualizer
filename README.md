# Arxem Performance

Streamlit dashboard for consolidating multiple MT trade history Excel reports.

## Features

- Upload multiple `.xlsx` reports with **Upload Excel files**.
- Extracts key analytics per report, including:
	- `Total Net Profit`
	- `Profit Factor`
	- `Expected Payoff`
	- `Recovery Factor`
	- `Balance Drawdown Absolute`
	- `Balance Drawdown Maximal`
	- `Balance Drawdown Relative`
- Builds an equity curve for each report starting at **$100,000**.
- Visualizes:
	- Profit vs drawdown comparison
	- Overlay equity curve chart across reports
- Export consolidated analytics and overlay charts with **Download Consolidate Report**.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Input format assumptions

This app expects MT-style trade history exports where:

- A `Deals` section exists with columns including `Time`, `Type`, `Direction`, `Profit`
- Summary rows near the end include `Total Net Profit` and `Balance Drawdown` fields
