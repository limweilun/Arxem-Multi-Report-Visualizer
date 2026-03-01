# Arxem Performance

Streamlit dashboard for consolidating multiple MT trade history Excel reports.

## Features

- Upload multiple `.xlsx` reports with **Upload Excel files**.
- Extracts and computes production-ready metrics per report:
	- `Total Net Profit`
	- `Max Drawdown ($)` from `Balance Drawdown Maximal`
	- `Win Rate` from `Profit Trades (% of total)`
	- `Profit Factor`
	- `Sharpe Ratio`
	- `Sortino Ratio`
	- `Recovery Factor`
	- `Time to Recovery` (`Xd Yh`) and numeric `Time to Recovery Days`
	- `Expectancy per Trade` (aligned to report `Expected Payoff`)
	- `Largest Single Loss`
	- `Total Trades`
- Uses deal-level net PnL (`Profit + Commission + Fee + Swap`) for equity/drawdown logic.
- Builds an equity curve for each report starting at **$100,000**.
- Ensures final equity aligns with `100,000 + Total Net Profit` from report summary.
- Visualizes:
	- Equity curve overlay across uploaded reports
	- Net profit vs max drawdown comparison
	- Relative quality profile (`Win Rate`, `Profit Factor`, `Sharpe Ratio`) as `% of best in current upload batch`
	- Time to recovery comparison with readable labels
	- Secondary metrics comparison
- Export consolidated analytics and charts with **Download Consolidate Report**.
- Maintains feature parity between the app's top summary table and exported `Summary` sheet.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Input format assumptions

This app expects MT-style trade history exports where:

- A `Deals` section exists with columns including `Time`, `Deal`, `Type`, `Direction`, `Profit`
- Optional columns `Commission`, `Fee`, and `Swap` are used when present
- Summary rows near the end include metrics such as `Total Net Profit`, `Profit Factor`, `Expected Payoff`, `Sharpe Ratio`, `Balance Drawdown`, and trade statistics

## Notes

- Relative Quality Profile is a within-batch ranking visualization. A value of `100` means best among uploaded files for that metric.
- Time to Recovery is shown in days/hours for readability and also charted in numeric days.
