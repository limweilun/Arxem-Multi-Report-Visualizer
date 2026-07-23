# Arxem Performance

Streamlit dashboard for consolidating multiple MT trade history Excel reports.

## Features

- Upload multiple `.xlsx` reports with **Upload Excel files**.
- Computes all performance metrics from the uploaded deal ledger rather than
  mixing broker summary values with locally calculated values:
	- net/gross profit and loss, costs/swap, return, and annualized return
	- max drawdown in dollars and percent
	- annualized Sharpe, Sortino, and Calmar ratios
	- recovery factor and maximum time to recovery
	- win rate, profit factor, trade count, expectancy, averages, and extremes
- Nets `Profit + Commission + Fee + Swap` into each completed position. Entry
  commissions are allocated to their completed trade and are never counted as
  separate losing trades.
- Uses zero-filled business-day returns against a fixed **$100,000** balance for
  Sharpe, Sortino, and annualization (`252` trading days, zero risk-free rate).
- Builds an equity curve for each report starting at **$100,000**.
- Calculates drawdown from completed-trade equity, including an initial $100,000
  anchor, so temporary entry-fee rows cannot create phantom drawdowns.
- Ensures final equity aligns with `100,000 + calculated Total Net Profit`.
- Visualizes:
	- Equity curve overlay across uploaded reports
	- Net profit vs max drawdown comparison
	- Raw `Win Rate`, `Profit Factor`, and `Sharpe Ratio` values on independent axes
	- Time to recovery comparison with readable labels
	- Secondary metrics comparison
- Exports the summary, equity overlay, completed trades, daily returns, raw
  deals, and charts with **Download Consolidate Report**.
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
- `Position`, `Symbol`, and `Volume` are used when available to reconstruct
  completed positions accurately. For exports without position identifiers,
  closing deals define the trades and unmatched entry costs are allocated by
  closing volume so totals still reconcile.

## Notes

- Win rate is expressed in percentage points (for example, `55.0` means 55%).
- Return and drawdown percentage fields are stored as decimal rates and formatted
  as percentages in the exported workbook.
- Time to recovery runs from the prior equity peak to recovery; an unrecovered
  drawdown runs through the final completed trade.
