from __future__ import annotations

from datetime import timedelta
from io import BytesIO
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from plotly import express as px

INITIAL_EQUITY = 100000.0


SUMMARY_KEYS = {
    "total net profit": "Total Net Profit",
    "profit factor": "Profit Factor",
    "expected payoff": "Expected Payoff",
    "recovery factor": "Recovery Factor",
    "balance drawdown absolute": "Balance Drawdown Absolute",
    "balance drawdown maximal": "Balance Drawdown Maximal",
    "balance drawdown relative": "Balance Drawdown Relative",
}


NUMERIC_METRICS = {
    "Total Net Profit",
    "Profit Factor",
    "Expected Payoff",
    "Recovery Factor",
    "Balance Drawdown Absolute",
}


def normalize_label(value: Any) -> str:
    return str(value).strip().lower().replace(":", "")


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace(",", "")
    if not text:
        return None

    if "(" in text and ")" in text:
        text = text.split("(")[0].strip()

    text = text.replace(" ", "")
    try:
        return float(text)
    except ValueError:
        return None


def find_best_value_in_row(row: tuple[Any, ...], label_index: int) -> Any:
    for value in row[label_index + 1 :]:
        if value not in (None, ""):
            return value
    return None


def extract_summary_metrics(ws) -> dict[str, Any]:
    metrics: dict[str, Any] = {
        "Total Net Profit": None,
        "Profit Factor": None,
        "Expected Payoff": None,
        "Recovery Factor": None,
        "Balance Drawdown Absolute": None,
        "Balance Drawdown Maximal": None,
        "Balance Drawdown Relative": None,
    }

    for row in ws.iter_rows(values_only=True):
        row_values = list(row)
        labels = [normalize_label(v) for v in row_values if v not in (None, "")]
        if not labels:
            continue

        for idx, raw in enumerate(row_values):
            if raw in (None, ""):
                continue
            label = normalize_label(raw)
            if label in SUMMARY_KEYS:
                metric_name = SUMMARY_KEYS[label]
                row_value = find_best_value_in_row(tuple(row_values), idx)
                if metric_name in NUMERIC_METRICS:
                    metrics[metric_name] = to_float(row_value)
                else:
                    metrics[metric_name] = row_value

    return metrics


def find_deals_header(ws) -> tuple[int, dict[str, int]]:
    max_col = ws.max_column
    for r in range(1, ws.max_row + 1):
        first_cell = ws.cell(r, 1).value
        if normalize_label(first_cell) == "deals":
            header_row = r + 1
            headers = [ws.cell(header_row, c).value for c in range(1, max_col + 1)]
            header_map = {
                normalize_label(name): idx + 1
                for idx, name in enumerate(headers)
                if name not in (None, "")
            }
            return header_row, header_map
    raise ValueError("Deals section not found")


def extract_deals_dataframe(ws) -> pd.DataFrame:
    header_row, header_map = find_deals_header(ws)
    required = ["time", "type", "direction", "profit"]
    for key in required:
        if key not in header_map:
            raise ValueError(f"Deals table is missing required column: {key}")

    rows: list[dict[str, Any]] = []
    r = header_row + 1

    while r <= ws.max_row:
        first_cell = ws.cell(r, 1).value
        if first_cell in (None, ""):
            # allow sparse blank rows after deals and before summary
            if normalize_label(ws.cell(r + 1, 1).value) in {
                "total net profit",
                "profit factor",
                "recovery factor",
                "balance drawdown",
            }:
                break
            r += 1
            continue

        first_label = normalize_label(first_cell)
        if first_label in {
            "total net profit",
            "profit factor",
            "recovery factor",
            "balance drawdown",
        }:
            break

        row_type = ws.cell(r, header_map["type"]).value
        direction = ws.cell(r, header_map["direction"]).value
        row_time = ws.cell(r, header_map["time"]).value
        row_profit = ws.cell(r, header_map["profit"]).value

        rows.append(
            {
                "Time": pd.to_datetime(row_time, errors="coerce"),
                "Type": str(row_type).lower() if row_type is not None else "",
                "Direction": str(direction).lower() if direction is not None else "",
                "Profit": to_float(row_profit) or 0.0,
            }
        )
        r += 1

    deals = pd.DataFrame(rows)
    deals = deals.dropna(subset=["Time"]).sort_values("Time").reset_index(drop=True)
    return deals


def build_equity_curve(deals: pd.DataFrame) -> pd.DataFrame:
    if deals.empty:
        return pd.DataFrame(columns=["Time", "Profit", "Equity"])

    realized = deals[
        (deals["Type"] != "balance")
        & ((deals["Direction"] == "out") | (deals["Profit"] != 0.0))
    ].copy()

    if realized.empty:
        start_time = deals["Time"].min()
        return pd.DataFrame(
            [
                {
                    "Time": start_time,
                    "Profit": 0.0,
                    "Equity": INITIAL_EQUITY,
                }
            ]
        )

    realized["Equity"] = INITIAL_EQUITY + realized["Profit"].cumsum()

    anchor_time = realized["Time"].min() - timedelta(seconds=1)
    anchor = pd.DataFrame(
        [{"Time": anchor_time, "Profit": 0.0, "Equity": INITIAL_EQUITY}]
    )

    return pd.concat([anchor, realized[["Time", "Profit", "Equity"]]], ignore_index=True)


def parse_report(uploaded_file) -> dict[str, Any]:
    workbook = load_workbook(BytesIO(uploaded_file.getvalue()), data_only=True, read_only=False)
    ws = workbook[workbook.sheetnames[0]]

    metrics = extract_summary_metrics(ws)
    deals = extract_deals_dataframe(ws)
    equity = build_equity_curve(deals)

    workbook.close()
    return {
        "name": uploaded_file.name,
        "metrics": metrics,
        "deals": deals,
        "equity": equity,
    }


def build_summary_table(reports: list[dict[str, Any]]) -> pd.DataFrame:
    rows = []
    for report in reports:
        row = {"Report": report["name"]}
        row.update(report["metrics"])
        rows.append(row)
    return pd.DataFrame(rows)


def build_equity_overlay_table(reports: list[dict[str, Any]]) -> pd.DataFrame:
    merged: pd.DataFrame | None = None
    for report in reports:
        curve = report["equity"][["Time", "Equity"]].rename(
            columns={"Equity": report["name"]}
        )
        if merged is None:
            merged = curve
        else:
            merged = merged.merge(curve, on="Time", how="outer")

    if merged is None:
        return pd.DataFrame(columns=["Time"])

    merged = merged.sort_values("Time")
    merged = merged.ffill()
    return merged


def build_drawdown_curve(equity_curve: pd.DataFrame) -> pd.DataFrame:
    if equity_curve.empty:
        return pd.DataFrame(columns=["Time", "Drawdown"])

    curve = equity_curve[["Time", "Equity"]].copy().sort_values("Time")
    curve["Peak"] = curve["Equity"].cummax()
    curve["Drawdown"] = (curve["Peak"] - curve["Equity"]).clip(lower=0)
    return curve[["Time", "Drawdown"]]


def make_download_workbook(reports: list[dict[str, Any]]) -> bytes:
    summary_df = build_summary_table(reports)
    equity_overlay = build_equity_overlay_table(reports)

    all_deals = []
    for report in reports:
        deals = report["deals"].copy()
        deals.insert(0, "Report", report["name"])
        all_deals.append(deals)
    deals_df = pd.concat(all_deals, ignore_index=True) if all_deals else pd.DataFrame()

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm:ss") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        equity_overlay.to_excel(writer, sheet_name="Equity Overlay", index=False)
        deals_df.to_excel(writer, sheet_name="Deals", index=False)

        workbook = writer.book
        summary_sheet = writer.sheets["Summary"]
        equity_sheet = writer.sheets["Equity Overlay"]
        charts_sheet = workbook.add_worksheet("Charts")

        currency_fmt = workbook.add_format({"num_format": "$#,##0.00"})
        for col_name in ["Total Net Profit", "Balance Drawdown Absolute"]:
            if col_name in summary_df.columns:
                col_idx = summary_df.columns.get_loc(col_name)
                summary_sheet.set_column(col_idx, col_idx, 20, currency_fmt)

        if "Balance Drawdown Relative" in summary_df.columns:
            col_idx = summary_df.columns.get_loc("Balance Drawdown Relative")
            summary_sheet.set_column(col_idx, col_idx, 24)

        equity_sheet.set_column(0, 0, 22)
        equity_sheet.set_column(1, max(1, len(equity_overlay.columns) - 1), 18, currency_fmt)

        row_count = len(summary_df)
        if row_count > 0 and "Total Net Profit" in summary_df.columns:
            chart_profit = workbook.add_chart({"type": "column"})
            metric_col = summary_df.columns.get_loc("Total Net Profit")
            chart_profit.add_series(
                {
                    "name": "Total Net Profit",
                    "categories": ["Summary", 1, 0, row_count, 0],
                    "values": ["Summary", 1, metric_col, row_count, metric_col],
                    "data_labels": {"value": True},
                }
            )
            chart_profit.set_title({"name": "Total Net Profit by Report"})
            chart_profit.set_y_axis({"name": "$", "num_format": "$#,##0"})
            chart_profit.set_legend({"none": True})
            charts_sheet.insert_chart("B2", chart_profit, {"x_scale": 1.2, "y_scale": 1.2})

        if row_count > 0 and "Balance Drawdown Absolute" in summary_df.columns:
            chart_drawdown = workbook.add_chart({"type": "column"})
            dd_col = summary_df.columns.get_loc("Balance Drawdown Absolute")
            chart_drawdown.add_series(
                {
                    "name": "Balance Drawdown Absolute",
                    "categories": ["Summary", 1, 0, row_count, 0],
                    "values": ["Summary", 1, dd_col, row_count, dd_col],
                    "data_labels": {"value": True},
                }
            )
            chart_drawdown.set_title({"name": "Absolute Drawdown by Report"})
            chart_drawdown.set_y_axis({"name": "$", "num_format": "$#,##0"})
            chart_drawdown.set_legend({"none": True})
            charts_sheet.insert_chart("B20", chart_drawdown, {"x_scale": 1.2, "y_scale": 1.2})

        if len(equity_overlay.columns) > 1 and len(equity_overlay) > 1:
            chart_equity = workbook.add_chart({"type": "line"})
            for col_idx in range(1, len(equity_overlay.columns)):
                chart_equity.add_series(
                    {
                        "name": ["Equity Overlay", 0, col_idx],
                        "categories": ["Equity Overlay", 1, 0, len(equity_overlay), 0],
                        "values": ["Equity Overlay", 1, col_idx, len(equity_overlay), col_idx],
                    }
                )
            chart_equity.set_title({"name": "Equity Curve Overlay (Start $100,000)"})
            chart_equity.set_y_axis({"name": "Equity ($)", "num_format": "$#,##0"})
            chart_equity.set_x_axis({"name": "Time"})
            charts_sheet.insert_chart("K2", chart_equity, {"x_scale": 1.5, "y_scale": 1.4})

    output.seek(0)
    return output.getvalue()


def render_dashboard(reports: list[dict[str, Any]]) -> None:
    summary_df = build_summary_table(reports)

    st.subheader("Key Metrics by Report")
    display_summary = summary_df.copy()
    display_summary.index = list(range(1, len(display_summary) + 1))
    display_summary.index.name = "Row"
    st.dataframe(display_summary, use_container_width=True)

    st.subheader("Profit and Drawdown Comparison")
    metric_cols = [c for c in ["Total Net Profit", "Balance Drawdown Absolute"] if c in summary_df.columns]
    if metric_cols:
        melted = summary_df.melt(
            id_vars=["Report"],
            value_vars=metric_cols,
            var_name="Metric",
            value_name="Value",
        )
        fig_bar = px.bar(
            melted,
            x="Report",
            y="Value",
            color="Metric",
            barmode="group",
            title="Total Net Profit vs Absolute Drawdown",
            labels={"Value": "$"},
        )
        fig_bar.update_layout(legend_title_text="Metric")
        st.plotly_chart(fig_bar, use_container_width=True)

    st.subheader("Equity Curve Overlay (Start $100,000)")
    curves = []
    for report in reports:
        curve = report["equity"].copy()
        curve["Report"] = report["name"]
        curves.append(curve)

    if curves:
        long_equity = pd.concat(curves, ignore_index=True)
        fig_line = px.line(
            long_equity,
            x="Time",
            y="Equity",
            color="Report",
            title="Equity Curves Across Uploaded Reports",
            labels={"Equity": "Equity ($)"},
        )
        st.plotly_chart(fig_line, use_container_width=True)

    st.subheader("Drawdown Analysis")
    drawdowns = []
    for report in reports:
        dd_curve = build_drawdown_curve(report["equity"])
        if dd_curve.empty:
            continue
        dd_curve["Report"] = report["name"]
        drawdowns.append(dd_curve)

    if drawdowns:
        long_drawdown = pd.concat(drawdowns, ignore_index=True)
        fig_drawdown = px.line(
            long_drawdown,
            x="Time",
            y="Drawdown",
            color="Report",
            title="Drawdown Amount Over Time",
            labels={"Drawdown": "Drawdown ($)"},
        )
        st.plotly_chart(fig_drawdown, use_container_width=True)


def main() -> None:
    st.set_page_config(page_title="Arxem Report Visualizer", layout="wide")
    st.title("Arxem Multi-Report Visualizer")

    uploaded_files = st.file_uploader(
        "Upload Excel files",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Upload one or more MT trade history reports in .xlsx format.",
    )

    if not uploaded_files:
        st.info("Upload Excel files to generate analytics and charts.")
        return

    reports = []
    parsing_errors = []
    for file in uploaded_files:
        try:
            reports.append(parse_report(file))
        except Exception as exc:
            parsing_errors.append(f"{file.name}: {exc}")

    if parsing_errors:
        st.error("Some files could not be parsed:")
        for error in parsing_errors:
            st.write(f"- {error}")

    if not reports:
        st.warning("No valid reports were parsed.")
        return

    render_dashboard(reports)

    consolidated = make_download_workbook(reports)
    st.download_button(
        "Download Consolidate Report",
        data=consolidated,
        file_name="consolidated_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )


if __name__ == "__main__":
    main()
