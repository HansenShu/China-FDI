"""
China Global Investment Tracker (CGIT) — annual cumulative FDI stock proxy analysis

Uses Dataset 1 (Investments) only.

IMPORTANT INTERPRETATION
------------------------
The CGIT workbook is transaction-level investment data. It does not provide a true
FDI position/stock series because it does not systematically record divestments,
asset depreciation, write-downs, exchange-rate valuation changes, or market-price
revaluations. This script therefore constructs a GROSS CUMULATIVE INVESTMENT STOCK
PROXY by summing all recorded investment transaction values up to each year.

This is useful for studying how the accumulated geographical and sectoral footprint
of recorded Chinese overseas investment changes over time, but it should not be
labelled an official FDI stock measure.

Usage
-----
    python china_fdi_stock_analysis.py "China-Global-Investment-Tracker-public.xlsx"

Optional:
    python china_fdi_stock_analysis.py workbook.xlsx --output-dir cgti_fdi_output --top-n 10

Outputs
-------
- Cleaned annual investment flows (used only as inputs to the cumulative stock)
- Annual cumulative stock proxy by country
- Annual cumulative stock proxy by sector
- Country and sector shares of cumulative stock
- Summary tables for latest year
- PNG visualizations
"""

from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.chart import AreaChart, LineChart, Reference
from openpyxl.drawing.line import LineProperties


SHEET_NAME = "Dataset 1"
HEADER_ROW = 5  # Excel row 6 contains column names; pandas is zero-indexed.
VALUE_COL = "Quantity in Millions"


def add_formula_analysis_sheet(workbook: Path) -> None:
    """Add annual flow and cumulative stock-proxy tables to the source workbook."""
    wb = openpyxl.load_workbook(workbook)
    source = wb[SHEET_NAME]
    analysis_name = "FDI Stock Analysis"
    if analysis_name in wb.sheetnames:
        del wb[analysis_name]
    ws = wb.create_sheet(analysis_name)

    data_start = HEADER_ROW + 2
    data_end = source.max_row
    years = sorted({
        source.cell(row, 2).value
        for row in range(data_start, data_end + 1)
        if isinstance(source.cell(row, 2).value, (int, float))
    })
    countries = sorted({
        source.cell(row, 10).value
        for row in range(data_start, data_end + 1)
        if source.cell(row, 10).value not in (None, "")
    })
    sectors = sorted({
        source.cell(row, 8).value
        for row in range(data_start, data_end + 1)
        if source.cell(row, 8).value not in (None, "")
    })

    ws["A1"] = "China FDI annual analysis from Dataset 1"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = (
        "Flows are annual sums of Quantity in Millions. Stock is a gross cumulative "
        "recorded-investment proxy, not an official FDI position measure."
    )
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max(3, len(countries) + 1))

    def write_table(start_row: int, title: str, category_header: str, categories: list[str]) -> int:
        flow_row = start_row + 1
        stock_row = flow_row + len(years) + 2
        end_col = len(categories) + 1
        ws.cell(start_row, 1, title).font = Font(bold=True, size=12)
        ws.cell(flow_row, 1, "Annual flow (USD millions)").font = Font(bold=True)
        ws.cell(stock_row, 1, "Cumulative stock proxy (USD millions)").font = Font(bold=True)
        for col, category in enumerate(categories, start=2):
            ws.cell(flow_row, col, category)
            ws.cell(stock_row, col, category)
        for offset, year in enumerate(years, start=1):
            row = flow_row + offset
            ws.cell(row, 1, year)
            stock_data_row = stock_row + offset
            ws.cell(stock_data_row, 1, year)
            for col in range(2, end_col + 1):
                category_ref = f"{openpyxl.utils.get_column_letter(col)}${flow_row}"
                ws.cell(
                    row,
                    col,
                    f'=SUMIFS(\'{SHEET_NAME}\'!$E${data_start}:$E${data_end},'
                    f'\'{SHEET_NAME}\'!$B${data_start}:$B${data_end},$A{row},'
                    f'\'{SHEET_NAME}\'!${category_header}${data_start}:'
                    f'${category_header}${data_end},{category_ref})',
                )
                ws.cell(
                    stock_data_row,
                    col,
                    f'=SUMIFS(\'{SHEET_NAME}\'!$E${data_start}:$E${data_end},'
                    f'\'{SHEET_NAME}\'!$B${data_start}:$B${data_end},"<="&$A{stock_data_row},'
                    f'\'{SHEET_NAME}\'!${category_header}${data_start}:'
                    f'${category_header}${data_end},{category_ref})',
                )
        for row in range(flow_row, stock_row + len(years) + 1):
            for col in range(2, end_col + 1):
                ws.cell(row, col).number_format = '#,##0.0'
        return stock_row + len(years) + 2

    next_row = write_table(4, "By destination country", "J", countries)
    sector_start = next_row
    write_table(sector_start, "By investment sector", "H", sectors)

    def category_totals(category_col: int, categories: list[str]) -> list[str]:
        totals = {category: 0 for category in categories}
        for row in range(data_start, data_end + 1):
            category = source.cell(row, category_col).value
            value = source.cell(row, 5).value
            if category in totals and isinstance(value, (int, float)):
                totals[category] += value
        return sorted(categories, key=lambda category: totals[category], reverse=True)

    def add_stock_chart(
        anchor: str,
        title: str,
        stock_header_row: int,
        stock_first_row: int,
        categories: list[str],
        max_categories: int = 10,
    ) -> None:
        chart = LineChart()
        chart.title = title
        chart.style = 13
        chart.y_axis.title = "Cumulative recorded investment (USD millions)"
        chart.x_axis.title = "Year"
        chart.x_axis.numFmt = "0"
        chart.x_axis.tickLblPos = "low"
        chart.x_axis.majorTickMark = "out"
        chart.x_axis.delete = False
        chart.height = 9
        chart.width = 16
        years_ref = Reference(ws, min_col=1, min_row=stock_first_row, max_row=stock_first_row + len(years) - 1)
        for category in categories[:max_categories]:
            col = next(
                col for col in range(2, ws.max_column + 1)
                if ws.cell(stock_header_row, col).value == category
            )
            values = Reference(
                ws, min_col=col, min_row=stock_header_row,
                max_row=stock_first_row + len(years) - 1,
            )
            chart.add_data(values, titles_from_data=True)
        chart.set_categories(years_ref)
        chart.legend.position = "r"
        chart_colors = [
            "1F77B4", "FF7F0E", "2CA02C", "D62728", "9467BD",
            "8C564B", "E377C2", "7F7F7F", "BCBD22", "17BECF",
            "393B79", "637939", "8C6D31", "843C39",
        ]
        for index, series in enumerate(chart.ser):
            color = chart_colors[index % len(chart_colors)]
            series.graphicalProperties.line = LineProperties(solidFill=color)
            series.graphicalProperties.line.width = 24000
            series.graphicalProperties.line.noFill = False
        ws.add_chart(chart, anchor)

    country_stock_row = 4 + 1 + len(years) + 2
    sector_stock_row = sector_start + 1 + len(years) + 2
    country_ranked = category_totals(10, countries)
    sector_ranked = category_totals(8, sectors)
    add_stock_chart(
        "A100",
        "China's cumulative recorded investment by top destination countries",
        country_stock_row,
        country_stock_row + 1,
        country_ranked,
    )
    add_stock_chart(
        "J100",
        "China's cumulative recorded investment by sector",
        sector_stock_row,
        sector_stock_row + 1,
        sector_ranked,
        max_categories=len(sectors),
    )

    def write_composition_table(
        start_row: int,
        start_col: int,
        stock_header_row: int,
        stock_first_row: int,
        ranked_categories: list[str],
    ) -> tuple[int, int]:
        categories = ranked_categories[:10]
        table_categories = categories + ["Other"]
        ws.cell(start_row, start_col, "100% stacked area chart data").font = Font(bold=True, size=12)
        for offset, category in enumerate(table_categories):
            cell = ws.cell(start_row + 1, start_col + offset)
            cell.value = category
        for year_offset, year in enumerate(years):
            row = start_row + 2 + year_offset
            source_row = stock_first_row + year_offset
            ws.cell(row, start_col - 1, year)
            for category_offset, category in enumerate(categories):
                target = ws.cell(row, start_col + category_offset)
                source_col = next(
                    col for col in range(2, ws.max_column + 1)
                    if ws.cell(stock_header_row, col).value == category
                )
                target.value = f"={ws.cell(source_row, source_col).coordinate}"
            total_ref = f"SUM({openpyxl.utils.get_column_letter(start_col)}{row}:{openpyxl.utils.get_column_letter(start_col + len(categories) - 1)}{row})"
            ws.cell(
                row,
                start_col + len(categories),
                f"=SUM($B{stock_first_row + year_offset}:$"
                f"{openpyxl.utils.get_column_letter(len(ranked_categories) + 1)}{stock_first_row + year_offset})-{total_ref}",
            )
        return start_row + 1, start_row + 2

    def add_composition_chart(
        anchor: str,
        title: str,
        header_row: int,
        first_data_row: int,
        start_col: int,
        category_count: int,
    ) -> None:
        chart = AreaChart()
        chart.title = title
        chart.style = 13
        chart.grouping = "percentStacked"
        chart.overlap = 100
        chart.y_axis.title = "Share of cumulative stock (%)"
        chart.x_axis.title = "Year"
        chart.x_axis.numFmt = "0"
        chart.x_axis.tickLblPos = "low"
        chart.x_axis.delete = False
        chart.height = 9
        chart.width = 16
        data = Reference(
            ws,
            min_col=start_col,
            max_col=start_col + category_count - 1,
            min_row=header_row,
            max_row=first_data_row + len(years) - 1,
        )
        categories = Reference(ws, min_col=start_col - 1, min_row=first_data_row, max_row=first_data_row + len(years) - 1)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(categories)
        chart.legend.position = "r"
        chart_colors = [
            "1F77B4", "FF7F0E", "2CA02C", "D62728", "9467BD",
            "8C564B", "E377C2", "7F7F7F", "BCBD22", "17BECF",
            "393B79",
        ]
        for index, series in enumerate(chart.ser):
            color = chart_colors[index % len(chart_colors)]
            series.graphicalProperties.solidFill = color
            series.graphicalProperties.line.solidFill = color
        ws.add_chart(chart, anchor)

    country_comp_header, country_comp_first = write_composition_table(
        130,
        2,
        country_stock_row,
        country_stock_row + 1,
        country_ranked,
    )
    sector_comp_header, sector_comp_first = write_composition_table(
        130, 16, sector_stock_row, sector_stock_row + 1, sector_ranked
    )
    add_composition_chart(
        "A165",
        "Changing country composition of cumulative recorded investment",
        country_comp_header,
        country_comp_first,
        2,
        11,
    )
    add_composition_chart(
        "J165",
        "Changing sector composition of cumulative recorded investment",
        sector_comp_header,
        sector_comp_first,
        16,
        11,
    )

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    for row in ws.iter_rows():
        for cell in row:
            if cell.row in (5, next_row + 1) or (
                cell.value in ("Annual flow (USD millions)", "Cumulative stock proxy (USD millions)")
            ):
                cell.fill = header_fill
    ws.freeze_panes = "B6"
    ws.column_dimensions["A"].width = 34
    for col in range(2, ws.max_column + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 15
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(workbook)


def load_and_clean(workbook: Path) -> pd.DataFrame:
    """Load Dataset 1 and retain valid annual investment observations."""
    df = pd.read_excel(workbook, sheet_name=SHEET_NAME, header=HEADER_ROW)

    required = {"Year", VALUE_COL, "Country", "Sector"}
    missing = required.difference(df.columns)
    if missing:
        raise ValueError(f"Missing required columns in {SHEET_NAME}: {sorted(missing)}")

    out = df.copy()
    out["Year"] = pd.to_numeric(out["Year"], errors="coerce")
    out[VALUE_COL] = pd.to_numeric(out[VALUE_COL], errors="coerce")

    # A transaction needs a year and a reported investment value to enter the stock proxy.
    out = out.dropna(subset=["Year", VALUE_COL]).copy()
    out["Year"] = out["Year"].astype(int)

    # Keep unknown geography/sector rather than silently dropping transaction values.
    out["Country"] = out["Country"].fillna("Unspecified").astype(str).str.strip()
    out["Sector"] = out["Sector"].fillna("Unspecified").astype(str).str.strip()

    # Negative values, if ever present, would imply reductions; CGIT is mainly gross deals.
    # We preserve them rather than force them to zero.
    return out


def annual_panel(df: pd.DataFrame, category: str) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Return annual flows, cumulative stock proxy, and cumulative-stock shares."""
    years = np.arange(df["Year"].min(), df["Year"].max() + 1)

    flows = (
        df.groupby(["Year", category], as_index=False)[VALUE_COL]
        .sum()
        .pivot(index="Year", columns=category, values=VALUE_COL)
        .reindex(years, fill_value=0)
        .fillna(0)
        .sort_index()
    )

    stock = flows.cumsum()
    total_stock = stock.sum(axis=1)
    shares = stock.div(total_stock.replace(0, np.nan), axis=0) * 100

    return flows, stock, shares


def latest_summary(stock: pd.DataFrame, shares: pd.DataFrame) -> pd.DataFrame:
    """Rank categories by cumulative stock proxy in the latest year."""
    latest_year = int(stock.index.max())
    summary = pd.DataFrame(
        {
            "Stock proxy (USD millions)": stock.loc[latest_year],
            "Share of total stock (%)": shares.loc[latest_year],
        }
    )
    summary.index.name = "Category"
    summary = summary.sort_values("Stock proxy (USD millions)", ascending=False)
    summary.insert(0, "Year", latest_year)
    return summary


def top_categories(stock: pd.DataFrame, top_n: int) -> list[str]:
    """Choose top categories based on latest-year cumulative stock."""
    return stock.loc[stock.index.max()].nlargest(top_n).index.tolist()


def save_line_chart(stock: pd.DataFrame, categories: list[str], title: str,
                    ylabel: str, output: Path) -> None:
    fig, ax = plt.subplots(figsize=(11, 6.5))
    for category in categories:
        ax.plot(stock.index, stock[category] / 1000, linewidth=2, label=category)
    ax.set_title(title)
    ax.set_xlabel("Year")
    ax.set_ylabel(ylabel)
    ax.grid(alpha=0.25)
    ax.legend(frameon=False, ncol=2)
    fig.tight_layout()
    fig.savefig(output, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_stacked_share_chart(shares: pd.DataFrame, categories: list[str], title: str,
                             output: Path) -> None:
    # Top categories are shown individually; everything else is combined into "Other".
    plot_df = shares[categories].copy()
    plot_df["Other"] = 100 - plot_df.sum(axis=1)

    fig, ax = plt.subplots(figsize=(11, 6.5))
    ax.stackplot(plot_df.index, [plot_df[c] for c in plot_df.columns],
                 labels=plot_df.columns, alpha=0.9)
    ax.set_title(title)
    ax.set_xlabel("Year")
    ax.set_ylabel("Share of cumulative stock proxy (%)")
    ax.set_ylim(0, 100)
    ax.grid(axis="y", alpha=0.25)
    ax.legend(frameon=False, ncol=2, loc="upper left", bbox_to_anchor=(1.01, 1.0))
    fig.tight_layout()
    fig.savefig(output, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_latest_bar(summary: pd.DataFrame, top_n: int, title: str, output: Path) -> None:
    top = summary.head(top_n).sort_values("Stock proxy (USD millions)")
    fig, ax = plt.subplots(figsize=(10, 6.5))
    ax.barh(top.index, top["Stock proxy (USD millions)"] / 1000)
    ax.set_title(title)
    ax.set_xlabel("Cumulative recorded investment (USD billions)")
    ax.set_ylabel("")
    ax.grid(axis="x", alpha=0.25)
    fig.tight_layout()
    fig.savefig(output, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_total_stock_chart(df: pd.DataFrame, output: Path) -> pd.DataFrame:
    annual_flow = df.groupby("Year")[VALUE_COL].sum().sort_index()
    years = np.arange(df["Year"].min(), df["Year"].max() + 1)
    annual_flow = annual_flow.reindex(years, fill_value=0)
    cumulative = annual_flow.cumsum()

    total = pd.DataFrame({
        "Annual recorded investment flow (USD millions)": annual_flow,
        "Cumulative stock proxy (USD millions)": cumulative,
    })
    total.index.name = "Year"

    fig, ax = plt.subplots(figsize=(10.5, 6))
    ax.plot(total.index, total["Cumulative stock proxy (USD millions)"] / 1000, linewidth=2.5)
    ax.set_title("China's cumulative recorded overseas investment stock proxy")
    ax.set_xlabel("Year")
    ax.set_ylabel("USD billions")
    ax.grid(alpha=0.25)
    fig.tight_layout()
    fig.savefig(output, dpi=180, bbox_inches="tight")
    plt.close(fig)

    return total


def save_change_table(shares: pd.DataFrame, stock: pd.DataFrame) -> pd.DataFrame:
    """Show how category importance changed from the first to latest year."""
    first_year = int(stock.index.min())
    latest_year = int(stock.index.max())

    result = pd.DataFrame({
        f"Stock {first_year} (USD mn)": stock.loc[first_year],
        f"Stock {latest_year} (USD mn)": stock.loc[latest_year],
        f"Share {first_year} (%)": shares.loc[first_year],
        f"Share {latest_year} (%)": shares.loc[latest_year],
    })
    result["Change in share (percentage points)"] = (
        result[f"Share {latest_year} (%)"] - result[f"Share {first_year} (%)"]
    )
    return result.sort_values(f"Stock {latest_year} (USD mn)", ascending=False)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Analyze annual cumulative Chinese overseas investment using CGIT Dataset 1."
    )
    default_workbook = Path(__file__).resolve().parent / "China-Global-Investment-Tracker-public.xlsx"
    parser.add_argument(
        "workbook",
        type=Path,
        nargs="?",
        default=default_workbook,
        help=f"Path to the CGIT .xlsx workbook (default: {default_workbook.name})",
    )
    parser.add_argument("--output-dir", type=Path, default=Path("china_fdi_stock_output"))
    parser.add_argument("--top-n", type=int, default=10,
                        help="Number of leading countries/sectors to show in charts (default: 10)")
    args = parser.parse_args()

    if not args.workbook.exists():
        raise FileNotFoundError(args.workbook)
    args.output_dir.mkdir(parents=True, exist_ok=True)
    add_formula_analysis_sheet(args.workbook)

    df = load_and_clean(args.workbook)
    country_flows, country_stock, country_shares = annual_panel(df, "Country")
    sector_flows, sector_stock, sector_shares = annual_panel(df, "Sector")

    country_summary = latest_summary(country_stock, country_shares)
    sector_summary = latest_summary(sector_stock, sector_shares)

    country_change = save_change_table(country_shares, country_stock)
    sector_change = save_change_table(sector_shares, sector_stock)

    # Save analysis tables.
    df.to_csv(args.output_dir / "cleaned_dataset1_transactions.csv", index=False)
    country_flows.to_csv(args.output_dir / "annual_flows_by_country_usd_millions.csv")
    country_stock.to_csv(args.output_dir / "cumulative_stock_proxy_by_country_usd_millions.csv")
    country_shares.to_csv(args.output_dir / "cumulative_stock_share_by_country_percent.csv")
    country_summary.to_csv(args.output_dir / "latest_country_stock_summary.csv")
    country_change.to_csv(args.output_dir / "country_stock_share_change.csv")

    sector_flows.to_csv(args.output_dir / "annual_flows_by_sector_usd_millions.csv")
    sector_stock.to_csv(args.output_dir / "cumulative_stock_proxy_by_sector_usd_millions.csv")
    sector_shares.to_csv(args.output_dir / "cumulative_stock_share_by_sector_percent.csv")
    sector_summary.to_csv(args.output_dir / "latest_sector_stock_summary.csv")
    sector_change.to_csv(args.output_dir / "sector_stock_share_change.csv")

    total_stock = save_total_stock_chart(
        df, args.output_dir / "01_total_cumulative_stock_proxy.png"
    )
    total_stock.to_csv(args.output_dir / "annual_total_flow_and_stock_proxy.csv")

    top_countries = top_categories(country_stock, args.top_n)
    top_sectors = top_categories(sector_stock, min(args.top_n, sector_stock.shape[1]))

    save_line_chart(
        country_stock,
        top_countries,
        "China's cumulative recorded investment by destination country",
        "Cumulative recorded investment (USD billions)",
        args.output_dir / "02_stock_by_top_countries.png",
    )
    save_stacked_share_chart(
        country_shares,
        top_countries,
        "Changing country composition of China's cumulative recorded investment",
        args.output_dir / "03_country_stock_shares_over_time.png",
    )
    save_latest_bar(
        country_summary,
        args.top_n,
        f"Largest destination countries by cumulative recorded investment, {country_stock.index.max()}",
        args.output_dir / "04_latest_country_stock_ranking.png",
    )

    save_line_chart(
        sector_stock,
        top_sectors,
        "China's cumulative recorded overseas investment by sector",
        "Cumulative recorded investment (USD billions)",
        args.output_dir / "05_stock_by_sector.png",
    )
    save_stacked_share_chart(
        sector_shares,
        top_sectors,
        "Changing sector composition of China's cumulative recorded investment",
        args.output_dir / "06_sector_stock_shares_over_time.png",
    )
    save_latest_bar(
        sector_summary,
        min(args.top_n, len(sector_summary)),
        f"Largest sectors by cumulative recorded investment, {sector_stock.index.max()}",
        args.output_dir / "07_latest_sector_stock_ranking.png",
    )

    latest_year = int(df["Year"].max())
    first_year = int(df["Year"].min())
    latest_total_bn = total_stock.loc[latest_year, "Cumulative stock proxy (USD millions)"] / 1000

    print("\nCGIT annual cumulative investment analysis complete")
    print("---------------------------------------------------")
    print(f"Years: {first_year}–{latest_year}")
    print(f"Transactions used: {len(df):,}")
    print(f"Latest cumulative stock proxy: USD {latest_total_bn:,.1f} billion")
    print(f"Top country in {latest_year}: {country_summary.index[0]}")
    print(f"Top sector in {latest_year}: {sector_summary.index[0]}")
    print(f"Outputs saved to: {args.output_dir.resolve()}")
    print("\nNOTE: This is a gross cumulative investment stock proxy, not an official FDI position stock.")


if __name__ == "__main__":
    main()
