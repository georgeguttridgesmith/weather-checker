"""
Tea Gardens Rainfall Data Processor
------------------------------------
Processes all daily Excel files and generates:
  1. Compiled raw data (all readings)
  2. Daily totals per garden
  3. Garden statistics summary
  4. Monthly summary
  5. Rainy event analysis

Usage:
    python process_tea_data.py <input_folder> <output_file.xlsx>

Example:
    python process_tea_data.py ~/Downloads/obubu-data/ tea_gardens_report.xlsx
"""

import sys
import os
import glob
import re
from datetime import datetime, timedelta

import pandas as pd
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import DataPoint


# ── Colours ──────────────────────────────────────────────────────────────────
C_HEADER    = "1F4E79"   # dark blue
C_SUBHEADER = "2E75B6"   # mid blue
C_ALT       = "D6E4F0"   # light blue row alt
C_ACCENT    = "70AD47"   # green highlight
C_WARN      = "FF0000"   # red (heavy rain)
C_WHITE     = "FFFFFF"
C_LIGHT     = "F2F2F2"


def thin_border():
    s = Side(border_style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)


def header_style(cell, bg=C_HEADER, font_size=11):
    cell.font = Font(bold=True, color=C_WHITE, size=font_size, name="Arial")
    cell.fill = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border()


def data_style(cell, bold=False, bg=None, align="right", color="000000"):
    cell.font = Font(bold=bold, color=color, size=10, name="Arial")
    if bg:
        cell.fill = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = thin_border()


def mm_fmt(cell, value, warn_threshold=None):
    cell.value = round(value, 2) if isinstance(value, float) else value
    cell.number_format = '#,##0.00'
    color = C_WARN if (warn_threshold and isinstance(value, (int, float)) and value >= warn_threshold) else "000000"
    data_style(cell, color=color)


# ── Data Loading ──────────────────────────────────────────────────────────────

def parse_date(d):
    s = str(int(d))
    return datetime.strptime(s, "%Y%m%d%H%M")


def load_all_files(folder):
    pattern = os.path.join(folder, "**", "*TeaGardensData.xlsx")
    files = sorted(glob.glob(pattern, recursive=True))
    if not files:
        pattern = os.path.join(folder, "*TeaGardensData.xlsx")
        files = sorted(glob.glob(pattern))
    print(f"Found {len(files)} files…")

    records = []
    garden_meta = {}

    for i, fpath in enumerate(files):
        if i % 50 == 0:
            print(f"  Processing file {i+1}/{len(files)}…")
        try:
            xl = pd.read_excel(fpath, sheet_name=None, header=0)
        except Exception as e:
            print(f"  Warning: could not read {fpath}: {e}")
            continue

        # Garden metadata
        if "Tea Gardens" in xl:
            meta = xl["Tea Gardens"]
            for _, row in meta.iterrows():
                name = row.get("Tea Garden Name")
                if pd.notna(name) and name not in garden_meta:
                    garden_meta[name] = {
                        "japanese": row.get("茶畑名前", ""),
                        "lat": row.get("Latitude"),
                        "lon": row.get("Longitude"),
                    }

        for sheet, df in xl.items():
            if sheet == "Tea Gardens":
                continue
            if df is None or df.empty:
                continue
            try:
                df.columns = ["Type", "Date", "Rainfall"]
                df = df.dropna(subset=["Date", "Rainfall"])
                df["datetime"] = df["Date"].apply(parse_date)
                df["date"]     = df["datetime"].dt.date
                df["garden"]   = sheet
                records.append(df[["garden", "datetime", "date", "Rainfall", "Type"]])
            except Exception:
                continue

    if not records:
        raise ValueError("No data found. Check the folder path.")

    raw = pd.concat(records, ignore_index=True)
    raw["Rainfall"] = pd.to_numeric(raw["Rainfall"], errors="coerce").fillna(0)
    return raw, garden_meta


# ── Aggregations ─────────────────────────────────────────────────────────────

def daily_totals(raw):
    return (
        raw.groupby(["garden", "date"])["Rainfall"]
        .sum()
        .reset_index()
        .rename(columns={"Rainfall": "daily_mm"})
    )


def garden_stats(daily):
    g = daily.groupby("garden")["daily_mm"]
    stats = pd.DataFrame({
        "Total Rainfall (mm)":      g.sum(),
        "Avg Daily Rainfall (mm)":  g.mean(),
        "Max Single Day (mm)":      g.max(),
        "Min Rainy Day (mm)":       daily[daily["daily_mm"] > 0].groupby("garden")["daily_mm"].min(),
        "Days with Rain":           (daily["daily_mm"] > 0).groupby(daily["garden"]).sum(),
        "Days Recorded":            g.count(),
        "Heavy Rain Days (>50mm)":  (daily["daily_mm"] > 50).groupby(daily["garden"]).sum(),
    }).reset_index().rename(columns={"garden": "Tea Garden"})
    stats["Rain Day %"] = stats["Days with Rain"] / stats["Days Recorded"] * 100
    return stats.sort_values("Total Rainfall (mm)", ascending=False)


def monthly_summary(daily):
    daily = daily.copy()
    daily["month"] = pd.to_datetime(daily["date"]).dt.to_period("M")
    return (
        daily.groupby(["month", "garden"])["daily_mm"]
        .sum()
        .reset_index()
        .pivot(index="month", columns="garden", values="daily_mm")
        .fillna(0)
    )


def daily_pivot(daily):
    return (
        daily.pivot(index="date", columns="garden", values="daily_mm")
        .fillna(0)
        .sort_index()
    )


def top_rain_events(daily, n=20):
    return (
        daily[daily["daily_mm"] > 0]
        .sort_values("daily_mm", ascending=False)
        .head(n)
        .copy()
    )


# ── Sheet Writers ─────────────────────────────────────────────────────────────

def write_cover(wb, daily, garden_meta):
    ws = wb.create_sheet("📋 Summary", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 30

    ws.merge_cells("A1:B1")
    ws["A1"] = "🍵 Tea Gardens Rainfall Report"
    ws["A1"].font = Font(bold=True, size=18, color=C_HEADER, name="Arial")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    ws.merge_cells("A2:B2")
    ws["A2"] = f"Generated: {datetime.now().strftime('%d %B %Y')}"
    ws["A2"].font = Font(size=11, color="666666", name="Arial")
    ws["A2"].alignment = Alignment(horizontal="center")

    info = [
        ("Total Files Processed", daily["date"].nunique()),
        ("Date Range",            f"{daily['date'].min()} → {daily['date'].max()}"),
        ("Tea Gardens",           daily["garden"].nunique()),
        ("Total Observations",    f"{len(daily):,}"),
    ]

    for r, (k, v) in enumerate(info, start=4):
        ws.cell(r, 1, k).font = Font(bold=True, size=11, name="Arial")
        ws.cell(r, 2, str(v)).font = Font(size=11, name="Arial")
        ws.cell(r, 2).alignment = Alignment(horizontal="right")

    # Garden list
    ws.cell(10, 1, "Gardens Included").font = Font(bold=True, size=11, color=C_HEADER, name="Arial")
    for i, g in enumerate(sorted(daily["garden"].unique()), start=11):
        meta = garden_meta.get(g, {})
        ws.cell(i, 1, g).font = Font(size=10, name="Arial")
        ws.cell(i, 2, meta.get("japanese", "")).font = Font(size=10, name="Arial")


def write_daily_totals(wb, pivot):
    ws = wb.create_sheet("📅 Daily Totals")
    ws.sheet_view.showGridLines = False

    gardens = list(pivot.columns)
    dates   = list(pivot.index)

    # Header row
    ws.cell(1, 1, "Date")
    header_style(ws.cell(1, 1))
    ws.column_dimensions["A"].width = 14

    for ci, g in enumerate(gardens, start=2):
        c = ws.cell(1, ci, g)
        header_style(c, bg=C_SUBHEADER, font_size=9)
        ws.column_dimensions[get_column_letter(ci)].width = 14

    # Total column
    tot_col = len(gardens) + 2
    ws.cell(1, tot_col, "Daily Total")
    header_style(ws.cell(1, tot_col))
    ws.column_dimensions[get_column_letter(tot_col)].width = 14

    # Data rows
    for ri, date in enumerate(dates, start=2):
        bg = C_ALT if ri % 2 == 0 else None
        ws.cell(ri, 1, str(date))
        data_style(ws.cell(ri, 1), bg=bg, align="center")

        for ci, g in enumerate(gardens, start=2):
            val = pivot.loc[date, g]
            c = ws.cell(ri, ci)
            mm_fmt(c, val, warn_threshold=50)
            if bg:
                c.fill = PatternFill("solid", start_color=bg)

        # Row total formula
        first_col = get_column_letter(2)
        last_col  = get_column_letter(len(gardens) + 1)
        ws.cell(ri, tot_col, f"=SUM({first_col}{ri}:{last_col}{ri})")
        data_style(ws.cell(ri, tot_col), bold=True)

    # Totals footer
    footer = len(dates) + 2
    ws.cell(footer, 1, "TOTAL")
    data_style(ws.cell(footer, 1), bold=True, bg=C_HEADER)
    ws.cell(footer, 1).font = Font(bold=True, color=C_WHITE, name="Arial")

    for ci in range(2, tot_col + 1):
        col_letter = get_column_letter(ci)
        ws.cell(footer, ci, f"=SUM({col_letter}2:{col_letter}{len(dates)+1})")
        ws.cell(footer, ci).number_format = '#,##0.00'
        ws.cell(footer, ci).font = Font(bold=True, color=C_WHITE, name="Arial")
        ws.cell(footer, ci).fill = PatternFill("solid", start_color=C_HEADER)
        ws.cell(footer, ci).border = thin_border()

    ws.freeze_panes = "B2"


def write_garden_stats(wb, stats):
    ws = wb.create_sheet("🌿 Garden Statistics")
    ws.sheet_view.showGridLines = False

    cols = list(stats.columns)
    widths = [22, 20, 22, 20, 20, 16, 16, 22, 14]

    for ci, (col, w) in enumerate(zip(cols, widths), start=1):
        c = ws.cell(1, ci, col)
        header_style(c)
        ws.column_dimensions[get_column_letter(ci)].width = w

    for ri, row in enumerate(stats.itertuples(index=False), start=2):
        bg = C_ALT if ri % 2 == 0 else None
        for ci, val in enumerate(row, start=1):
            c = ws.cell(ri, ci)
            if isinstance(val, float):
                c.value = round(val, 2)
                c.number_format = '#,##0.00'
            else:
                c.value = val
            data_style(c, bg=bg, align="left" if ci == 1 else "right")

    ws.freeze_panes = "A2"


def write_monthly(wb, monthly):
    ws = wb.create_sheet("📆 Monthly Summary")
    ws.sheet_view.showGridLines = False

    gardens = list(monthly.columns)
    months  = list(monthly.index)

    ws.cell(1, 1, "Month")
    header_style(ws.cell(1, 1))
    ws.column_dimensions["A"].width = 14

    for ci, g in enumerate(gardens, start=2):
        header_style(ws.cell(1, ci, g), bg=C_SUBHEADER, font_size=9)
        ws.column_dimensions[get_column_letter(ci)].width = 14

    tot_col = len(gardens) + 2
    header_style(ws.cell(1, tot_col, "Monthly Total"))
    ws.column_dimensions[get_column_letter(tot_col)].width = 16

    for ri, month in enumerate(months, start=2):
        bg = C_ALT if ri % 2 == 0 else None
        ws.cell(ri, 1, str(month))
        data_style(ws.cell(ri, 1), bg=bg, align="center")

        for ci, g in enumerate(gardens, start=2):
            val = monthly.loc[month, g]
            c = ws.cell(ri, ci)
            mm_fmt(c, val, warn_threshold=300)
            if bg:
                c.fill = PatternFill("solid", start_color=bg)

        fc = get_column_letter(2)
        lc = get_column_letter(len(gardens) + 1)
        ws.cell(ri, tot_col, f"=SUM({fc}{ri}:{lc}{ri})")
        data_style(ws.cell(ri, tot_col), bold=True)
        ws.cell(ri, tot_col).number_format = '#,##0.00'

    footer = len(months) + 2
    ws.cell(footer, 1, "TOTAL")
    data_style(ws.cell(footer, 1), bold=True, bg=C_HEADER)
    ws.cell(footer, 1).font = Font(bold=True, color=C_WHITE, name="Arial")
    for ci in range(2, tot_col + 1):
        cl = get_column_letter(ci)
        ws.cell(footer, ci, f"=SUM({cl}2:{cl}{len(months)+1})")
        ws.cell(footer, ci).number_format = '#,##0.00'
        ws.cell(footer, ci).font = Font(bold=True, color=C_WHITE, name="Arial")
        ws.cell(footer, ci).fill = PatternFill("solid", start_color=C_HEADER)
        ws.cell(footer, ci).border = thin_border()

    ws.freeze_panes = "B2"


def write_top_events(wb, events):
    ws = wb.create_sheet("⛈️ Top Rain Events")
    ws.sheet_view.showGridLines = False

    headers = ["Rank", "Tea Garden", "Date", "Rainfall (mm)", "Intensity"]
    widths  = [8, 25, 16, 18, 20]

    for ci, (h, w) in enumerate(zip(headers, widths), start=1):
        header_style(ws.cell(1, ci, h))
        ws.column_dimensions[get_column_letter(ci)].width = w

    for ri, row in enumerate(events.itertuples(index=False), start=2):
        bg = C_ALT if ri % 2 == 0 else None
        mm = row.daily_mm

        if mm >= 200:   intensity = "🌊 Extreme (≥200mm)"
        elif mm >= 100: intensity = "🌧️ Very Heavy (≥100mm)"
        elif mm >= 50:  intensity = "🌦️ Heavy (≥50mm)"
        elif mm >= 10:  intensity = "🌂 Moderate (≥10mm)"
        else:           intensity = "🌤️ Light"

        vals = [ri - 1, row.garden, str(row.date), round(mm, 2), intensity]
        for ci, val in enumerate(vals, start=1):
            c = ws.cell(ri, ci, val)
            data_style(c, bg=bg, align="center" if ci in (1, 3) else "left")
            if ci == 4:
                c.number_format = '#,##0.00'
                if mm >= 100:
                    c.font = Font(bold=True, color=C_WARN, size=10, name="Arial")


def write_raw(wb, raw):
    ws = wb.create_sheet("📊 Raw Data")
    ws.sheet_view.showGridLines = False

    headers = ["Garden", "Datetime", "Date", "Rainfall (mm)", "Type"]
    widths  = [22, 20, 14, 16, 14]

    for ci, (h, w) in enumerate(zip(headers, widths), start=1):
        header_style(ws.cell(1, ci, h))
        ws.column_dimensions[get_column_letter(ci)].width = w

    for ri, row in enumerate(raw.itertuples(index=False), start=2):
        bg = C_ALT if ri % 2 == 0 else None
        vals = [row.garden, str(row.datetime), str(row.date), row.Rainfall, row.Type]
        for ci, val in enumerate(vals, start=1):
            c = ws.cell(ri, ci, val)
            data_style(c, bg=bg, align="left" if ci in (1, 5) else "center")
            if ci == 4:
                c.number_format = '#,##0.00'

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{len(raw)+1}"


# ── Main ──────────────────────────────────────────────────────────────────────

def main(input_folder, output_path):
    print("Loading files…")
    raw, garden_meta = load_all_files(input_folder)
    print(f"Loaded {len(raw):,} readings across {raw['garden'].nunique()} gardens and {raw['date'].nunique()} days.")

    print("Aggregating…")
    daily   = daily_totals(raw)
    stats   = garden_stats(daily)
    monthly = monthly_summary(daily)
    pivot   = daily_pivot(daily)
    events  = top_rain_events(daily, n=30)

    print("Writing Excel report…")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    write_cover(wb, daily, garden_meta)
    write_daily_totals(wb, pivot)
    write_garden_stats(wb, stats)
    write_monthly(wb, monthly)
    write_top_events(wb, events)
    # write_raw(wb, raw.sort_values(["garden", "datetime"]))

    wb.save(output_path)
    print(f"\n✅ Report saved to: {output_path}")
    print(f"   Sheets: {', '.join(wb.sheetnames)}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
