import os
import re
import sys
from datetime import datetime

import pandas as pd

# ============================================================
# RISK DASHBOARD AUTOMATION
# ============================================================
# How to run in VS Code terminal:
#   python risk_dashboard_automation.py "PRP Sample Jun (2).xlsx"
#
# If no input file is passed, the script uses INPUT_FILE below.
# ============================================================

INPUT_FILE = "PRP Sample Jun (2).xlsx"
SHEET_NAME = "OneTrust - Risk Export"
OUTPUT_FILE = "Risk_Output.xlsx"
HISTORY_FILE = "History.xlsx"

YEAR = 2026
OPEN_STAGES = ["Evaluation", "Identified", "Treatment"]
CLOSED_STAGE = "Monitoring"

# Fixed baseline values exactly as per your image/table
BASELINE_LABEL = "Baseline '25"
BASELINE_OPEN_RISK = 536
BASELINE_CLOSED_RISK = 42
BASELINE_TOTAL_RISK = 578

TARGET_LABEL = "Target '26"
TARGET_PERCENT = 0.80

# If History.xlsx does not exist, these starting values make the first output
# look like your sample image. Set this to False if you want History.xlsx empty.
SEED_HISTORY_WHEN_MISSING = True
SEED_ROWS = [
    {"Month": "Q1 '26", "Open risk as on date": 655, "Closed Risk in 2026": 41, "Total Risk": 696, "Risk Created in 2026": 118},
    {"Month": "Apr '26", "Open risk as on date": 694, "Closed Risk in 2026": 42, "Total Risk": 736, "Risk Created in 2026": 158},
]


def clean_column_name(value):
    """Remove hidden Excel characters such as carriage return and extra spaces."""
    return str(value).replace("\r", "").replace("\n", "").strip()


def normalize_stage(value):
    """Normalize stage values so filters do not fail because of spaces/hidden characters."""
    if pd.isna(value):
        return ""
    return str(value).replace("\r", "").replace("\n", "").strip()


def find_required_column(df, wanted_name):
    """Find required column even if Excel has hidden spaces/newline characters."""
    cleaned_map = {clean_column_name(col).lower(): col for col in df.columns}
    key = clean_column_name(wanted_name).lower()
    if key not in cleaned_map:
        raise ValueError(f"Required column '{wanted_name}' was not found. Available columns: {list(df.columns)}")
    return cleaned_map[key]


def get_period_label(df, date_created_col, date_closed_col):
    """
    Detect reporting month from the latest 2026 date available.
    This matches your sample where May'26 is calculated from the Jun sample file.
    """
    dates = []
    for col in [date_created_col, date_closed_col]:
        if col in df.columns:
            s = pd.to_datetime(df[col], errors="coerce")
            s = s[s.dt.year == YEAR]
            if not s.empty:
                dates.append(s.max())
    if not dates:
        raise ValueError(f"No valid {YEAR} dates found in Date created/Date closed columns.")

    latest_date = max(dates)
    month_abbr = latest_date.strftime("%b")
    if month_abbr == "Mar":
        return "Q1 '26"
    return f"{month_abbr} '26"


def calculate_metrics(input_file):
    """Read Excel and calculate Open, Closed, Total, Target, Risk Created."""
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file not found: {input_file}")

    excel_file = pd.ExcelFile(input_file, engine="openpyxl")
    if SHEET_NAME not in excel_file.sheet_names:
        raise ValueError(f"Sheet '{SHEET_NAME}' not found. Available sheets: {excel_file.sheet_names}")

    df = pd.read_excel(input_file, sheet_name=SHEET_NAME, engine="openpyxl")
    df.columns = [clean_column_name(col) for col in df.columns]

    stage_col = find_required_column(df, "Stage")
    date_created_col = find_required_column(df, "Date created")
    date_closed_col = find_required_column(df, "Date closed")

    df[stage_col] = df[stage_col].apply(normalize_stage)
    df[date_created_col] = pd.to_datetime(df[date_created_col], errors="coerce")
    df[date_closed_col] = pd.to_datetime(df[date_closed_col], errors="coerce")

    # Matches your screenshot logic:
    # Open risk as on date = all rows in open stages, not limited by year.
    open_risk = int(df[df[stage_col].isin(OPEN_STAGES)].shape[0])

    # Closed Risk in 2026 = Monitoring rows closed in YEAR.
    closed_risk = int(df[(df[stage_col] == CLOSED_STAGE) & (df[date_closed_col].dt.year == YEAR)].shape[0])

    total_risk = int(open_risk + closed_risk)

    # Risk Created in 2026 = all rows where Date created year is 2026.
    risk_created = int(df[df[date_created_col].dt.year == YEAR].shape[0])

    target_value = int(round(total_risk * TARGET_PERCENT, 0))
    period_label = get_period_label(df, date_created_col, date_closed_col)

    return {
        "Month": period_label,
        "Open risk as on date": open_risk,
        "Closed Risk in 2026": closed_risk,
        "Total Risk": total_risk,
        "Risk Created in 2026": risk_created,
        "Target 80%": target_value,
        "Closed %": closed_risk / total_risk if total_risk else 0,
        "Input File": os.path.basename(input_file),
        "Processed On": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def load_or_create_history():
    """Load History.xlsx or create it with optional seed values."""
    columns = [
        "Month",
        "Open risk as on date",
        "Closed Risk in 2026",
        "Total Risk",
        "Risk Created in 2026",
        "Target 80%",
        "Closed %",
        "Input File",
        "Processed On",
    ]

    if os.path.exists(HISTORY_FILE):
        history = pd.read_excel(HISTORY_FILE, engine="openpyxl")
        for col in columns:
            if col not in history.columns:
                history[col] = ""
        return history[columns]

    if SEED_HISTORY_WHEN_MISSING:
        history = pd.DataFrame(SEED_ROWS)
        history["Target 80%"] = (history["Total Risk"] * TARGET_PERCENT).round(0).astype(int)
        history["Closed %"] = history["Closed Risk in 2026"] / history["Total Risk"]
        history["Input File"] = "Seed from sample image"
        history["Processed On"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        return history[columns]

    return pd.DataFrame(columns=columns)


def update_history(metrics):
    """Add the month only if it does not already exist."""
    history = load_or_create_history()

    month = metrics["Month"]
    if month in history["Month"].astype(str).values:
        # Do not duplicate month. Keep old calculation exactly as requested.
        return history, False

    history = pd.concat([history, pd.DataFrame([metrics])], ignore_index=True)
    return history, True


def period_sort_key(label):
    """Sort Q1 first, then months Apr to Dec."""
    label = str(label)
    if label.startswith("Q1"):
        return 3
    month_order = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
                   "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}
    m = re.match(r"([A-Za-z]{3})", label)
    return month_order.get(m.group(1), 99) if m else 99


def build_dashboard_frames(history, current_total_for_target):
    """Create dashboard tables like your screenshot."""
    history = history.copy()
    history["_sort"] = history["Month"].apply(period_sort_key)
    history = history.sort_values("_sort").drop(columns=["_sort"])

    # Only last 3 periods appear in the chart/table between Baseline and Target.
    graph_history = history.tail(3).reset_index(drop=True)

    period_labels = graph_history["Month"].tolist()
    target_value = int(round(current_total_for_target * TARGET_PERCENT, 0))

    wide_columns = [BASELINE_LABEL] + period_labels + [TARGET_LABEL]

    row_open = [BASELINE_OPEN_RISK] + graph_history["Open risk as on date"].astype(int).tolist() + [""]
    row_closed = [BASELINE_CLOSED_RISK] + graph_history["Closed Risk in 2026"].astype(int).tolist() + [""]
    row_total = [BASELINE_TOTAL_RISK] + graph_history["Total Risk"].astype(int).tolist() + [""]
    row_blank = [""] * len(wide_columns)
    row_created = [0] + graph_history["Risk Created in 2026"].astype(int).tolist() + [""]
    row_baseline = [BASELINE_TOTAL_RISK] + [BASELINE_TOTAL_RISK] * len(period_labels) + [""]

    table_rows = [
        ["1. Open risk as on date"] + row_open,
        ["Closed Risk in 2026"] + row_closed,
        ["Total Risk"] + row_total,
        [""] + row_blank,
        ["Risk Created in 2026"] + row_created,
        [""] + row_baseline,
    ]
    table_df = pd.DataFrame(table_rows, columns=[""] + wide_columns)

    # Calculation table for chart
    chart_labels = [BASELINE_LABEL] + period_labels + [TARGET_LABEL]
    chart_values = [BASELINE_TOTAL_RISK] + graph_history["Open risk as on date"].astype(int).tolist() + [target_value]
    chart_percents = [""] + graph_history["Closed %"].astype(float).tolist() + [TARGET_PERCENT]

    calc_df = pd.DataFrame({
        "Label": chart_labels,
        "Value": chart_values,
        "Percent": chart_percents,
    })

    return table_df, calc_df


def write_files(history, table_df, calc_df):
    """Write History.xlsx and Risk_Output.xlsx with formatted table and chart."""
    history.to_excel(HISTORY_FILE, index=False, engine="openpyxl")

    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        workbook = writer.book
        ws = workbook.add_worksheet("Dashboard")
        writer.sheets["Dashboard"] = ws

        header_fmt = workbook.add_format({"bold": True, "bg_color": "#FFC000", "border": 1, "align": "center", "valign": "vcenter"})
        left_fmt = workbook.add_format({"bold": True, "bg_color": "#FFC000", "border": 1})
        num_fmt = workbook.add_format({"border": 1, "align": "right"})
        pct_fmt = workbook.add_format({"num_format": "0%", "border": 1, "align": "center"})
        peach_fmt = workbook.add_format({"bg_color": "#F4B183", "border": 1, "align": "center"})
        border_fmt = workbook.add_format({"border": 1})
        title_fmt = workbook.add_format({"bold": True, "font_size": 14})

        # Main matrix table
        for col_idx, col_name in enumerate(table_df.columns):
            ws.write(0, col_idx, col_name, header_fmt)
        for row_idx, row in table_df.iterrows():
            for col_idx, value in enumerate(row):
                if col_idx == 0:
                    ws.write(row_idx + 1, col_idx, value, left_fmt)
                else:
                    ws.write(row_idx + 1, col_idx, value, num_fmt)

        ws.set_column(0, 0, 28)
        ws.set_column(1, len(table_df.columns), 14)

        # Calculation table for chart
        start_row = 10
        ws.write(start_row, 0, "Calculation used for chart", title_fmt)
        for r, row in calc_df.iterrows():
            excel_row = start_row + 1 + r
            ws.write(excel_row, 0, row["Label"], peach_fmt)
            ws.write(excel_row, 1, row["Value"], border_fmt)
            if row["Percent"] == "":
                ws.write(excel_row, 2, "", border_fmt)
            else:
                ws.write_number(excel_row, 2, float(row["Percent"]), pct_fmt)

        ws.write(start_row + 1 + len(calc_df), 0, "Note", left_fmt)
        ws.write(start_row + 1 + len(calc_df), 1, "Target = 80% of latest Total Risk", border_fmt)

        # Chart helper headers
        chart_start = start_row + 10
        ws.write(chart_start, 0, "Chart Label")
        ws.write(chart_start, 1, "Chart Value")
        ws.write(chart_start, 2, "Percent Label")
        for r, row in calc_df.iterrows():
            ws.write(chart_start + 1 + r, 0, row["Label"])
            ws.write(chart_start + 1 + r, 1, row["Value"])
            if row["Percent"] == "":
                ws.write(chart_start + 1 + r, 2, "")
            else:
                ws.write(chart_start + 1 + r, 2, float(row["Percent"]), pct_fmt)

        chart = workbook.add_chart({"type": "column"})
        n = len(calc_df)
        points = []
        for i in range(n):
            if i == 0 or i == n - 1:
                points.append({"fill": {"color": "#BFBFBF"}, "border": {"color": "#BFBFBF"}})
            else:
                points.append({"fill": {"color": "#FFC000"}, "border": {"color": "#FFC000"}})

        chart.add_series({
            "name": "Cummulative Risk Treatment Progress",
            "categories": ["Dashboard", chart_start + 1, 0, chart_start + n, 0],
            "values": ["Dashboard", chart_start + 1, 1, chart_start + n, 1],
            "points": points,
            "data_labels": {"value": True, "position": "center", "font": {"color": "white", "bold": True}},
        })
        chart.set_title({"name": "Cummulative Risk Treatment\nProgress", "name_font": {"bold": True, "color": "white", "size": 18}})
        chart.set_legend({"none": True})
        chart.set_chartarea({"fill": {"color": "black"}, "border": {"color": "black"}})
        chart.set_plotarea({"fill": {"color": "black"}, "border": {"color": "black"}})
        chart.set_x_axis({"label_position": "low", "num_font": {"color": "white"}, "line": {"color": "white"}})
        chart.set_y_axis({"visible": False, "major_gridlines": {"visible": False}})
        chart.set_size({"width": 720, "height": 430})
        ws.insert_chart("G2", chart)

        # History sheet in output file too
        history.to_excel(writer, sheet_name="History Used", index=False)


def main():
    input_file = sys.argv[1] if len(sys.argv) > 1 else INPUT_FILE

    metrics = calculate_metrics(input_file)
    history, added = update_history(metrics)
    table_df, calc_df = build_dashboard_frames(history, metrics["Total Risk"])
    write_files(history, table_df, calc_df)

    print("Done.")
    print(f"Input file: {input_file}")
    print(f"Month calculated: {metrics['Month']}")
    print(f"Open risk as on date: {metrics['Open risk as on date']}")
    print(f"Closed Risk in {YEAR}: {metrics['Closed Risk in 2026']}")
    print(f"Total Risk: {metrics['Total Risk']}")
    print(f"Target 80%: {metrics['Target 80%']}")
    print(f"History updated with new month: {'Yes' if added else 'No - month already existed'}")
    print(f"Created/updated: {OUTPUT_FILE}")
    print(f"Created/updated: {HISTORY_FILE}")


if __name__ == "__main__":
    main()
