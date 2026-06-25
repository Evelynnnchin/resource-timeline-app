import io
import re
import calendar
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="T&C Resource Planning", layout="wide")
st.title("T&C Resource Planning Excel Generator")

MONTH_MAP = {m.upper(): i for i, m in enumerate(calendar.month_abbr) if m}
MONTH_MAP.update({m.upper(): i for i, m in enumerate(calendar.month_name) if m})

LINE_NAMES = ["DTL", "JRL", "CRL", "RTS"]

# =========================================================
# BASIC HELPERS
# =========================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    text = str(value).replace("\xa0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text

def to_number(value):
    if pd.isna(value):
        return None
    if isinstance(value, (int, float, np.integer, np.floating)):
        if np.isnan(value):
            return None
        return float(value)
    text = clean_text(value)
    if text in ["", "–", "—"]:
        return None
    try:
        return float(text.replace(",", ""))
    except Exception:
        return None

def parse_month(value):
    if pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.month
    text = clean_text(value).upper()
    if not text:
        return None
    text = text[:3]
    return MONTH_MAP.get(text)

def parse_year(value):
    if pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.year
    if isinstance(value, (int, float, np.integer, np.floating)):
        year = int(value)
        return year if 1900 <= year <= 2100 else None
    text = clean_text(value)
    match = re.search(r"(19\d{2}|20\d{2})", text)
    if match:
        return int(match.group(1))
    return None

def is_item_number(value):
    if pd.isna(value):
        return False
    if isinstance(value, (int, float, np.integer, np.floating)):
        return not np.isnan(value)
    text = clean_text(value)
    return bool(re.fullmatch(r"\d+(\.\d+)?", text))

def normalise_line(value):
    text = clean_text(value).upper()
    text = text.replace(" ", "")
    for line in LINE_NAMES:
        if text == line:
            return line
    return ""

def get_cell(df, row, col):
    if row >= df.shape[0] or col >= df.shape[1]:
        return None
    return df.iat[row, col]

def is_tbc_name(name):
    text = clean_text(name).upper()
    if text == "":
        return True
    if "TBC" in text:
        return True
    if text in ["TBD", "NA", "N/A"]:
        return True
    return False

# =========================================================
# MONTH HEADER PARSING
# =========================================================

def get_month_columns(tc_df):
    month_cols = []
    current_year = None
    for col in range(4, tc_df.shape[1]):
        year_here = parse_year(get_cell(tc_df, 0, col))
        if year_here:
            current_year = year_here
        month_num = parse_month(get_cell(tc_df, 1, col))
        if current_year and month_num:
            month_date = pd.Timestamp(current_year, month_num, 1)
            month_cols.append(
                {
                    "col_index": col,
                    "month": month_date,
                    "month_label": month_date.strftime("%b %y"),
                }
            )
    return month_cols

# =========================================================
# PARSE T&C ACTIVITIES
# =========================================================

def parse_tc_activities(tc_df):
    month_cols = get_month_columns(tc_df)
    if not month_cols:
        raise ValueError(
            "Cannot find month headers. Expected years on row 1 and months on row 2 from column E onwards."
        )

    resource_rows = []
    monthly_records = []

    current_line = ""
    current_item_no = ""
    current_item = ""
    current_requirement = ""
    current_parent_values = {}

    for r in range(2, tc_df.shape[0]):
        a = get_cell(tc_df, r, 0)
        b = clean_text(get_cell(tc_df, r, 1))
        c = clean_text(get_cell(tc_df, r, 2))
        d = clean_text(get_cell(tc_df, r, 3))

        line = normalise_line(a)
        if line:
            current_line = line
            current_item_no = ""
            current_item = ""
            current_requirement = ""
            current_parent_values = {}
            continue

        if not b and not c and not d:
            continue

        numbered_item = is_item_number(a)
        is_header = numbered_item or (not numbered_item and b and d and not c)

        if is_header:
            current_item_no = clean_text(a) if pd.notna(a) else ""
            current_item = b
            current_requirement = d

            current_parent_values = {}
            for m in month_cols:
                v = to_number(get_cell(tc_df, r, m["col_index"]))
                if v is not None:
                    current_parent_values[m["col_index"]] = v

            has_children = False
            for curr_r in range(r + 1, tc_df.shape[0]):
                nxt_a = get_cell(tc_df, curr_r, 0)
                nxt_b = clean_text(get_cell(tc_df, curr_r, 1))
                if normalise_line(nxt_a) or is_item_number(nxt_a):
                    break
                if nxt_b:
                    has_children = True
                    break

            if has_children:
                continue
            else:
                role = "Item level demand"
                assigned_name = "TBC"
                row_type = "Item level demand"
                row_requirement = d
                values_by_month = current_parent_values

        else:
            if not current_item:
                current_item = "General / Unassigned item"
                current_item_no = ""
                current_requirement = d

            role = b
            assigned_name = c if c else "TBC"
            row_type = "Role row"
            row_requirement = d if d else current_requirement

            values_by_month = {}
            for m in month_cols:
                v = to_number(get_cell(tc_df, r, m["col_index"]))
                if v is not None:
                    values_by_month[m["col_index"]] = v

            if not values_by_month:
                values_by_month = current_parent_values.copy()

        if not values_by_month:
            continue

        source_row = r + 1
        row_id = f"R{source_row}_{current_line}_{current_item_no}_{role}_{assigned_name}"

        load_list = []
        for m in month_cols:
            if m["col_index"] in values_by_month:
                load_list.append((m, values_by_month[m["col_index"]]))

        start_month = min([x[0]["month"] for x in load_list], default=pd.NaT)
        finish_month = max([x[0]["month"] for x in load_list], default=pd.NaT)

        resource_rows.append(
            {
                "Row ID": row_id,
                "Source Row": source_row,
                "Line": current_line,
                "Item No": current_item_no,
                "Activity / Item": current_item,
                "Item Requirement": row_requirement,
                "Role": role,
                "Current Name": assigned_name,
                "Assigned Name": assigned_name,
                "Row Type": row_type,
                "Start Month": start_month,
                "Finish Month": finish_month,
            }
        )

        for m, load in load_list:
            monthly_records.append(
                {
                    "Row ID": row_id,
                    "Source Row": source_row,
                    "Line": current_line,
                    "Item No": current_item_no,
                    "Activity / Item": current_item,
                    "Item Requirement": row_requirement,
                    "Role": role,
                    "Current Name": assigned_name,
                    "Month": m["month"],
                    "Month Label": m["month_label"],
                    "Load": load,
                }
            )

    resources = pd.DataFrame(resource_rows)
    monthly = pd.DataFrame(monthly_records)

    if resources.empty:
        raise ValueError("No resource rows found from T&C Activities.")

    return resources, monthly

# =========================================================
# PARSE MANPOWER TAB
# =========================================================

def parse_manpower(mp_df):
    records = []
    group_starts = [0, 3, 6, 9]

    for start_col in group_starts:
        line = clean_text(get_cell(mp_df, 1, start_col)).upper()
        if not line:
            continue

        for r in range(2, mp_df.shape[0]):
            name = clean_text(get_cell(mp_df, r, start_col))
            company = clean_text(get_cell(mp_df, r, start_col + 1))
            if not name:
                continue

            company_upper = company.upper()
            if company_upper in ["SC", "SUBCON", "SUBCONTRACTOR"]:
                company = "SC"
            elif company_upper == "SIEMENS":
                company = "Siemens"
            elif not company:
                company = "Unknown"

            records.append(
                {
                    "Line": line,
                    "Name": name,
                    "Company": company,
                }
            )

    manpower = pd.DataFrame(records)
    if manpower.empty:
        manpower = pd.DataFrame(columns=["Line", "Name", "Company"])
    manpower["Name Key"] = manpower["Name"].str.strip().str.upper()
    return manpower

# =========================================================
# ASSIGNMENT + COMPANY CLASSIFICATION
# =========================================================

def classify_company(row, name_to_company, infer_unknown=True):
    assigned = clean_text(row.get("Assigned Name", ""))
    role = clean_text(row.get("Role", ""))
    if is_tbc_name(assigned):
        return "TBC"
    key = assigned.upper()
    if key in name_to_company:
        return name_to_company[key]
    role_upper = role.upper()
    assigned_upper = assigned.upper()
    if (
        "SUBCON" in role_upper
        or assigned_upper.startswith("SUBCON")
        or assigned_upper == "SC"
    ):
        return "SC"
    if infer_unknown:
        return "Siemens"
    return "Unknown"

def apply_assignments(resources, monthly, assignments, manpower, infer_unknown=True):
    resources = resources.copy()
    monthly = monthly.copy()
    resources["Assigned Name"] = resources["Row ID"].map(assignments).fillna(resources["Assigned Name"])
    resources["Assigned Name"] = resources["Assigned Name"].apply(
        lambda x: clean_text(x) if clean_text(x) else "TBC"
    )
    resources["TBC?"] = resources["Assigned Name"].apply(
        lambda x: "YES" if is_tbc_name(x) else "NO"
    )
    resources["Role + Name"] = (
        resources["Role"].astype(str) + " / " + resources["Assigned Name"].astype(str)
    )
    if not manpower.empty:
        name_to_company = dict(zip(manpower["Name Key"], manpower["Company"]))
    else:
        name_to_company = {}
    resources["Company"] = resources.apply(
        lambda row: classify_company(row, name_to_company, infer_unknown),
        axis=1,
    )
    join_cols = resources[
        ["Row ID", "Assigned Name", "TBC?", "Role + Name", "Company"]
    ]
    monthly = monthly.drop(
        columns=[
            c
            for c in ["Assigned Name", "TBC?", "Role + Name", "Company"]
            if c in monthly.columns
        ],
        errors="ignore",
    )
    monthly = monthly.merge(join_cols, on="Row ID", how="left")
    return resources, monthly

# =========================================================
# DEMAND TABLES
# =========================================================

def make_demand_tables(monthly, total_threshold):
    if monthly.empty:
        return pd.DataFrame(), pd.DataFrame()
    demand_line = monthly.pivot_table(
        index="Month",
        columns="Line",
        values="Load",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()
    for line in LINE_NAMES:
        if line not in demand_line.columns:
            demand_line[line] = 0.0
    demand_line = demand_line[["Month"] + LINE_NAMES]
    demand_line["Total"] = demand_line[LINE_NAMES].sum(axis=1)
    demand_line["Threshold"] = float(total_threshold)
    demand_line["Month Label"] = demand_line["Month"].dt.strftime("%b %y")

    companies = ["Siemens", "SC", "TBC", "Unknown"]
    demand_company = monthly.pivot_table(
        index="Month",
        columns="Company",
        values="Load",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()
    for comp in companies:
        if comp not in demand_company.columns:
            demand_company[comp] = 0.0
    demand_company = demand_company[["Month"] + companies]
    demand_company["Total"] = demand_company[companies].sum(axis=1)
    demand_company["Threshold"] = float(total_threshold)
    demand_company["Month Label"] = demand_company["Month"].dt.strftime("%b %y")
    return demand_line, demand_company

def make_person_load(monthly, person_threshold=1.0):
    if monthly.empty:
        return pd.DataFrame(), pd.DataFrame()
    valid = monthly[~monthly["Assigned Name"].apply(is_tbc_name)].copy()
    if valid.empty:
        return pd.DataFrame(), pd.DataFrame()
    person_load = (
        valid.groupby(["Month", "Month Label", "Assigned Name", "Company"], as_index=False)["Load"]
        .sum()
        .sort_values(["Month", "Assigned Name"])
    )
    person_load["Person Threshold"] = float(person_threshold)
    person_load["Exceeded?"] = np.where(
        person_load["Load"] > float(person_threshold),
        "YES",
        "NO",
    )
    overload = person_load[person_load["Load"] > float(person_threshold)].copy()
    return person_load, overload

# =========================================================
# PLOTLY PREVIEW CHARTS
# =========================================================

def plot_line_chart(demand_line, threshold):
    fig = go.Figure()
    for line in LINE_NAMES:
        fig.add_trace(
            go.Bar(
                x=demand_line["Month Label"],
                y=demand_line[line],
                name=line,
                text=demand_line[line].round(2),
                textposition="inside",
            )
        )
    fig.add_trace(
        go.Scatter(
            x=demand_line["Month Label"],
            y=[threshold] * len(demand_line),
            mode="lines",
            name="Threshold",
            line=dict(dash="dash", width=3),
        )
    )
    fig.update_layout(
        barmode="stack",
        title="Monthly Manpower Demand by Line",
        xaxis_title="Month",
        yaxis_title="Manpower / No. of People",
        legend_title="Line",
        height=520,
    )
    return fig

def plot_company_chart(demand_company, threshold):
    fig = go.Figure()
    for comp in ["Siemens", "SC", "TBC", "Unknown"]:
        fig.add_trace(
            go.Bar(
                x=demand_company["Month Label"],
                y=demand_company[comp],
                name=comp,
                text=demand_company[comp].round(2),
                textposition="inside",
            )
        )
    fig.add_trace(
        go.Scatter(
            x=demand_company["Month Label"],
            y=[threshold] * len(demand_company),
            mode="lines",
            name="Threshold",
            line=dict(dash="dash", width=3),
        )
    )
    fig.update_layout(
        barmode="stack",
        title="Monthly Manpower Demand by Siemens / SC / TBC",
        xaxis_title="Month",
        yaxis_title="Manpower / No. of People",
        legend_title="Company / Status",
        height=520,
    )
    return fig

# =========================================================
# EXCEL OUTPUT
# =========================================================

def safe_sheet_name(name):
    text = re.sub(r"[\\/*?:\[\]]", "_", str(name))[:31]
    return text or "Sheet"

def write_df(writer, df, sheet_name, index=False):
    sheet_name = safe_sheet_name(sheet_name)
    df.to_excel(writer, sheet_name=sheet_name, index=index)
    ws = writer.sheets[sheet_name]
    wb = writer.book

    header_fmt = wb.add_format(
        {
            "bold": True,
            "bg_color": "#1F4E78",
            "font_color": "white",
            "border": 1,
        }
    )
    cell_fmt = wb.add_format({"border": 1})
    date_fmt = wb.add_format({"num_format": "mmm yy", "border": 1})

    for col_idx, col_name in enumerate(df.columns):
        ws.write(0, col_idx, col_name, header_fmt)
        if len(df):
            series = df[col_name].astype(str).replace("NaT", "")
            width = min(
                max(
                    len(str(col_name)) + 2,
                    int(series.str.len().quantile(0.9)) + 2,
                ),
                35,
            )
        else:
            width = len(str(col_name)) + 2
        ws.set_column(col_idx, col_idx, width)
        if pd.api.types.is_datetime64_any_dtype(df[col_name]):
            ws.set_column(col_idx, col_idx, 12, date_fmt)
    if len(df) > 0:
        ws.conditional_format(
            1,
            0,
            len(df),
            len(df.columns) - 1,
            {
                "type": "no_errors",
                "format": cell_fmt,
            },
        )
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(len(df), 1), max(len(df.columns) - 1, 0))

def add_stacked_chart(
    workbook,
    worksheet,
    sheet_name,
    first_data_row,
    last_data_row,
    cat_col,
    series_cols,
    threshold_col,
    title,
    anchor_cell,
):
    if last_data_row < first_data_row:
        return
    col_chart = workbook.add_chart(
        {
            "type": "column",
            "subtype": "stacked",
        }
    )
    for col_idx in series_cols:
        col_chart.add_series(
            {
                "name": [sheet_name, first_data_row - 1, col_idx],
                "categories": [sheet_name, first_data_row, cat_col, last_data_row, cat_col],
                "values": [sheet_name, first_data_row, col_idx, last_data_row, col_idx],
                "data_labels": {
                    "value": True,
                    "num_format": "0.0",
                },
            }
        )
    line_chart = workbook.add_chart({"type": "line"})
    line_chart.add_series(
        {
            "name": [sheet_name, first_data_row - 1, threshold_col],
            "categories": [sheet_name, first_data_row, cat_col, last_data_row, cat_col],
            "values": [sheet_name, first_data_row, threshold_col, last_data_row, threshold_col],
            "line": {
                "color": "#C00000",
                "width": 2.5,
                "dash_type": "dash",
            },
        }
    )
    col_chart.combine(line_chart)
    col_chart.set_title({"name": title})
    col_chart.set_x_axis(
        {
            "name": "Month",
            "num_font": {
                "rotation": -45,
            },
        }
    )
    col_chart.set_y_axis(
        {
            "name": "Manpower / No. of People",
            "major_gridlines": {
                "visible": True,
            },
        }
    )
    col_chart.set_legend({"position": "bottom"})
    col_chart.set_size({"width": 920, "height": 420})
    worksheet.insert_chart(anchor_cell, col_chart)

def create_excel_report(
    resources,
    monthly,
    demand_line,
    demand_company,
    person_load,
    overload,
    tbc_list,
    manpower,
    total_threshold,
    person_threshold,
):
    output = io.BytesIO()
    with pd.ExcelWriter(
        output,
        engine="xlsxwriter",
        datetime_format="mmm yy",
        date_format="mmm yy",
    ) as writer:
        workbook = writer.book
        dashboard = workbook.add_worksheet("Dashboard")
        writer.sheets["Dashboard"] = dashboard

        title_fmt = workbook.add_format(
            {
                "bold": True,
                "font_size": 18,
                "font_color": "#1F4E78",
            }
        )
        subtitle_fmt = workbook.add_format(
            {
                "bold": True,
                "font_size": 11,
                "font_color": "#666666",
            }
        )
        header_fmt = workbook.add_format(
            {
                "bold": True,
                "bg_color": "#1F4E78",
                "font_color": "white",
                "border": 1,
            }
        )
        note_fmt = workbook.add_format(
            {
                "italic": True,
                "font_color": "#666666",
            }
        )
        num_fmt = workbook.add_format(
            {
                "num_format": "0.0",
                "border": 1,
            }
        )
        text_fmt = workbook.add_format({"border": 1})
        date_fmt = workbook.add_format({"num_format": "mmm yy", "border": 1})
        red_fmt = workbook.add_format(
            {
                "bg_color": "#FFC7CE",
                "font_color": "#9C0006",
            }
        )

        dashboard.write("A1", "T&C Resource Planning Dashboard", title_fmt)
        dashboard.write("A2", "Generated from T&C Activities and Manpower tabs", subtitle_fmt)
        dashboard.write("A4", "Total manpower threshold from Manpower tab", header_fmt)
        dashboard.write("B4", total_threshold, num_fmt)
        dashboard.write("A5", "Per person overload threshold", header_fmt)
        dashboard.write("B5", person_threshold, num_fmt)
        dashboard.write(
            "A6",
            "SC = Subcon. Requirement text is carried from the whole activity item to the roles under it.",
            note_fmt,
        )

        # Dashboard table 1: Demand by Line
        dl = demand_line.copy()
        dl = dl[["Month", "DTL", "JRL", "CRL", "RTS", "Total", "Threshold"]]
        start_row = 8
        for col_idx, col in enumerate(dl.columns):
            dashboard.write(start_row, col_idx, col, header_fmt)
        for r_idx, row in dl.iterrows():
            excel_row = start_row + 1 + r_idx
            dashboard.write_datetime(excel_row, 0, row["Month"].to_pydatetime(), date_fmt)
            for col_idx, col in enumerate(dl.columns[1:], start=1):
                dashboard.write_number(excel_row, col_idx, float(row[col]), num_fmt)
        dashboard.set_column(0, 0, 12)
        dashboard.set_column(1, 6, 12)

        add_stacked_chart(
            workbook=workbook,
            worksheet=dashboard,
            sheet_name="Dashboard",
            first_data_row=start_row + 1,
            last_data_row=start_row + len(dl),
            cat_col=0,
            series_cols=[1, 2, 3, 4],
            threshold_col=6,
            title="Monthly Manpower Demand by Line",
            anchor_cell="I4",
        )

        # Dashboard table 2: Demand by Company
        dc = demand_company.copy()
        dc = dc[["Month", "Siemens", "SC", "TBC", "Unknown", "Total", "Threshold"]]
        comp_start_row = start_row + len(dl) + 4
        for col_idx, col in enumerate(dc.columns):
            dashboard.write(comp_start_row, col_idx, col, header_fmt)
        for r_idx, row in dc.iterrows():
            excel_row = comp_start_row + 1 + r_idx
            dashboard.write_datetime(excel_row, 0, row["Month"].to_pydatetime(), date_fmt)
            for col_idx, col in enumerate(dc.columns[1:], start=1):
                dashboard.write_number(excel_row, col_idx, float(row[col]), num_fmt)

        add_stacked_chart(
            workbook=workbook,
            worksheet=dashboard,
            sheet_name="Dashboard",
            first_data_row=comp_start_row + 1,
            last_data_row=comp_start_row + len(dc),
            cat_col=0,
            series_cols=[1, 2, 3, 4],
            threshold_col=6,
            title="Monthly Manpower Demand by Siemens / SC / TBC",
            anchor_cell=f"I{comp_start_row + 1}",
        )

        # Summary block
        summary_start = 8
        dashboard.write(summary_start, 15, "Summary", header_fmt)
        summary_pairs = [
            ("Total resource rows", len(resources)),
            ("Rows with TBC", len(tbc_list)),
            ("Overload person months", len(overload)),
            (
                "Peak monthly demand",
                float(demand_line["Total"].max()) if not demand_line.empty else 0,
            ),
        ]
        for idx, (label, value) in enumerate(summary_pairs, start=1):
            dashboard.write(summary_start + idx, 15, label, text_fmt)
            if isinstance(value, (int, float)):
                dashboard.write(summary_start + idx, 16, value, num_fmt)
            else:
                dashboard.write(summary_start + idx, 16, value, text_fmt)
        dashboard.set_column(15, 15, 24)
        dashboard.set_column(16, 16, 14)

        # Data sheets
        write_df(writer, resources, "Assignment_Overview")
        write_df(writer, monthly, "Monthly_Detail")
        write_df(writer, person_load, "Person_Load")
        write_df(writer, overload, "Overload_List")
        write_df(writer, tbc_list, "TBC_List")
        write_df(writer, demand_line, "Demand_by_Line")
        write_df(writer, demand_company, "Demand_by_Company")
        write_df(writer, manpower.drop(columns=["Name Key"], errors="ignore"), "Manpower_Master")

        if not overload.empty:
            ws = writer.sheets["Overload_List"]
            ws.conditional_format(
                1,
                0,
                len(overload),
                max(len(overload.columns) - 1, 0),
                {
                    "type": "no_blanks",
                    "format": red_fmt,
                },
            )
        if not tbc_list.empty:
            ws = writer.sheets["TBC_List"]
            ws.conditional_format(
                1,
                0,
                len(tbc_list),
                max(len(tbc_list.columns) - 1, 0),
                {
                    "type": "no_blanks",
                    "format": red_fmt,
                },
            )

    output.seek(0)
    return output.getvalue()

# =========================================================
# STREAMLIT APP
# =========================================================

uploaded_file = st.file_uploader(
    "Upload Resource Planning Excel file",
    type=["xlsx"],
)

if uploaded_file is None:
    st.info(
        "Upload the workbook here. The app reads only `T&C Activities` and `Manpower`, then generates a downloadable Excel report."
    )
    st.stop()

try:
    file_bytes = uploaded_file.getvalue()
    tc_df = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name="T&C Activities",
        header=None,
        engine="openpyxl",
    )
    mp_df = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name="Manpower",
        header=None,
        engine="openpyxl",
    )
    base_resources, base_monthly = parse_tc_activities(tc_df)
    manpower = parse_manpower(mp_df)

except Exception as e:
    st.error(
        "The workbook could not be parsed. Check that the sheets are named exactly `T&C Activities` and `Manpower`."
    )
    st.exception(e)
    st.stop()

if "assignments" not in st.session_state:
    st.session_state.assignments = dict(
        zip(base_resources["Row ID"], base_resources["Assigned Name"])
    )

# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:
    st.header("Filters / Settings")
    selected_lines = st.multiselect(
        "Line",
        LINE_NAMES,
        default=LINE_NAMES,
    )
    role_search = st.text_input(
        "Search role / activity / person",
        "",
    )
    show_only_tbc = st.checkbox(
        "Show only TBC rows",
        value=False,
    )
    infer_unknown = st.checkbox(
        "For names not in Manpower tab: infer non Subcon roles as Siemens",
        value=True,
    )
    default_threshold = int(len(manpower)) if len(manpower) else 1
    total_threshold = st.number_input(
        "Total manpower threshold",
        min_value=0.0,
        value=float(default_threshold),
        step=1.0,
    )
    person_threshold = st.number_input(
        "Per person overload threshold",
        min_value=0.0,
        value=1.0,
        step=0.1,
    )

# =========================================================
# APPLY ASSIGNMENTS
# =========================================================

resources, monthly = apply_assignments(
    base_resources,
    base_monthly,
    st.session_state.assignments,
    manpower,
    infer_unknown,
)

filtered = resources[resources["Line"].isin(selected_lines)].copy()

if role_search.strip():
    q = role_search.strip().lower()
    filtered = filtered[
        filtered["Role"].str.lower().str.contains(q, na=False)
        | filtered["Activity / Item"].str.lower().str.contains(q, na=False)
        | filtered["Assigned Name"].str.lower().str.contains(q, na=False)
    ]

if show_only_tbc:
    filtered = filtered[filtered["TBC?"] == "YES"]

# =========================================================
# ASSIGNMENT TABLE
# =========================================================

st.subheader("1. Assign person to each role")

st.write(
    "Edit only the `Assigned Name` column. Blank or `TBC` will be treated as TBC. "
    "Requirement is inherited from the whole activity item and is not counted as manpower by itself."
)

edit_cols = [
    "Row ID",
    "Line",
    "Source Row",
    "Item No",
    "Activity / Item",
    "Item Requirement",
    "Role",
    "Current Name",
    "Assigned Name",
    "Company",
    "TBC?",
    "Start Month",
    "Finish Month",
]

edited = st.data_editor(
    filtered[edit_cols],
    hide_index=True,
    use_container_width=True,
    num_rows="fixed",
    disabled=[c for c in edit_cols if c not in ["Assigned Name"]],
    column_config={
        "Row ID": None,
        "Assigned Name": st.column_config.TextColumn(
            "Assigned Name",
            help="Type the actual person name, or leave as TBC.",
        ),
        "Start Month": st.column_config.DateColumn(
            "Start Month",
            format="MMM YYYY",
        ),
        "Finish Month": st.column_config.DateColumn(
            "Finish Month",
            format="MMM YYYY",
        ),
    },
)

for _, row in edited.iterrows():
    st.session_state.assignments[row["Row ID"]] = (
        clean_text(row["Assigned Name"]) if clean_text(row["Assigned Name"]) else "TBC"
    )

resources, monthly = apply_assignments(
    base_resources,
    base_monthly,
    st.session_state.assignments,
    manpower,
    infer_unknown,
)

report_resources = resources[resources["Line"].isin(selected_lines)].copy()
report_monthly = monthly[monthly["Line"].isin(selected_lines)].copy()

demand_line, demand_company = make_demand_tables(
    report_monthly,
    total_threshold,
)

person_load, overload = make_person_load(
    report_monthly,
    person_threshold,
)

tbc_list = report_resources[report_resources["TBC?"] == "YES"].copy()

# =========================================================
# PREVIEW
# =========================================================

st.subheader("2. Preview manpower curves")

col1, col2, col3, col4 = st.columns(4)

col1.metric("Resource rows", len(report_resources))
col2.metric("TBC rows", len(tbc_list))
col3.metric("Overload person months", len(overload))

if not demand_line.empty:
    peak_monthly_demand = round(float(demand_line["Total"].max()), 2)
else:
    peak_monthly_demand = 0

col4.metric("Peak monthly demand", peak_monthly_demand)

if demand_line.empty:
    st.warning("No monthly manpower values found for the current filter.")
else:
    st.plotly_chart(
        plot_line_chart(demand_line, total_threshold),
        use_container_width=True,
    )

    st.plotly_chart(
        plot_company_chart(demand_company, total_threshold),
        use_container_width=True,
    )

# =========================================================
# CHECK TABLES
# =========================================================

st.subheader("3. Checks")

left, right = st.columns(2)

with left:
    st.write("TBC List")
    if tbc_list.empty:
        st.success("No TBC rows.")
    else:
        st.dataframe(
            tbc_list[
                [
                    "Line",
                    "Activity / Item",
                    "Item Requirement",
                    "Role",
                    "Assigned Name",
                    "Role + Name",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )

with right:
    st.write("Overload List")
    if overload.empty:
        st.success("No named person exceeds the per person threshold for the current filter.")
    else:
        st.dataframe(
            overload,
            use_container_width=True,
            hide_index=True,
        )

# =========================================================
# DOWNLOAD EXCEL
# =========================================================

st.subheader("4. Download Excel output")

try:
    excel_bytes = create_excel_report(
        resources=report_resources,
        monthly=report_monthly,
        demand_line=demand_line,
        demand_company=demand_company,
        person_load=person_load,
        overload=overload,
        tbc_list=tbc_list,
        manpower=manpower,
        total_threshold=total_threshold,
        person_threshold=person_threshold,
    )

    st.download_button(
        label="Download Excel Resource Planning Report",
        data=excel_bytes,
        file_name="T&C_Resource_Planning_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

except Exception as e:
    st.error("Excel output could not be generated.")
    st.exception(e)

with st.expander("Manpower master read from Manpower tab"):
    st.dataframe(
        manpower.drop(columns=["Name Key"], errors="ignore"),
        use_container_width=True,
        hide_index=True,
    )
