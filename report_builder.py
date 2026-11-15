"""Build a portable Excel workbook with cleaned data and schema-driven analysis."""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLUMNS = 16_384
MAX_CHARTS = 8
MAX_CATEGORY_CHART_VALUES = 20

_HEADER_FILL = PatternFill("solid", fgColor="205B78")
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_TITLE_FONT = Font(color="17324D", bold=True, size=16)
_SUBTITLE_FONT = Font(color="536878", italic=True)


def _safe_excel_value(value: object) -> object:
    if value is None or value is pd.NA:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if hasattr(value, "item"):
        value = value.item()
    if isinstance(value, str):
        return value[:32767]
    return value


def _is_risky_excel_string(value: str) -> bool:
    if value.startswith(("=", "@")):
        return True
    if value == "-" or re.fullmatch(r"[+-]\d+(?:[.,]\d+)?", value):
        return False
    if value.startswith("+") and re.fullmatch(r"\+\d[\d\s().-]*", value):
        return False
    return value.startswith(("+", "-"))


def _set_excel_cell(cell, value: object) -> None:
    safe_value = _safe_excel_value(value)
    cell.value = safe_value
    if isinstance(safe_value, str) and _is_risky_excel_string(safe_value):
        cell.data_type = "s"
        cell.quotePrefix = True


def _append_safe_row(sheet, values: list[object]) -> None:
    sheet.append([_safe_excel_value(value) for value in values])
    for column_index, value in enumerate(values, start=1):
        if isinstance(value, str):
            _set_excel_cell(sheet.cell(sheet.max_row, column_index), value)


def _write_frame(
    sheet,
    frame: pd.DataFrame,
    start_row: int = 1,
    column_types: dict[str, str] | None = None,
) -> int:
    if len(frame.columns) > EXCEL_MAX_COLUMNS:
        raise ValueError(f"{sheet.title} exceeds Excel's {EXCEL_MAX_COLUMNS:,}-column limit.")
    if start_row + len(frame) > EXCEL_MAX_ROWS:
        raise ValueError(f"{sheet.title} exceeds Excel's {EXCEL_MAX_ROWS:,}-row limit.")

    for column_index, name in enumerate(frame.columns, start=1):
        cell = sheet.cell(start_row, column_index)
        _set_excel_cell(cell, str(name))
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    formats = {
        "DATE": "yyyy-mm-dd hh:mm:ss",
        "PERCENT": "0.00%",
        "MONEY": "#,##0.00",
        "NUMBER": "#,##0.00",
        "INTEGER": "#,##0",
    }
    column_types = column_types or {}
    for row_index, row in enumerate(frame.itertuples(index=False, name=None), start=start_row + 1):
        sheet.append([_safe_excel_value(value) for value in row])
        for column_index, (name, value) in enumerate(zip(frame.columns, row), start=1):
            cell = sheet.cell(row_index, column_index)
            if isinstance(value, str):
                _set_excel_cell(cell, value)
            number_format = formats.get(column_types.get(name, ""))
            if number_format:
                cell.number_format = number_format

    sheet.freeze_panes = f"A{start_row + 1}"
    sheet.auto_filter.ref = (
        f"A{start_row}:{get_column_letter(len(frame.columns))}{start_row + len(frame)}"
    )
    sheet.row_dimensions[start_row].height = 30
    for column_index, name in enumerate(frame.columns, start=1):
        values = [str(name)] + [str(value) for value in frame[name].head(100) if pd.notna(value)]
        width = min(max((len(value) for value in values), default=12) + 2, 42)
        sheet.column_dimensions[get_column_letter(column_index)].width = max(width, 12)
    if len(frame) and len(frame.columns):
        table_name = re.sub(r"\W", "", sheet.title)[:10] + f"R{start_row}"
        table_name = (table_name or "Data")[:250]
        suffix = 2
        existing_names = {table.displayName for table in sheet.tables.values()}
        base_name = table_name
        while table_name in existing_names:
            table_name = f"{base_name}{suffix}"
            suffix += 1
        table = Table(
            displayName=table_name,
            ref=f"A{start_row}:{get_column_letter(len(frame.columns))}{start_row + len(frame)}",
        )
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        sheet.add_table(table)
    return start_row + len(frame)


def _field_profile(frame: pd.DataFrame, column_types: dict[str, str]) -> pd.DataFrame:
    profiles = []
    for column in frame.columns:
        values = frame[column]
        non_null = values.dropna()
        common = non_null.astype(str).value_counts().head(3)
        profile: dict[str, object] = {
            "Column": column,
            "Detected type": column_types.get(column, "TEXT"),
            "Non-empty rows": int(values.notna().sum()),
            "Missing rows": int(values.isna().sum()),
            "Missing %": round(float(values.isna().mean() * 100), 2) if len(values) else 0,
            "Unique values": int(values.nunique(dropna=True)),
            "Most common values": "; ".join(f"{key} ({count})" for key, count in common.items()),
        }
        if column_types.get(column) in {"MONEY", "NUMBER", "INTEGER", "PERCENT"}:
            numeric = pd.to_numeric(values, errors="coerce").dropna()
            if not numeric.empty:
                profile.update({
                    "Minimum": float(numeric.min()),
                    "Maximum": float(numeric.max()),
                    "Average": float(numeric.mean()),
                    "Median": float(numeric.median()),
                })
        elif column_types.get(column) == "DATE":
            dates = pd.to_datetime(values, errors="coerce").dropna()
            if not dates.empty:
                profile.update({"Earliest": dates.min(), "Latest": dates.max()})
        profiles.append(profile)
    return pd.DataFrame(profiles)


def _title(sheet, title: str, subtitle: str | None = None) -> None:
    sheet["A1"] = title
    sheet["A1"].font = _TITLE_FONT
    if subtitle:
        sheet["A2"] = subtitle
        sheet["A2"].font = _SUBTITLE_FONT
    sheet.sheet_view.showGridLines = False


def _numeric_columns(frame: pd.DataFrame, column_types: dict[str, str]) -> list[str]:
    return [
        column for column in frame.columns
        if column_types.get(column) in {"MONEY", "NUMBER", "INTEGER", "PERCENT"}
        and pd.to_numeric(frame[column], errors="coerce").notna().any()
    ]


def _primary_measure(columns: list[str], column_types: dict[str, str]) -> str | None:
    if not columns:
        return None
    priorities = ("revenue", "sales", "amount", "total", "price", "cost", "profit", "value")
    return max(
        columns,
        key=lambda column: (
            next((len(priorities) - i for i, word in enumerate(priorities) if word in column.casefold()), 0),
            column_types.get(column) == "MONEY",
        ),
    )


def _measure_aggregation(column: str, column_types: dict[str, str]) -> str:
    words = set(re.findall(r"[a-z0-9]+", column.casefold()))
    if column_types.get(column) == "PERCENT" or words.intersection(
        {"unit", "rate", "average", "avg"}
    ):
        return "mean"
    if words.intersection({"price", "cost"}) and not words.intersection(
        {"total", "subtotal", "extended", "revenue", "sales"}
    ):
        return "mean"
    return "sum"


def build_excel_report(
    raw: pd.DataFrame,
    clean: pd.DataFrame,
    column_types: dict[str, str],
    report: dict[str, object],
    output_path: str | Path,
    measure: str | None = None,
    dimensions: list[str] | None = None,
) -> Path:
    """Write all records and broad, data-driven summaries to an .xlsx workbook."""
    if len(clean) + 1 > EXCEL_MAX_ROWS:
        raise ValueError(
            f"The cleaned table has {len(clean):,} records, exceeding Excel's "
            f"{EXCEL_MAX_ROWS - 1:,}-data-row worksheet limit. The input file is unchanged."
        )
    if len(clean.columns) > EXCEL_MAX_COLUMNS:
        raise ValueError(f"The input has more than Excel's {EXCEL_MAX_COLUMNS:,}-column limit.")

    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    overview = workbook.active
    overview.title = "Overview"
    _title(overview, "Data Cleaning & Analysis Report", "Automatically inferred from the uploaded file; review detected types before relying on interpretations.")
    _append_safe_row(overview, [])
    _append_safe_row(overview, ["Input rows", int(report["input_rows"])])
    _append_safe_row(overview, ["Cleaned rows", len(clean)])
    _append_safe_row(overview, ["Input columns", len(raw.columns)])
    _append_safe_row(overview, ["Blank rows removed", int(report["blank_rows_removed"])])
    _append_safe_row(overview, ["Exact duplicate rows removed", int(report["duplicate_rows_removed"])])
    _append_safe_row(overview, ["Numeric/date conversion issues", sum(report["conversion_issues"].values())])
    numeric_columns = _numeric_columns(clean, column_types)
    available_dimensions = [
        column for column in clean.columns
        if column_types.get(column) in {"CATEGORY", "BOOLEAN"} and clean[column].nunique(dropna=True) > 0
    ]
    selected_measure = measure if measure in numeric_columns else _primary_measure(numeric_columns, column_types)
    selected_dimensions = (
        [column for column in dimensions if column in available_dimensions]
        if dimensions is not None
        else available_dimensions
    )
    _append_safe_row(overview, ["Selected chart measure", selected_measure or "Record counts"])
    _append_safe_row(overview, ["Selected chart dimension(s)", ", ".join(selected_dimensions) or "Record counts"])
    _append_safe_row(overview, ["Analysis approach", "Column roles are inferred from names and values. Unknown fields are retained and profiled; no fixed business schema is assumed."])
    _append_safe_row(overview, ["Data quality flags", len(report.get("quality_issues", []))])
    _append_safe_row(overview, ["Repeated header rows removed", int(report.get("repeated_header_rows_removed", 0))])
    _append_safe_row(overview, ["Empty columns removed", ", ".join(report.get("empty_columns_removed", [])) or "None"])
    overview.column_dimensions["A"].width = 34
    overview.column_dimensions["B"].width = 100
    overview.freeze_panes = "A4"
    for row in overview.iter_rows(min_row=4, max_col=2):
        row[0].font = Font(bold=True, color="205B78")
        row[1].alignment = Alignment(wrap_text=True, vertical="top")

    dictionary = workbook.create_sheet("Column Insights")
    _title(dictionary, "Column Insights", "Completeness, uniqueness, representative values, and numeric/date ranges for every field.")
    profiles = _field_profile(clean, column_types)
    _write_frame(dictionary, profiles, start_row=4)

    cleaned_sheet = workbook.create_sheet("Cleaned Data")
    _write_frame(cleaned_sheet, clean, column_types=column_types)

    raw_sheet = workbook.create_sheet("Original Preview")
    _title(raw_sheet, "Original File Preview", "First 1,000 original records for comparison; the full original file is not modified.")
    _write_frame(raw_sheet, raw.head(1000), start_row=4)

    quality_sheet = workbook.create_sheet("Data Quality")
    _title(
        quality_sheet,
        "Data Quality Review",
        "Flags are informational only. Suspicious, negative, or statistical outlier values are retained.",
    )
    quality_fields = ["Issue", "Column", "Count", "Rows (source row numbers)", "Details"]
    quality_rows: list[dict[str, object]] = []
    for issue in report.get("quality_issues", []):
        quality_rows.append({
            "Issue": issue.get("kind", ""),
            "Column": issue.get("column", ""),
            "Count": issue.get("count", 0),
            "Rows (source row numbers)": ", ".join(str(row) for row in issue.get("rows", [])),
            "Details": issue.get("message", ""),
        })
    if not quality_rows:
        quality_rows.append({
            "Issue": "none",
            "Column": "",
            "Count": 0,
            "Rows (source row numbers)": "",
            "Details": "No flagged data-quality issues were found.",
        })
    quality_rows.append({
        "Issue": "category_merges",
        "Column": "",
        "Count": sum(
            len(merge.get("merged", []))
            for merges in report.get("category_merges", {}).values()
            for merge in merges
        ),
        "Rows (source row numbers)": "",
        "Details": "; ".join(
            f"{column}: {merge.get('merged', [])!r} → {merge.get('canonical')!r}"
            for column, merges in report.get("category_merges", {}).items()
            for merge in merges
        ) or "No category labels were merged.",
    })
    _write_frame(quality_sheet, pd.DataFrame(quality_rows, columns=quality_fields), start_row=4)
    if report.get("conversion_failures"):
        start_row = quality_sheet.max_row + 2
        quality_sheet.cell(start_row, 1, "Conversion failure details").font = Font(
            bold=True, size=12, color="205B78"
        )
        _write_frame(
            quality_sheet,
            pd.DataFrame(report["conversion_failures"]),
            start_row=start_row + 1,
        )

    category_columns = available_dimensions
    insights = workbook.create_sheet("Category Insights")
    _title(insights, "Category Insights", "Record counts for every observed value in each detected categorical column.")
    row_cursor = 4
    for column in category_columns:
        counts = clean[column].astype("string").fillna("(Missing)").value_counts(dropna=False)
        if row_cursor + len(counts) + 2 > EXCEL_MAX_ROWS:
            insights.cell(row_cursor, 1, f"Stopped before Excel's {EXCEL_MAX_ROWS:,}-row worksheet limit.")
            break
        insights.cell(row_cursor, 1, f"Field: {column}").font = Font(bold=True, size=12, color="205B78")
        row_cursor += 1
        grouped = counts.rename_axis("Value").rename("Records").reset_index()
        _write_frame(insights, grouped, start_row=row_cursor)
        row_cursor += len(grouped) + 3
    if not category_columns:
        insights["A4"] = "No categorical columns were detected. Review the types on Column Insights."

    measure_insights = workbook.create_sheet("Measure Insights")
    summary_subject = selected_measure or "Record count"
    _title(
        measure_insights,
        "Selected Measure Analysis",
        f"Aggregations for {summary_subject} by the selected category fields.",
    )
    row_cursor = 4
    for dimension in selected_dimensions:
        labels = clean[dimension].astype("string").fillna("(Missing)")
        grouped = labels.value_counts(dropna=False).rename("Records").to_frame()
        if selected_measure:
            values = pd.to_numeric(clean[selected_measure], errors="coerce")
            joined = pd.DataFrame({"label": labels, "measure": values})
            aggregation = _measure_aggregation(selected_measure, column_types)
            grouped[selected_measure] = joined.groupby("label")["measure"].agg(
                "mean" if aggregation == "mean" else lambda items: items.sum(min_count=1)
            )
        grouped.index.name = dimension
        grouped = grouped.reset_index().sort_values(
            selected_measure if selected_measure else "Records",
            ascending=False,
            na_position="last",
        )
        if row_cursor + len(grouped) + 2 > EXCEL_MAX_ROWS:
            measure_insights.cell(
                row_cursor,
                1,
                f"Stopped before Excel's {EXCEL_MAX_ROWS:,}-row worksheet limit; full records remain in Cleaned Data.",
            )
            break
        measure_insights.cell(row_cursor, 1, f"Breakdown: {dimension}").font = Font(
            bold=True, size=12, color="205B78"
        )
        row_cursor += 1
        _write_frame(measure_insights, grouped, start_row=row_cursor)
        row_cursor += len(grouped) + 3
    if not selected_dimensions:
        measure_insights["A4"] = "Choose a detected category field to see grouped analysis; all numeric summaries are in Column Insights."

    chart_sheet = workbook.create_sheet("Charts")
    _title(chart_sheet, "Visual Analysis", "Charts are selected automatically from detected columns; summaries and cleaned records remain available in the other sheets.")
    chart_data = workbook.create_sheet("Chart Data")
    chart_data.sheet_state = "hidden"
    chart_cursor = 1
    charts_added = 0
    measure = selected_measure
    dimensions = selected_dimensions
    date_columns = [column for column in clean.columns if column_types.get(column) == "DATE"]

    def add_bar_chart(title: str, labels: pd.Series, values: pd.Series) -> None:
        nonlocal chart_cursor, charts_added
        if charts_added >= MAX_CHARTS or labels.empty:
            return
        start = chart_cursor
        chart_data.cell(start, 1, "Value")
        chart_data.cell(start, 2, "Measure")
        for offset, (label, value) in enumerate(zip(labels, values), start=1):
            _set_excel_cell(chart_data.cell(start + offset, 1), str(label))
            _set_excel_cell(chart_data.cell(start + offset, 2), value)
        chart = BarChart()
        chart.type = "bar"
        chart.style = 10
        chart.title = title
        chart.y_axis.title = "Group or range"
        chart.x_axis.title = "Records or total"
        chart.add_data(Reference(chart_data, min_col=2, min_row=start, max_row=start + len(labels)), titles_from_data=True)
        chart.set_categories(Reference(chart_data, min_col=1, min_row=start + 1, max_row=start + len(labels)))
        chart.height = 8
        chart.width = 15
        chart.legend = None
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        chart_sheet.add_chart(chart, f"A{3 + charts_added * 18}")
        chart_cursor += len(labels) + 3
        charts_added += 1

    def add_line_chart(title: str, labels: pd.Series, values: pd.Series) -> None:
        nonlocal chart_cursor, charts_added
        if charts_added >= MAX_CHARTS or labels.empty:
            return
        start = chart_cursor
        chart_data.cell(start, 1, "Period")
        chart_data.cell(start, 2, "Measure")
        for offset, (label, value) in enumerate(zip(labels, values), start=1):
            _set_excel_cell(chart_data.cell(start + offset, 1), str(label))
            _set_excel_cell(chart_data.cell(start + offset, 2), value)
        chart = LineChart()
        chart.style = 13
        chart.title = title
        chart.y_axis.title = "Value"
        chart.x_axis.title = "Period"
        chart.add_data(Reference(chart_data, min_col=2, min_row=start, max_row=start + len(labels)), titles_from_data=True)
        chart.set_categories(Reference(chart_data, min_col=1, min_row=start + 1, max_row=start + len(labels)))
        chart.height = 8
        chart.width = 15
        chart.legend = None
        chart_sheet.add_chart(chart, f"A{3 + charts_added * 18}")
        chart_cursor += len(labels) + 3
        charts_added += 1

    if measure:
        for dimension in dimensions:
            values = clean[dimension].astype("string").fillna("(Missing)")
            numeric = pd.to_numeric(clean[measure], errors="coerce")
            aggregation = _measure_aggregation(measure, column_types)
            if aggregation == "mean":
                grouped = numeric.groupby(values).mean().dropna()
            else:
                grouped = numeric.groupby(values).sum(min_count=1).dropna()
            grouped = grouped.sort_values(ascending=False).head(MAX_CATEGORY_CHART_VALUES)
            if len(grouped) >= 2:
                aggregation_label = "Average" if aggregation == "mean" else "Sum"
                add_bar_chart(f"{aggregation_label} of {measure} by {dimension}", grouped.index.to_series(), grouped)
    for dimension in dimensions:
        if charts_added >= MAX_CHARTS:
            break
        if measure and dimensions:
            continue
        counts = clean[dimension].astype("string").fillna("(Missing)").value_counts().head(MAX_CATEGORY_CHART_VALUES)
        if len(counts) >= 2:
            add_bar_chart(f"Records by {dimension}", counts.index.to_series(), counts)
    for column in date_columns:
        if charts_added >= MAX_CHARTS:
            break
        dates = pd.to_datetime(clean[column], errors="coerce").dropna()
        if dates.nunique() < 2:
            continue
        periods = dates.dt.to_period("M")
        if measure:
            data = pd.DataFrame({"period": periods, "value": pd.to_numeric(clean.loc[dates.index, measure], errors="coerce")})
            aggregation = _measure_aggregation(measure, column_types)
            if aggregation == "mean":
                grouped = data.groupby("period")["value"].mean().dropna()
                title = f"Average {measure} over {column}"
            else:
                grouped = data.groupby("period")["value"].sum(min_count=1).dropna()
                title = f"Sum of {measure} over {column}"
        else:
            grouped = periods.value_counts().sort_index()
            title = f"Records over {column}"
        if len(grouped) >= 2:
            add_line_chart(title, grouped.index.to_series(), grouped)

    for column in numeric_columns:
        if charts_added >= MAX_CHARTS:
            break
        values = pd.to_numeric(clean[column], errors="coerce").dropna()
        if values.nunique() < 2:
            continue
        buckets = pd.cut(values, bins=min(10, int(values.nunique())), duplicates="drop", include_lowest=True)
        distribution = buckets.value_counts().sort_index()
        add_bar_chart(
            f"Distribution of {column}",
            distribution.index.to_series().astype(str),
            distribution,
        )

    if not charts_added:
        chart_sheet["A4"] = "No chartable patterns were detected. Review Column Insights and Category Insights for available analysis."

    for sheet in workbook.worksheets:
        sheet.sheet_view.showGridLines = False
    workbook.active = 0
    workbook.save(destination)
    return destination
