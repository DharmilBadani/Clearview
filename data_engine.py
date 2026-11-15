"""Schema-agnostic loading, type detection, and cleaning for tabular files."""

from __future__ import annotations

import csv
import re
import warnings
from pathlib import Path
from typing import Callable

import pandas as pd

TYPE_CHOICES = (
    "DATE",
    "BOOLEAN",
    "MONEY",
    "PERCENT",
    "INTEGER",
    "NUMBER",
    "CATEGORY",
    "IDENTIFIER",
    "TEXT",
)

_NULL_VALUES = {
    "", "null", "none", "nan", "n/a", "n.a.", "na", "?", "--", "-",
    "#n/a", "#div/0!", "#value!", "#ref!",
}
_MISSING_CONVERSION_MARKERS = {"", "?", "tbd", "unknown", "00/00/0000", "na", "n/a", "n.a.", "none", "null", "nan", "-"}
_NUMBER_WORDS = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
    "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14,
    "fifteen": 15, "sixteen": 16, "seventeen": 17, "eighteen": 18,
    "nineteen": 19, "twenty": 20, "thirty": 30, "forty": 40,
    "fifty": 50, "sixty": 60, "seventy": 70, "eighty": 80, "ninety": 90,
}
_ID_WORDS = {"id", "uuid", "sku", "code", "phone", "telephone", "postcode", "zip"}
_MONEY_WORDS = {
    "amount", "cost", "price", "revenue", "sales", "total", "subtotal",
    "fee", "salary", "income", "spend", "balance", "profit",
}
_COUNT_WORDS = {"age", "count", "inventory", "qty", "quantity", "satisfaction", "stock", "unit", "units"}
_PERCENT_WORDS = {"percent", "percentage", "discount", "margin", "rate", "ratio"}
_BOOLEAN_VALUES = {
    "y": "Yes",
    "yes": "Yes",
    "true": "Yes",
    "1": "Yes",
    "t": "Yes",
    "n": "No",
    "no": "No",
    "false": "No",
    "0": "No",
    "f": "No",
}

_CURRENCY_PATTERN = r"[$\u20ac\u00a3\u00a5\u20b9\u20bd]"


def _clean_cell_text(value: object) -> object:
    if not isinstance(value, str):
        return value
    value = (
        value.replace("\ufeff", "")
        .replace("\u200b", "")
        .replace("\u200c", "")
        .replace("\u200d", "")
        .replace("\u00a0", " ")
        .strip()
    )
    return pd.NA if value.casefold() in _NULL_VALUES else value


def _unique_headers(columns: object) -> list[str]:
    headers: list[str] = []
    seen_headers: dict[str, int] = {}
    used_headers: set[str] = set()
    for index, column in enumerate(columns):
        base = str(column).strip() or f"Column_{index + 1}"
        occurrence = seen_headers.get(base.casefold(), 0) + 1
        seen_headers[base.casefold()] = occurrence
        candidate = base if occurrence == 1 else f"{base}_{occurrence}"
        while candidate.casefold() in used_headers:
            occurrence += 1
            candidate = f"{base}_{occurrence}"
        seen_headers[base.casefold()] = occurrence
        used_headers.add(candidate.casefold())
        headers.append(candidate)
    return headers


def _header_row_index(rows: list[list[object]]) -> int:
    """Choose a real header after banner/title rows in an Excel sheet."""
    for index, row in enumerate(rows[:100]):
        values = [_clean_cell_text(value) for value in row]
        non_empty = [str(value).strip() for value in values if not pd.isna(value) and str(value).strip()]
        if len(values) < 2 or len(non_empty) / max(1, len(values)) < 0.6:
            continue
        unique = {value.casefold() for value in non_empty}
        mostly_text = sum(not re.fullmatch(r"[-+]?\d+(?:\.\d+)?", value) for value in non_empty) / len(non_empty)
        if len(unique) / len(non_empty) >= 0.6 and mostly_text >= 0.6:
            return index
    return 0


def _load_excel_rows(file_path: Path, include_hidden_sheets: bool) -> tuple[pd.DataFrame, list[str], dict[str, object]]:
    try:
        from openpyxl import load_workbook
    except ImportError:
        with pd.ExcelFile(file_path) as workbook:
            sheets = list(workbook.sheet_names)
        parsed = pd.read_excel(file_path, header=None, dtype=str, keep_default_na=False)
        return parsed, sheets, {"excel_header_row": 1}

    workbook = load_workbook(file_path, read_only=False, data_only=True)
    all_sheets = [sheet.title for sheet in workbook.worksheets]
    visible_sheets = []
    for sheet in workbook.worksheets:
        if not include_hidden_sheets and sheet.sheet_state != "visible":
            continue
        if any(value is not None and str(value).strip() for row in sheet.iter_rows(values_only=True) for value in row):
            visible_sheets.append(sheet.title)
    selected = visible_sheets or all_sheets
    selected_sheet = selected[0] if selected else None
    if selected_sheet is None:
        raise ValueError("The Excel workbook does not contain any worksheets.")
    worksheet = workbook[selected_sheet]
    rows = [list(row) for row in worksheet.iter_rows(values_only=True)]
    return pd.DataFrame(rows), visible_sheets, {"excel_header_row": _header_row_index(rows) + 1}


def load_table(
    path: str | Path,
    sheet_name: str | None = None,
    include_hidden_sheets: bool = False,
) -> tuple[pd.DataFrame, list[str]]:
    """Load one selected worksheet or a delimited text file as strings."""
    file_path = Path(path)
    suffix = file_path.suffix.casefold()

    load_metadata: dict[str, object] = {}
    if suffix in {".xlsx", ".xlsm"}:
        workbook_frame, sheets, load_metadata = _load_excel_rows(file_path, include_hidden_sheets)
        selected_sheet = sheet_name if sheet_name in sheets else (sheets[0] if sheets else None)
        if selected_sheet is None:
            raise ValueError("The Excel workbook does not contain any visible worksheets with data.")
        # Re-read the requested sheet while retaining cached formula values.
        from openpyxl import load_workbook
        workbook = load_workbook(file_path, read_only=False, data_only=True)
        rows = [list(row) for row in workbook[selected_sheet].iter_rows(values_only=True)]
        header_index = _header_row_index(rows)
        parsed_file = pd.DataFrame(rows)
        if parsed_file.empty:
            raise ValueError("The selected file does not contain a header row.")
        headers = parsed_file.iloc[header_index].tolist()
        frame = parsed_file.iloc[header_index + 1:].copy()
        frame.columns = headers
        load_metadata["excel_header_row"] = header_index + 1
        load_metadata["excel_sheet"] = selected_sheet
        load_metadata["hidden_sheets_ignored"] = not include_hidden_sheets
    elif suffix == ".xls":
        with pd.ExcelFile(file_path) as workbook:
            sheets = list(workbook.sheet_names)
            if not sheets:
                raise ValueError("The Excel workbook does not contain any worksheets.")
            selected_sheet = sheet_name if sheet_name in sheets else sheets[0]
            parsed_file = pd.read_excel(workbook, sheet_name=selected_sheet, header=None, dtype=str, keep_default_na=False)
            if parsed_file.empty:
                raise ValueError("The selected file does not contain a header row.")
            header_index = _header_row_index(parsed_file.values.tolist())
            headers = parsed_file.iloc[header_index].tolist()
            frame = parsed_file.iloc[header_index + 1:].copy()
            frame.columns = headers
    elif suffix in {".csv", ".tsv", ".tab", ".txt"}:
        sheets = []
        with file_path.open("rb") as source:
            sample = source.read(65536)
        decoded = None
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                decoded = sample.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        if decoded is None:
            raise UnicodeError("The file encoding could not be detected.")

        delimiter = "\t" if suffix in {".tsv", ".tab"} else None
        if delimiter is None:
            try:
                delimiter = csv.Sniffer().sniff(decoded, delimiters=",;\t|").delimiter
            except csv.Error:
                delimiter = ","

        last_error: UnicodeDecodeError | None = None
        frame = None
        ragged_rows: list[dict[str, object]] = []
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                with file_path.open("r", encoding=encoding, newline="") as source:
                    rows = list(csv.reader(source, delimiter=delimiter))
                if not rows:
                    raise ValueError("The selected file does not contain a header row.")
                expected_fields = len(rows[0])
                normalized_rows: list[list[str]] = [rows[0]]
                overflow_column = "__overflow__"
                for source_row, row in enumerate(rows[1:], start=2):
                    if len(row) != expected_fields:
                        extras = row[expected_fields:]
                        normalized = row[:expected_fields] if extras else row[:expected_fields] + [""] * (expected_fields - len(row))
                        ragged_rows.append({"row": source_row, "expected": expected_fields, "actual": len(row), "overflow": extras})
                        normalized_rows.append(normalized)
                    else:
                        normalized_rows.append(row)
                if any(item["overflow"] for item in ragged_rows):
                    normalized_rows[0] = normalized_rows[0] + [overflow_column]
                    for index, item in enumerate(ragged_rows, start=1):
                        normalized_rows[index] = normalized_rows[index] + [" | ".join(item["overflow"])]
                    for index, row in enumerate(normalized_rows[1:], start=1):
                        if len(row) == expected_fields:
                            normalized_rows[index] = row + [""]
                parsed_file = pd.DataFrame(normalized_rows[1:], columns=normalized_rows[0])
                headers = parsed_file.columns.tolist()
                frame = parsed_file.copy()
                load_metadata["ragged_rows_fixed"] = ragged_rows
                frame.columns = headers
                break
            except UnicodeDecodeError as error:
                last_error = error
        if frame is None:
            raise UnicodeError("The file encoding could not be detected.") from last_error
    else:
        raise ValueError("Choose a CSV, TSV, TXT, XLS, XLSX, or XLSM file.")

    if frame.empty and len(frame.columns) == 0:
        raise ValueError("The selected file does not contain a header row.")
    if len(frame.columns) == 0:
        raise ValueError("The selected sheet does not contain any columns.")

    frame.columns = _unique_headers(frame.columns)
    frame.attrs["load_report"] = load_metadata
    return frame, sheets


def _column_words(column: object) -> set[str]:
    return set(re.findall(r"[a-z0-9]+", str(column).casefold()))


def _repeated_header_mask(frame: pd.DataFrame) -> pd.Series:
    if not len(frame.columns):
        return pd.Series(False, index=frame.index, dtype=bool)
    matches = [
        frame.iloc[:, index]
        .map(_clean_cell_text)
        .astype("string")
        .str.strip()
        .str.casefold()
        .eq(str(column).strip().casefold())
        .fillna(False)
        for index, column in enumerate(frame.columns)
    ]
    return pd.concat(matches, axis=1).sum(axis=1).ge((len(frame.columns) + 1) // 2)


def _duplicate_token(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\u00a0", " ")).strip().casefold()


def _structural_junk_mask(frame: pd.DataFrame) -> pd.Series:
    if frame.empty or not len(frame.columns):
        return pd.Series(False, index=frame.index, dtype=bool)
    first = frame.iloc[:, 0].map(_duplicate_token)
    return first.str.startswith("#") | first.str.fullmatch(r"(?:total|subtotal|grand total)", na=False)


def _parse_number_legacy(value: object) -> float | None:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text or text.casefold() in _NULL_VALUES:
        return None

    negative = text.startswith("(") and text.endswith(")")
    text = text.strip().strip("()").replace("−", "-")
    text = re.sub(
        r"(?i)\b(?:usd|eur|gbp|cad|aud|inr|jpy|units?|pcs?|pieces?|items?|kgs?|kg|g|lbs?|lb|hours?|hrs?)\b",
        "",
        text,
    )
    multiplier = 1
    suffix = re.search(r"(?i)([kml])\s*$", text)
    if suffix:
        multiplier = {
            "k": 1_000,
            "m": 1_000_000,
            "l": 100_000,
        }[suffix.group(1).casefold()]
        text = text[:suffix.start()]
    text = re.sub(r"[₹$€£¥]", "", text)
    if re.search(r"[A-Za-z]", text):
        return None
    text = re.sub(r"[^\d,.\-+]", "", text)
    if not text or text in {"-", "+", ".", ","}:
        return None

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        if re.fullmatch(r"[+-]?\d{1,3}(?:,\d{3})+", text):
            text = text.replace(",", "")
        elif text.count(",") == 1:
            text = text.replace(",", ".")
        else:
            return None

    try:
        number = float(text)
    except ValueError:
        return None
    if negative:
        number = -abs(number)
    return number * multiplier


# Re-define the parser with an explicit currency pattern. This keeps symbols
# and Unicode minus signs from being confused with numeric content.
def _parse_number(value: object) -> float | None:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text or text.casefold() in _NULL_VALUES:
        return None
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip().strip("()").replace("\u2212", "-").replace("\u2013", "-").replace("\u2014", "-")
    text = re.sub(
        r"(?i)\b(?:usd|eur|gbp|cad|aud|inr|jpy|units?|pcs?|pieces?|items?|kgs?|kg|g|lbs?|lb|hours?|hrs?)\b",
        "",
        text,
    )
    multiplier = 1
    suffix = re.search(r"(?i)([kml])\s*$", text)
    if suffix:
        multiplier = {"k": 1_000, "m": 1_000_000, "l": 100_000}[suffix.group(1).casefold()]
        text = text[:suffix.start()]
    text = re.sub(_CURRENCY_PATTERN, "", text)
    scientific = re.fullmatch(r"[+-]?\d+(?:\.\d+)?[eE][+-]?\d+", text.strip())
    if scientific:
        try:
            number = float(text)
        except ValueError:
            return None
        return (-abs(number) if negative else number) * multiplier
    word_number = _NUMBER_WORDS.get(text.casefold())
    if word_number is not None:
        return (-abs(word_number) if negative else word_number) * multiplier
    if re.search(r"[A-Za-z]", text):
        return None
    text = re.sub(r"[^\d,.\-+]", "", text)
    if not text or text in {"-", "+", ".", ","}:
        return None
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".") if text.rfind(",") > text.rfind(".") else text.replace(",", "")
    elif "," in text:
        if re.fullmatch(r"[+-]?\d{1,3}(?:,\d{3})+", text):
            text = text.replace(",", "")
        elif text.count(",") == 1:
            text = text.replace(",", ".")
        else:
            return None
    try:
        number = float(text)
    except ValueError:
        return None
    return (-abs(number) if negative else number) * multiplier


def _numeric_values(series: pd.Series) -> pd.Series:
    return series.map(_parse_number).astype("float64")


def _date_values(series: pd.Series, day_first: bool | None = None) -> pd.Series:
    text = series.astype("string").str.strip()
    present = text.notna() & ~text.str.casefold().isin(_NULL_VALUES)
    inferred_day_first, _ = date_order_info(text[present])
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", FutureWarning)
        parsed = pd.to_datetime(text, format="mixed", errors="coerce", dayfirst=inferred_day_first if day_first is None else day_first, utc=True)
    numeric = pd.to_numeric(text, errors="coerce")
    excel_mask = numeric.between(1, 100000) & numeric.notna()
    unix_ms_mask = numeric.between(100000000000, 10000000000000) & numeric.notna()
    unix_s_mask = numeric.between(1000000000, 100000000000) & numeric.notna()
    if excel_mask.any():
        parsed.loc[excel_mask] = pd.to_datetime(
            numeric.loc[excel_mask], unit="D", origin="1899-12-30", errors="coerce", utc=True
        )
    if unix_ms_mask.any():
        parsed.loc[unix_ms_mask] = pd.to_datetime(numeric.loc[unix_ms_mask], unit="ms", errors="coerce", utc=True)
    if unix_s_mask.any():
        parsed.loc[unix_s_mask] = pd.to_datetime(numeric.loc[unix_s_mask], unit="s", errors="coerce", utc=True)
    return parsed.dt.tz_convert("UTC").dt.tz_localize(None)


def date_order_info(series: pd.Series) -> tuple[bool, bool]:
    """Infer day-first only from decisive evidence and flag ambiguous slash dates."""
    day_first_evidence = False
    month_first_evidence = False
    ambiguous_seen = False
    slash_date = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})(?:\s.*)?$")
    for raw in series.dropna().astype(str):
        match = slash_date.match(raw.strip())
        if not match:
            continue
        first, second = int(match.group(1)), int(match.group(2))
        if first > 12 and second <= 12:
            day_first_evidence = True
        elif second > 12 and first <= 12:
            month_first_evidence = True
        elif 1 <= first <= 12 and 1 <= second <= 12 and first != second:
            ambiguous_seen = True
    if day_first_evidence and not month_first_evidence:
        return True, False
    if month_first_evidence and not day_first_evidence:
        return False, False
    if day_first_evidence and month_first_evidence:
        return False, True
    return False, ambiguous_seen


def _detect_column_types_unsafe(frame: pd.DataFrame) -> dict[str, str]:
    """Infer a type for every column from its label and observed values."""
    result: dict[str, str] = {}
    for column in frame.columns:
        series = frame[column].map(_clean_cell_text).dropna()
        words = _column_words(column)
        if series.empty:
            result[column] = "TEXT"
            continue

        values = series.astype(str).str.strip()
        lowered = values.str.casefold()
        if lowered.isin(_BOOLEAN_VALUES).mean() >= 0.9:
            result[column] = "BOOLEAN"
            continue

        id_hint = bool(words & _ID_WORDS) or any(
            word.endswith("_id") for word in str(column).casefold().split()
        )
        if id_hint:
            result[column] = "IDENTIFIER"
            continue

        try:
            parsed_dates = _date_values(values)
        except Exception:
            parsed_dates = pd.Series(pd.NaT, index=values.index)
        date_ratio = parsed_dates.notna().mean()
        date_hint = bool(words & {"date", "time", "dob", "joined", "created", "updated", "signed", "at"})
        non_numeric_dates = parsed_dates.notna() & pd.to_numeric(values, errors="coerce").isna()
        if date_ratio >= 0.5 and (date_hint or non_numeric_dates.mean() >= 0.5):
            result[column] = "DATE"
            continue

        try:
            numeric = _numeric_values(values)
        except Exception:
            numeric = pd.Series(float("nan"), index=values.index, dtype="float64")
        numeric_ratio = numeric.notna().mean()
        has_percent = values.str.contains("%", regex=False).mean() >= 0.2
        if bool(words & _PERCENT_WORDS) or has_percent:
            result[column] = "PERCENT"
            continue
        currency_ratio = values.str.contains(
            r"[₹$€£¥]|(?i:\b(?:usd|eur|gbp|cad|aud|inr|jpy)\b)", regex=True
        ).mean()
        currency_ratio = max(
            currency_ratio,
            float(values.str.contains(_CURRENCY_PATTERN, regex=True).mean()),
        )
        suffix_ratio = values.str.contains(r"(?i)\d(?:[.,]\d+)?\s*[kml]\s*$", regex=True).mean()
        if numeric_ratio >= 0.6 and (
            currency_ratio >= 0.6
            or suffix_ratio >= 0.6
            or bool(words & _MONEY_WORDS)
        ):
            result[column] = "MONEY"
            continue
        if numeric_ratio >= 0.75:
            leading_zero_values = values.str.fullmatch(r"0\d+").fillna(False)
            if leading_zero_values.mean() >= 0.5:
                result[column] = "IDENTIFIER"
                continue
            if bool(words & {"qty", "quantity", "units", "count", "items"}):
                result[column] = "INTEGER" if (numeric.dropna() % 1 == 0).all() else "NUMBER"
            elif (numeric.dropna() % 1 == 0).all():
                result[column] = "INTEGER"
            else:
                result[column] = "NUMBER"
            continue

        unique_ratio = values.nunique() / max(1, len(values))
        if bool(words & {
            "category", "region", "channel", "status", "method", "country", "city",
            "state", "type", "department", "group", "product", "segment", "team",
        }) or (values.nunique() <= 50 and unique_ratio <= 0.25):
            result[column] = "CATEGORY"
        else:
            result[column] = "TEXT"
    return result


def detect_column_types(frame: pd.DataFrame) -> dict[str, str]:
    """Infer each column independently so one malformed column cannot abort detection."""
    result: dict[str, str] = {}
    errors: list[dict[str, object]] = []
    for column in frame.columns:
        try:
            result[column] = _detect_column_types_unsafe(frame[[column]])[column]
        except Exception as error:
            result[column] = "TEXT"
            errors.append({"column": str(column), "error": str(error)})
    frame.attrs["detection_errors"] = errors
    return result


def clean_table(
    raw: pd.DataFrame,
    column_types: dict[str, str] | None = None,
    log: Callable[[str], None] | None = None,
    normalize_categories: bool = True,
    date_orders: dict[str, str] | None = None,
    normalize_booleans: bool = True,
    validation_rules: list[dict[str, str]] | None = None,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """Normalize strings and detected types without dropping partially populated rows."""

    def emit(message: str) -> None:
        if log is not None:
            log(message)

    frame = raw.copy()
    original_nonempty_columns = {
        column: bool(frame[column].map(lambda value: value is not None and not (isinstance(value, float) and pd.isna(value)) and str(value).strip()).any())
        for column in frame.columns
    }
    for column in frame.columns:
        frame[column] = frame[column].map(_clean_cell_text)
    load_report = dict(frame.attrs.get("load_report", {}))
    source_start = int(load_report.get("excel_header_row", 1)) + 1 if load_report.get("excel_header_row") else 2
    source_rows = pd.Series(range(source_start, source_start + len(frame)), index=frame.index)

    repeated_header_mask = _repeated_header_mask(frame)
    repeated_header_rows_removed = int(repeated_header_mask.sum())
    repeated_header_source_rows = [
        int(source_rows.loc[index])
        for index in repeated_header_mask[repeated_header_mask].index
    ]
    frame = frame.loc[~repeated_header_mask].copy()
    source_rows = source_rows.loc[frame.index]

    structural_mask = _structural_junk_mask(frame)
    structural_rows_removed = [int(source_rows.loc[index]) for index in structural_mask[structural_mask].index]
    frame = frame.loc[~structural_mask].copy()
    source_rows = source_rows.loc[frame.index]

    empty_columns_removed = [
        column for column in frame.columns
        if not frame[column].notna().any() and not original_nonempty_columns.get(column, False)
    ]
    if empty_columns_removed:
        frame.drop(columns=empty_columns_removed, inplace=True)

    blank_mask = frame.isna().all(axis=1) if len(frame.columns) else pd.Series(True, index=frame.index)
    blank_rows = int(blank_mask.sum())
    frame = frame.loc[~blank_mask].copy()
    source_rows = source_rows.loc[frame.index]
    before_duplicates = len(frame)
    normalized_frame = frame.apply(lambda column: column.map(_duplicate_token))
    unique_mask = ~normalized_frame.duplicated(keep="first")
    id_columns = [column for column in frame.columns if "id" in str(column).casefold() or str(column).casefold() in {"code", "order no", "order number"}]
    conflicting_duplicate_rows: list[int] = []
    if id_columns:
        id_column = id_columns[0]
        for _, group in normalized_frame.groupby(id_column, dropna=False, sort=False):
            if not group[id_column].iloc[0] or len(group) < 2:
                continue
            distinct_payloads = group.drop(columns=[id_column]).drop_duplicates()
            if len(distinct_payloads) > 1:
                conflicting_duplicate_rows.extend(int(source_rows.loc[index]) for index in group.index)
    frame = frame.loc[unique_mask].copy()
    source_rows = source_rows.loc[frame.index]
    duplicate_rows = before_duplicates - len(frame)
    frame.reset_index(drop=True, inplace=True)
    source_rows.reset_index(drop=True, inplace=True)
    if frame.empty:
        raise ValueError("The selected file contains no non-empty records.")
    source_values = frame.copy()

    detected_types = column_types or detect_column_types(frame)
    types = {column: detected_types.get(column, "TEXT") for column in frame.columns}
    detection_errors = list(frame.attrs.get("detection_errors", []))
    issues: dict[str, int] = {}
    conversion_failures: list[dict[str, object]] = []
    category_merges: dict[str, list[dict[str, object]]] = {}
    ambiguous_dates: dict[str, bool] = {}
    column_errors: list[dict[str, object]] = []
    date_timezone_used = False
    date_orders = date_orders or {}
    for column in frame.columns:
        kind = types.get(column, "TEXT")
        source = frame[column]
        if kind in {"MONEY", "NUMBER", "INTEGER", "PERCENT"}:
            try:
                parsed = _numeric_values(source)
            except Exception as error:
                types[column] = "TEXT"
                frame[column] = source.astype("string").str.strip()
                column_errors.append({"column": str(column), "error": str(error)})
                emit(f"Fell back to TEXT for {column}: {error}")
                continue
            source_present = source.notna() & ~source.astype("string").str.strip().str.casefold().isin(_MISSING_CONVERSION_MARKERS)
            invalid_mask = source_present & parsed.isna()
            if kind == "PERCENT":
                has_percent = source.astype("string").str.strip().str.endswith("%", na=False)
                parsed = parsed.where(~has_percent, parsed / 100)
                bare_percentage = ~has_percent & parsed.abs().gt(1)
                parsed = parsed.where(~bare_percentage, parsed / 100)
            if kind == "INTEGER":
                non_integer_mask = source_present & parsed.notna() & parsed.mod(1).ne(0)
                invalid_mask = invalid_mask | non_integer_mask
                parsed = parsed.mask(non_integer_mask)
                frame[column] = parsed.astype("Int64")
            else:
                frame[column] = parsed
            invalid = int(invalid_mask.sum())
            if invalid:
                issues[column] = invalid
                for position in invalid_mask[invalid_mask].index:
                    conversion_failures.append({
                        "column": column,
                        "value": str(source.loc[position]),
                        "row": int(source_rows.loc[position]),
                        "reason": (
                            "Not a valid integer"
                            if kind == "INTEGER" and non_integer_mask.loc[position]
                            else "Could not convert value"
                        ),
                    })
        elif kind == "DATE":
            requested_order = date_orders.get(column)
            try:
                inferred_day_first, ambiguous = date_order_info(source)
                day_first = requested_order == "day_first" if requested_order in {"day_first", "month_first"} else inferred_day_first
                text = source.astype("string").str.strip()
                date_timezone_used = date_timezone_used or bool(text.str.contains(r"(?:Z|[+-]\d{2}:?\d{2})$", regex=True, na=False).any())
                numeric_dates = pd.to_numeric(text, errors="coerce")
                date_timezone_used = date_timezone_used or bool(numeric_dates.between(1, 100000).any() or numeric_dates.between(1000000000, 10000000000000).any())
                parsed = _date_values(text, day_first=day_first)
            except Exception as error:
                types[column] = "TEXT"
                frame[column] = source.astype("string").str.strip()
                column_errors.append({"column": str(column), "error": str(error)})
                emit(f"Fell back to TEXT for {column}: {error}")
                continue
            invalid_mask = source.notna() & ~text.str.casefold().isin(_MISSING_CONVERSION_MARKERS) & parsed.isna()
            invalid = int(invalid_mask.sum())
            frame[column] = parsed
            ambiguous_dates[column] = ambiguous
            if invalid:
                issues[column] = invalid
                for position in invalid_mask[invalid_mask].index:
                    conversion_failures.append({
                        "column": column,
                        "value": str(source.loc[position]),
                        "row": int(source_rows.loc[position]),
                        "reason": "Could not parse date",
                    })
        elif kind == "BOOLEAN":
            normalized = source.astype("string").str.strip()
            if normalize_booleans:
                parsed = normalized.str.casefold().map(_BOOLEAN_VALUES)
                invalid_mask = source.notna() & parsed.isna()
                frame[column] = parsed
                invalid = int(invalid_mask.sum())
                if invalid:
                    issues[column] = invalid
                    for position in invalid_mask[invalid_mask].index:
                        conversion_failures.append({
                            "column": column,
                            "value": str(source.loc[position]),
                            "row": int(source_rows.loc[position]),
                            "reason": "Not a recognized YES/NO value",
                        })
            else:
                frame[column] = normalized
        elif kind == "CATEGORY":
            normalized = source.astype("string").str.strip().str.replace(r"\s+", " ", regex=True)
            if normalize_categories:
                merges: list[dict[str, object]] = []
                non_null = normalized.dropna()
                for _, members in non_null.groupby(non_null.str.casefold(), sort=False):
                    original_spellings = source.loc[members.index].astype("string").str.strip()
                    spellings = original_spellings.value_counts(sort=False)
                    if len(spellings) < 2:
                        continue
                    first_seen = {value: index for index, value in enumerate(spellings.index)}
                    winning_spelling = max(
                        spellings.index,
                        key=lambda value: (
                            int(spellings[value]),
                            int(value == value.title()),
                            -first_seen[value],
                        ),
                    )
                    canonical = " ".join(winning_spelling.split())
                    if winning_spelling.islower() or winning_spelling.isupper():
                        canonical = canonical.title()
                    variants = [value for value in spellings.index if value != canonical]
                    merged_row_count = int(sum(spellings[value] for value in variants))
                    normalized.loc[members.index] = canonical
                    merges.append({
                        "canonical": canonical,
                        "merged": variants,
                        "count": merged_row_count,
                    })
                if merges:
                    category_merges[column] = merges
            frame[column] = normalized
        else:
            frame[column] = source.astype("string").str.strip().str.replace(r"\s+", " ", regex=True)
        emit(f"Prepared {column} as {kind}")

    quality_issues: list[dict[str, object]] = []
    for error in detection_errors:
        quality_issues.append({
            "kind": "column_error",
            "column": error["column"],
            "count": 1,
            "rows": [],
            "message": f"Column {error['column']} fell back to TEXT during detection: {error['error']}",
        })
    for error in column_errors:
        quality_issues.append({
            "kind": "column_error",
            "column": error["column"],
            "count": 1,
            "rows": [],
            "message": f"Column {error['column']} fell back to TEXT during cleaning: {error['error']}",
        })
    ragged_rows = load_report.get("ragged_rows_fixed", [])
    if ragged_rows:
        quality_issues.append({
            "kind": "ragged_rows_fixed",
            "column": "",
            "count": len(ragged_rows),
            "rows": [int(item["row"]) for item in ragged_rows[:10]],
            "message": f"Fixed {len(ragged_rows):,} ragged CSV row(s) by padding blanks and preserving overflow values",
        })
    if structural_rows_removed:
        quality_issues.append({
            "kind": "structural_rows_removed",
            "column": "",
            "count": len(structural_rows_removed),
            "rows": structural_rows_removed[:10],
            "message": f"Removed {len(structural_rows_removed):,} structural total/comment row(s)",
        })
    if conflicting_duplicate_rows:
        quality_issues.append({
            "kind": "conflicting_duplicate",
            "column": id_columns[0] if id_columns else "",
            "count": len(conflicting_duplicate_rows),
            "rows": conflicting_duplicate_rows[:10],
            "message": f"Found {len(conflicting_duplicate_rows):,} rows sharing an ID with different data; all rows were kept",
        })
    if load_report.get("excel_header_row"):
        quality_issues.append({
            "kind": "excel_header_detected",
            "column": "",
            "count": 1,
            "rows": [int(load_report["excel_header_row"])],
            "message": f"Detected the Excel header on source row {load_report['excel_header_row']}",
        })
    if date_timezone_used:
        quality_issues.append({
            "kind": "date_timezone_normalized",
            "column": "",
            "count": 1,
            "rows": [],
            "message": "Parsed date/time values are normalized to UTC and exported without a timezone offset",
        })
    for rule in validation_rules or []:
        column = str(rule.get("column", ""))
        operator = str(rule.get("operator", ""))
        value = str(rule.get("value", ""))
        if column not in frame.columns:
            continue
        series = frame[column]
        text = series.astype("string").str.strip()
        matches = text.fillna("").str.casefold()
        if operator == "required":
            invalid_mask = series.isna() | text.eq("")
            description = f"{column} is required"
        elif operator in {"min", "max"}:
            threshold = _parse_number(value)
            if threshold is None:
                continue
            numeric = pd.to_numeric(series, errors="coerce")
            invalid_mask = numeric.notna() & (numeric < threshold if operator == "min" else numeric > threshold)
            description = f"{column} must be {operator}imum {value}"
        elif operator in {"equals", "not_equals", "contains"}:
            expected = value.casefold()
            if operator == "equals":
                invalid_mask = series.notna() & matches.ne(expected)
                description = f"{column} must equal {value!r}"
            elif operator == "not_equals":
                invalid_mask = series.notna() & matches.eq(expected)
                description = f"{column} must not equal {value!r}"
            else:
                invalid_mask = series.notna() & ~matches.str.contains(expected, regex=False, na=False)
                description = f"{column} must contain {value!r}"
        elif operator == "one_of":
            allowed = {item.strip().casefold() for item in value.split(",") if item.strip()}
            if not allowed:
                continue
            invalid_mask = series.notna() & ~matches.isin(allowed)
            description = f"{column} must be one of {value}"
        else:
            continue
        if invalid_mask.any():
            count = int(invalid_mask.sum())
            quality_issues.append({
                "kind": "custom_validation",
                "column": column,
                "count": count,
                "rows": [int(source_rows.iloc[index]) for index in invalid_mask[invalid_mask].index[:10]],
                "message": f"{count:,} value(s) failed custom rule: {description}",
            })
    if repeated_header_rows_removed:
        quality_issues.append({
            "kind": "repeated_header_rows_removed",
            "column": "",
            "count": repeated_header_rows_removed,
            "rows": repeated_header_source_rows[:10],
            "message": f"Removed {repeated_header_rows_removed:,} repeated header row(s)",
        })
    for column in empty_columns_removed:
        quality_issues.append({
            "kind": "empty_column_removed",
            "column": column,
            "count": 1,
            "rows": [],
            "message": f"Removed empty column {column}",
        })

    numeric_types = {"INTEGER", "NUMBER", "MONEY", "PERCENT"}
    for column in frame.columns:
        values = frame[column]
        missing = int(values.isna().sum())
        if missing:
            quality_issues.append({
                "kind": "missing",
                "column": column,
                "count": missing,
                "percent": round(missing / max(1, len(frame)) * 100, 2),
                "message": f"{missing:,} missing value(s) in {column}",
            })
        if types.get(column) in numeric_types:
            numeric = pd.to_numeric(values, errors="coerce").dropna()
            if _column_words(column) & _COUNT_WORDS:
                negatives = numeric[numeric < 0]
            else:
                negatives = numeric.iloc[0:0]
            if not negatives.empty:
                quality_issues.append({
                    "kind": "negative",
                    "column": column,
                    "count": int(len(negatives)),
                    "rows": [int(source_rows.iloc[index]) for index in negatives.index[:10]],
                    "message": f"{len(negatives):,} negative value(s) in {column}",
                })

            out_of_range_rows: set[int] = set()
            source_text = source_values[column].astype("string").str.strip()
            if types.get(column) == "PERCENT":
                over_100 = numeric[numeric > 1]
                out_of_range_rows.update(int(index) for index in over_100.index)

            percent_suffix = source_text.str.endswith("%", na=False)
            scale_values = _numeric_values(source_values.loc[~percent_suffix, column]).dropna()
            if (
                len(scale_values) >= 2
                and scale_values.between(1, 5).all()
            ):
                percent_magnitudes = _numeric_values(source_values[column])
                scale_violations = percent_magnitudes[
                    percent_suffix & percent_magnitudes.gt(5)
                ]
                out_of_range_rows.update(int(index) for index in scale_violations.index)
            if out_of_range_rows:
                issue_rows = sorted(out_of_range_rows)
                quality_issues.append({
                    "kind": "out_of_range",
                    "column": column,
                    "count": len(issue_rows),
                    "rows": [int(source_rows.iloc[index]) for index in issue_rows[:10]],
                    "message": f"{len(issue_rows):,} value(s) are outside the expected range in {column}",
                })

            if len(numeric) >= 4:
                q1, q3 = numeric.quantile([0.25, 0.75])
                spread = q3 - q1
                outliers = numeric[(numeric < q1 - 1.5 * spread) | (numeric > q3 + 1.5 * spread)]
                if len(outliers):
                    quality_issues.append({
                        "kind": "outlier",
                        "column": column,
                        "count": int(len(outliers)),
                        "rows": [int(source_rows.iloc[index]) for index in outliers.index[:10]],
                        "message": f"{len(outliers):,} statistical outlier(s) in {column}",
                    })
    for column, count in issues.items():
        failures = [item for item in conversion_failures if item["column"] == column]
        if failures:
            quality_issues.append({
                "kind": "conversion",
                "column": column,
                "count": count,
                "rows": [item["row"] for item in failures[:10]],
                "message": f"{count:,} value(s) could not be converted in {column}",
            })
    for column, merges in category_merges.items():
        count = sum(int(item["count"]) for item in merges)
        quality_issues.append({
            "kind": "category_merge",
            "column": column,
            "count": count,
            "message": f"Normalized inconsistent capitalization/spacing in {column}",
        })
    for column in frame.columns:
        if "email" not in str(column).casefold():
            continue
        values = frame[column].astype("string")
        invalid_mask = values.notna() & ~values.str.fullmatch(
            r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?)+",
            na=False,
        )
        if invalid_mask.any():
            quality_issues.append({
                "kind": "invalid_email",
                "column": column,
                "count": int(invalid_mask.sum()),
                "rows": [int(source_rows.iloc[index]) for index in invalid_mask[invalid_mask].index[:10]],
                "message": f"{int(invalid_mask.sum()):,} value(s) do not look like valid email addresses in {column}",
            })

    report: dict[str, object] = {
        "input_rows": len(raw),
        "blank_rows_removed": blank_rows,
        "duplicate_rows_removed": duplicate_rows,
        "repeated_header_rows_removed": repeated_header_rows_removed,
        "empty_columns_removed": empty_columns_removed,
        "output_rows": len(frame),
        "column_types": types,
        "conversion_issues": issues,
        "conversion_failures": conversion_failures,
        "category_merges": category_merges,
        "normalize_categories": normalize_categories,
        "normalize_booleans": normalize_booleans,
        "date_order_ambiguous": ambiguous_dates,
        "quality_issues": sorted(quality_issues, key=lambda item: int(item.get("count", 0)), reverse=True),
        "columns": len(frame.columns),
    }
    emit(
        f"Kept {len(frame):,} records; removed {blank_rows:,} blank rows, "
        f"{duplicate_rows:,} duplicates, {repeated_header_rows_removed:,} repeated header rows, "
        f"and {len(empty_columns_removed):,} empty columns."
    )
    return frame, report


def filter_table(
    frame: pd.DataFrame,
    column_types: dict[str, str],
    filters: list[dict[str, str]] | None = None,
) -> pd.DataFrame:
    """Return rows matching every user-selected exploration filter."""
    if not filters:
        return frame.copy()
    mask = pd.Series(True, index=frame.index)
    numeric_types = {"MONEY", "NUMBER", "INTEGER", "PERCENT"}
    for item in filters:
        column = str(item.get("column", ""))
        operator = str(item.get("operator", ""))
        value = str(item.get("value", ""))
        if column not in frame.columns:
            raise ValueError(f"Filter column {column!r} does not exist.")
        series = frame[column]
        text = series.astype("string").fillna("").str.strip()
        if operator == "is_empty":
            current = series.isna() | text.eq("")
        elif operator == "is_not_empty":
            current = series.notna() & text.ne("")
        elif column_types.get(column) in numeric_types:
            expected = _parse_number(value)
            if expected is None:
                raise ValueError(f"Enter a valid number for the {column!r} filter.")
            numeric = series.map(_parse_number)
            if operator == "equals":
                current = numeric.eq(expected)
            elif operator == "not_equals":
                current = numeric.ne(expected)
            elif operator == "greater_than":
                current = numeric.gt(expected)
            elif operator == "less_than":
                current = numeric.lt(expected)
            else:
                raise ValueError(f"Unsupported filter operator {operator!r}.")
        else:
            expected = value.casefold()
            lowered = text.str.casefold()
            if operator == "equals":
                current = lowered.eq(expected)
            elif operator == "not_equals":
                current = lowered.ne(expected)
            elif operator == "contains":
                current = lowered.str.contains(expected, regex=False, na=False)
            elif operator == "greater_than":
                current = lowered.gt(expected)
            elif operator == "less_than":
                current = lowered.lt(expected)
            else:
                raise ValueError(f"Unsupported filter operator {operator!r}.")
        mask &= current.fillna(False)
    return frame.loc[mask].copy()


def suggest_revenue_columns(
    frame: pd.DataFrame,
    column_types: dict[str, str],
) -> dict[str, str] | None:
    """Suggest (never apply) a revenue calculation when semantic name hints exist."""
    def find_column(hints: tuple[str, ...], allowed: set[str]) -> str | None:
        for column in frame.columns:
            words = _column_words(column)
            if column_types.get(column) in allowed and words.intersection(hints):
                return column
        return None

    quantity = find_column(("quantity", "qty", "units", "items"), {"INTEGER", "NUMBER"})
    price = find_column(("price", "unitprice", "cost"), {"MONEY", "NUMBER"})
    discount = find_column(("discount", "rebate"), {"PERCENT"}) if quantity and price else None
    if not quantity or not price or any(str(column).casefold() == "revenue" for column in frame.columns):
        return None
    suggestion = {"quantity": quantity, "price": price}
    if discount:
        suggestion["discount"] = discount
    return suggestion


def add_computed_revenue(frame: pd.DataFrame, suggestion: dict[str, str]) -> pd.DataFrame:
    """Return a copy containing an explicitly requested revenue calculation."""
    result = frame.copy()
    quantity = pd.to_numeric(result[suggestion["quantity"]], errors="coerce")
    price = pd.to_numeric(result[suggestion["price"]], errors="coerce")
    revenue = quantity * price
    if "discount" in suggestion:
        discount = pd.to_numeric(result[suggestion["discount"]], errors="coerce").fillna(0)
        revenue = revenue * (1 - discount)
    result["Computed Revenue"] = revenue.round(2)
    return result
