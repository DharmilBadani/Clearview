"""Schema-agnostic loading, type detection, and cleaning for tabular files."""

from __future__ import annotations

import csv
import re
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

_NULL_VALUES = {"", "null", "none", "nan", "n/a", "n.a."}
_MISSING_CONVERSION_MARKERS = {"", "?", "tbd", "00/00/0000", "na", "n/a", "n.a.", "none", "null", "nan", "-"}
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
    value = value.strip()
    return pd.NA if value.casefold() in _NULL_VALUES else value


def load_table(path: str | Path, sheet_name: str | None = None) -> tuple[pd.DataFrame, list[str]]:
    """Load one selected worksheet or a delimited text file as strings."""
    file_path = Path(path)
    suffix = file_path.suffix.casefold()

    if suffix in {".xlsx", ".xlsm", ".xls"}:
        with pd.ExcelFile(file_path) as workbook:
            sheets = list(workbook.sheet_names)
            if not sheets:
                raise ValueError("The Excel workbook does not contain any worksheets.")
            selected_sheet = sheet_name if sheet_name in sheets else sheets[0]
            parsed_file = pd.read_excel(
                workbook,
                sheet_name=selected_sheet,
                header=None,
                dtype=str,
                keep_default_na=False,
            )
            if parsed_file.empty:
                raise ValueError("The selected file does not contain a header row.")
            headers = parsed_file.iloc[0].tolist()
            frame = parsed_file.iloc[1:].copy()
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
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                try:
                    parsed_file = pd.read_csv(
                        file_path,
                        sep=delimiter,
                        header=None,
                        dtype=str,
                        keep_default_na=False,
                        encoding=encoding,
                        low_memory=False,
                    )
                except pd.errors.EmptyDataError as error:
                    raise ValueError("The selected file does not contain a header row.") from error
                if parsed_file.empty:
                    raise ValueError("The selected file does not contain a header row.")
                headers = parsed_file.iloc[0].tolist()
                frame = parsed_file.iloc[1:].copy()
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

    headers: list[str] = []
    seen_headers: dict[str, int] = {}
    used_headers: set[str] = set()
    for index, column in enumerate(frame.columns):
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
    frame.columns = headers
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


def _date_values(series: pd.Series) -> pd.Series:
    text = series.astype("string").str.strip()
    present = text.notna() & ~text.str.casefold().isin(_NULL_VALUES)
    day_first, _ = date_order_info(text[present])
    parsed = pd.to_datetime(text, format="mixed", errors="coerce", dayfirst=day_first)
    return parsed


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


def detect_column_types(frame: pd.DataFrame) -> dict[str, str]:
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

        parsed_dates = _date_values(values)
        date_ratio = parsed_dates.notna().mean()
        if date_ratio >= 0.5:
            result[column] = "DATE"
            continue

        numeric = _numeric_values(values)
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
    for column in frame.columns:
        frame[column] = frame[column].map(_clean_cell_text)
    source_rows = pd.Series(range(2, len(frame) + 2), index=frame.index)

    repeated_header_mask = _repeated_header_mask(frame)
    repeated_header_rows_removed = int(repeated_header_mask.sum())
    repeated_header_source_rows = [
        int(source_rows.loc[index])
        for index in repeated_header_mask[repeated_header_mask].index
    ]
    frame = frame.loc[~repeated_header_mask].copy()
    source_rows = source_rows.loc[frame.index]

    empty_columns_removed = [
        column for column in frame.columns
        if not frame[column].notna().any()
    ]
    if empty_columns_removed:
        frame.drop(columns=empty_columns_removed, inplace=True)

    blank_mask = frame.isna().all(axis=1) if len(frame.columns) else pd.Series(True, index=frame.index)
    blank_rows = int(blank_mask.sum())
    frame = frame.loc[~blank_mask].copy()
    source_rows = source_rows.loc[frame.index]
    before_duplicates = len(frame)
    unique_mask = ~frame.duplicated(keep="first")
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
    issues: dict[str, int] = {}
    conversion_failures: list[dict[str, object]] = []
    category_merges: dict[str, list[dict[str, object]]] = {}
    ambiguous_dates: dict[str, bool] = {}
    date_orders = date_orders or {}
    for column in frame.columns:
        kind = types.get(column, "TEXT")
        source = frame[column]
        if kind in {"MONEY", "NUMBER", "INTEGER", "PERCENT"}:
            parsed = _numeric_values(source)
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
            inferred_day_first, ambiguous = date_order_info(source)
            day_first = requested_order == "day_first" if requested_order in {"day_first", "month_first"} else inferred_day_first
            text = source.astype("string").str.strip()
            parsed = pd.to_datetime(text, format="mixed", errors="coerce", dayfirst=day_first)
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
