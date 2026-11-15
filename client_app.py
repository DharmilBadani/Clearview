"""Local web interface for ClearView. Start with `python client_app.py`."""

from __future__ import annotations

import atexit
import os
import shutil
import tempfile
import time
import uuid
from datetime import datetime, timedelta
from pathlib import Path
from threading import RLock, Timer
from urllib.parse import urlsplit

import pandas as pd
from flask import Flask, jsonify, render_template, request, send_from_directory
from werkzeug.datastructures import FileStorage
from werkzeug.utils import secure_filename

from data_engine import (
    TYPE_CHOICES,
    add_computed_revenue,
    clean_table,
    date_order_info,
    detect_column_types,
    filter_table,
    load_table,
    suggest_revenue_columns,
)
from report_builder import EXCEL_MAX_ROWS, build_excel_report

PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


def _environment_integer(name: str, default: int) -> int:
    try:
        return int(os.environ.get(name, str(default)))
    except ValueError:
        raise RuntimeError(f"{name} must be a whole number.") from None


def _environment_bool(name: str, default: bool) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    normalized = value.strip().casefold()
    if normalized in {"1", "true", "yes", "on"}:
        return True
    if normalized in {"0", "false", "no", "off"}:
        return False
    raise RuntimeError(f"{name} must be true/false.")


def _environment_list(name: str) -> set[str]:
    return {item.strip().casefold() for item in os.environ.get(name, "").split(",") if item.strip()}


MAX_UPLOAD_BYTES = max(1, _environment_integer("CLEARVIEW_MAX_UPLOAD_MB", 100)) * 1024 * 1024
MAX_SESSIONS = max(1, _environment_integer("CLEARVIEW_MAX_SESSIONS", 20))
SESSION_TTL_SECONDS = max(60, _environment_integer("CLEARVIEW_SESSION_TTL_MINUTES", 30) * 60)
OUTPUT_TTL_DAYS = max(1, _environment_integer("CLEARVIEW_OUTPUT_TTL_DAYS", 7))
LOCAL_PORT = max(1, min(65535, _environment_integer("CLEARVIEW_PORT", int(os.environ.get("PORT", "8765")))))
BIND_HOST = os.environ.get("CLEARVIEW_HOST", "127.0.0.1").strip() or "127.0.0.1"
RUN_MODE = os.environ.get("CLEARVIEW_ENV", "development").strip().casefold()
OPEN_BROWSER = _environment_bool("CLEARVIEW_OPEN_BROWSER", RUN_MODE != "production")
SERVER_THREADS = max(1, _environment_integer("CLEARVIEW_THREADS", 4))
TRUSTED_HOSTS = _environment_list("CLEARVIEW_ALLOWED_HOSTS")
ALLOWED_ORIGINS = _environment_list("CLEARVIEW_ALLOWED_ORIGINS")
ALLOWED_EXTENSIONS = {".csv", ".tsv", ".tab", ".txt", ".xlsx", ".xlsm", ".xls"}
UPLOAD_ROOT = Path(tempfile.mkdtemp(prefix="clearview-uploads-"))
_datasets: dict[str, dict[str, object]] = {}
_datasets_lock = RLock()

app = Flask(__name__, template_folder="templates", static_folder="static")
app.config.update(
    MAX_CONTENT_LENGTH=MAX_UPLOAD_BYTES,
    OUTPUT_DIR=OUTPUT_DIR,
    LOCAL_PORT=LOCAL_PORT,
    BIND_HOST=BIND_HOST,
    RUN_MODE=RUN_MODE,
)


@app.context_processor
def inject_upload_limit():
    return {"max_upload_mb": MAX_UPLOAD_BYTES // (1024 * 1024)}


def _cleanup_uploads() -> None:
    shutil.rmtree(UPLOAD_ROOT, ignore_errors=True)


def _cleanup_old_outputs(now: float | None = None) -> None:
    current_time = time.time() if now is None else now
    threshold = current_time - timedelta(days=OUTPUT_TTL_DAYS).total_seconds()
    output_dir: Path = Path(app.config["OUTPUT_DIR"])
    output_dir.mkdir(parents=True, exist_ok=True)
    for path in output_dir.iterdir():
        try:
            is_generated = (
                path.name.startswith(("Analysis_", "Cleaned_"))
                and path.suffix.casefold() in {".xlsx", ".csv"}
            )
            if is_generated and path.is_file() and path.stat().st_mtime < threshold:
                path.unlink()
        except OSError:
            app.logger.warning("Could not remove expired output file %s", path.name)


def _remove_session(session_id: str, dataset: dict[str, object]) -> None:
    _datasets.pop(session_id, None)
    path = Path(str(dataset["path"]))
    try:
        path.unlink(missing_ok=True)
    except OSError:
        app.logger.warning("Could not remove expired upload %s", path.name)


def _expire_sessions(now: float | None = None) -> None:
    current_time = time.time() if now is None else now
    with _datasets_lock:
        expired = [
            session_id
            for session_id, dataset in _datasets.items()
            if current_time - float(dataset["last_access"]) > SESSION_TTL_SECONDS
        ]
        for session_id in expired:
            dataset = _datasets.get(session_id)
            if dataset:
                _remove_session(session_id, dataset)


def _get_session(session_id: str) -> dict[str, object] | None:
    _expire_sessions()
    with _datasets_lock:
        dataset = _datasets.get(session_id)
        if dataset is not None:
            dataset["last_access"] = time.time()
        return dataset


def _store_session(
    session_id: str,
    *,
    path: Path,
    name: str,
    sheets: list[str],
    selected_sheet: str | None,
    types: dict[str, str],
) -> None:
    _expire_sessions()
    with _datasets_lock:
        while len(_datasets) >= MAX_SESSIONS:
            oldest_id = min(_datasets, key=lambda key: float(_datasets[key]["last_access"]))
            _remove_session(oldest_id, _datasets[oldest_id])
        _datasets[session_id] = {
            "path": path,
            "name": name,
            "sheets": sheets,
            "sheet": selected_sheet,
            "types": types,
            "last_access": time.time(),
        }


def _read_dataset(dataset: dict[str, object], sheet_name: str | None = None) -> tuple[pd.DataFrame, dict[str, str]]:
    frame, _ = load_table(
        Path(str(dataset["path"])),
        sheet_name if sheet_name is not None else dataset.get("sheet"),
    )
    if frame.empty:
        raise ValueError("This file has headers but no data rows.")
    if len(frame) + 1 > EXCEL_MAX_ROWS:
        raise ValueError(
            f"This file exceeds Excel's {EXCEL_MAX_ROWS - 1:,}-data-row worksheet limit. "
            "Please split it into smaller files."
        )
    types = detect_column_types(frame)
    return frame, types


def _validated_filters(value: object, types: dict[str, str]) -> list[dict[str, str]]:
    if value in (None, []):
        return []
    if not isinstance(value, list):
        raise ValueError("Filters must be supplied as a list.")
    allowed = {"equals", "not_equals", "contains", "greater_than", "less_than", "is_empty", "is_not_empty"}
    result = []
    for item in value:
        if not isinstance(item, dict):
            continue
        column = item.get("column")
        operator = item.get("operator")
        if not isinstance(column, str) or column not in types or operator not in allowed:
            continue
        result.append({"column": column, "operator": str(operator), "value": str(item.get("value", ""))})
    return result


def _validated_rules(value: object, types: dict[str, str]) -> list[dict[str, str]]:
    if value in (None, []):
        return []
    if not isinstance(value, list):
        raise ValueError("Validation rules must be supplied as a list.")
    allowed = {"required", "min", "max", "equals", "not_equals", "contains", "one_of"}
    result = []
    for item in value:
        if not isinstance(item, dict):
            continue
        column = item.get("column")
        operator = item.get("operator")
        if not isinstance(column, str) or column not in types or operator not in allowed:
            continue
        result.append({"column": column, "operator": str(operator), "value": str(item.get("value", ""))})
    return result


def _json_value(value: object) -> object:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.isoformat(sep=" ")
    if hasattr(value, "item"):
        value = value.item()
    return value


def _preview(frame: pd.DataFrame, limit: int = 12) -> list[dict[str, object]]:
    return [
        {str(column): _json_value(value) for column, value in zip(frame.columns, row)}
        for row in frame.head(limit).itertuples(index=False, name=None)
    ]


def _upload_payload(session_id: str, dataset: dict[str, object], frame: pd.DataFrame, types: dict[str, str]):
    ambiguous_dates = {
        column: date_order_info(frame[column])[1]
        for column, kind in types.items()
        if kind == "DATE"
    }
    return {
        "session_id": session_id,
        "filename": dataset["name"],
        "rows": len(frame),
        "column_count": len(frame.columns),
        "sheets": dataset["sheets"],
        "columns": [
            {"name": str(column), "type": types[column], "date_order_ambiguous": ambiguous_dates.get(column, False)}
            for column in frame.columns
        ],
        "revenue_suggestion": suggest_revenue_columns(frame, types),
        "preview": _preview(frame),
    }


def _host_is_allowed(host: str, port: int | None, allowed_port: int) -> bool:
    local = (
        host.casefold() in {"localhost", "127.0.0.1", "::1"}
        and (port == allowed_port or (port is None and app.testing))
    )
    trusted = ("*" in TRUSTED_HOSTS or host.casefold() in TRUSTED_HOSTS) and (
        port == allowed_port or port is None
    )
    return local or trusted


@app.before_request
def enforce_local_request_origin():
    host_parts = urlsplit("//" + request.host)
    host = host_parts.hostname or ""
    try:
        port = host_parts.port
    except ValueError:
        return jsonify({"error": "Invalid local host."}), 403
    allowed_port = int(app.config["LOCAL_PORT"])
    if not _host_is_allowed(host, port, allowed_port):
        return jsonify({"error": "ClearView only accepts requests addressed to this local server."}), 403
    origin = request.headers.get("Origin")
    if origin:
        origin_parts = urlsplit(origin)
        try:
            origin_port = origin_parts.port
        except ValueError:
            return jsonify({"error": "Invalid request origin."}), 403
        if (
            origin_parts.scheme not in {"http", "https"}
            or (
                origin_parts.netloc.casefold() != request.host.casefold()
                and origin.casefold() not in ALLOWED_ORIGINS
            )
        ):
            return jsonify({"error": "Cross-origin requests are not allowed."}), 403
    _cleanup_old_outputs()


@app.after_request
def add_security_headers(response):
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "DENY")
    response.headers.setdefault("Referrer-Policy", "same-origin")
    response.headers.setdefault(
        "Content-Security-Policy",
        "default-src 'self'; script-src 'self'; style-src 'self'; "
        "img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; "
        "base-uri 'self'; form-action 'self'",
    )
    if request.is_secure:
        response.headers.setdefault("Strict-Transport-Security", "max-age=31536000; includeSubDomains")
    return response


@app.get("/")
def home():
    return render_template("index.html")


@app.get("/api/health")
def health():
    return jsonify({"status": "ok"})


def _save_upload(uploaded_file, original_name: str):
    extension = Path(original_name).suffix.casefold()
    if extension not in ALLOWED_EXTENSIONS:
        return jsonify({"error": "Upload a CSV, TSV, TXT, XLS, XLSX, or XLSM file."}), 400

    session_id = uuid.uuid4().hex
    stored_path = UPLOAD_ROOT / f"{session_id}{extension}"
    uploaded_file.save(stored_path)
    dataset: dict[str, object] = {
        "path": stored_path,
        "name": secure_filename(original_name) or "uploaded-data",
        "sheets": [],
        "sheet": None,
    }
    try:
        frame, sheets = load_table(stored_path)
        if frame.empty:
            raise ValueError("This file has headers but no data rows.")
        if len(frame) + 1 > EXCEL_MAX_ROWS:
            raise ValueError(
                f"This file exceeds Excel's {EXCEL_MAX_ROWS - 1:,}-data-row worksheet limit. "
                "Please split it into smaller files."
            )
        types = detect_column_types(frame)
        dataset["sheets"] = sheets
        _store_session(
            session_id,
            path=stored_path,
            name=str(dataset["name"]),
            sheets=sheets,
            selected_sheet=None,
            types=types,
        )
        return jsonify(_upload_payload(session_id, dataset, frame, types))
    except (OSError, UnicodeError, ValueError) as error:
        stored_path.unlink(missing_ok=True)
        return jsonify({"error": str(error)}), 400
    except Exception:
        stored_path.unlink(missing_ok=True)
        app.logger.exception("Upload could not be read")
        return jsonify({"error": "The uploaded file could not be read. Check the server log for details."}), 500


@app.post("/api/upload")
def upload():
    uploaded_file = request.files.get("file")
    if uploaded_file is None or not uploaded_file.filename:
        return jsonify({"error": "Choose a file to upload."}), 400
    original_name = Path(uploaded_file.filename).name
    return _save_upload(uploaded_file, original_name)


@app.post("/api/sample")
def sample():
    sample_path = PROJECT_ROOT / "data" / "raw_dirty_sales_data.csv"
    if not sample_path.is_file():
        return jsonify({"error": "The optional sample dataset is not available."}), 404
    with sample_path.open("rb") as source:
        uploaded_file = FileStorage(stream=source, filename=sample_path.name)
        return _save_upload(uploaded_file, sample_path.name)


@app.post("/api/sheet")
def change_sheet():
    payload = request.get_json(silent=True) or {}
    session_id = str(payload.get("session_id", ""))
    dataset = _get_session(session_id)
    if dataset is None:
        return jsonify({"error": "Upload the workbook again; this session is no longer available."}), 404

    selected = payload.get("sheet")
    if selected not in dataset["sheets"]:
        return jsonify({"error": "Select a worksheet from this workbook."}), 400
    try:
        frame, types = _read_dataset(dataset, str(selected))
        dataset["sheet"] = str(selected)
        dataset["types"] = types
    except (OSError, UnicodeError, ValueError) as error:
        return jsonify({"error": str(error)}), 400
    return jsonify(_upload_payload(session_id, dataset, frame, types))


@app.post("/api/explore")
def explore():
    payload = request.get_json(silent=True) or {}
    session_id = str(payload.get("session_id", ""))
    dataset = _get_session(session_id)
    if dataset is None:
        return jsonify({"error": "Upload your file again; this session is no longer available."}), 404
    try:
        frame, types = _read_dataset(dataset)
        filters = _validated_filters(payload.get("filters"), types)
        filtered = filter_table(frame, types, filters)
    except (OSError, UnicodeError, ValueError) as error:
        return jsonify({"error": str(error)}), 400
    return jsonify({
        "total_rows": len(frame),
        "filtered_rows": len(filtered),
        "preview": _preview(filtered),
        "filters": filters,
    })


def _csv_export_copy(frame: pd.DataFrame, column_types: dict[str, str]) -> pd.DataFrame:
    export = frame.copy()
    for column, kind in column_types.items():
        if column not in export.columns or kind != "DATE":
            continue
        parsed = pd.to_datetime(export[column], errors="coerce")
        formatted = parsed.dt.strftime("%Y-%m-%d")
        export[column] = formatted.where(parsed.notna(), "")
    return export


@app.post("/api/analyze")
def analyze():
    payload = request.get_json(silent=True) or {}
    session_id = str(payload.get("session_id", ""))
    dataset = _get_session(session_id)
    if dataset is None:
        return jsonify({"error": "Upload your file again; this session is no longer available."}), 404

    supplied_types = payload.get("types", {})
    if not isinstance(supplied_types, dict):
        return jsonify({"error": "Column types must be supplied as an object."}), 400
    try:
        raw, detected_types = _read_dataset(dataset)
    except (OSError, UnicodeError, ValueError) as error:
        return jsonify({"error": str(error)}), 400
    types = dict(detected_types)
    for column, kind in supplied_types.items():
        if column in types and kind in TYPE_CHOICES:
            types[column] = kind

    measure = payload.get("measure")
    dimension = payload.get("dimension")
    if not isinstance(measure, str):
        measure = None
    if not isinstance(dimension, str):
        dimension = None
    if measure not in types or (measure and types[measure] not in {"MONEY", "NUMBER", "INTEGER", "PERCENT"}):
        measure = None
    if dimension not in types or (dimension and types[dimension] not in {"CATEGORY", "BOOLEAN"}):
        dimension = None
    normalize_categories = payload.get("normalize_categories", True) is not False
    normalize_booleans = payload.get("normalize_booleans", True) is not False
    date_orders = payload.get("date_orders", {})
    if not isinstance(date_orders, dict):
        return jsonify({"error": "Date order preferences must be supplied as an object."}), 400
    validated_date_orders = {
        column: order
        for column, order in date_orders.items()
        if isinstance(order, str)
        and column in types
        and types[column] == "DATE"
        and order in {"day_first", "month_first"}
    }

    try:
        filters = _validated_filters(payload.get("filters"), types)
        validation_rules = _validated_rules(payload.get("validation_rules"), types)
        raw = filter_table(raw, types, filters)
        clean, report = clean_table(
            raw,
            types,
            normalize_categories=normalize_categories,
            date_orders=validated_date_orders,
            normalize_booleans=normalize_booleans,
            validation_rules=validation_rules,
        )
        revenue_suggestion = suggest_revenue_columns(clean, types)
        if payload.get("derive_revenue") is True and revenue_suggestion:
            clean = add_computed_revenue(clean, revenue_suggestion)
            types["Computed Revenue"] = "MONEY"
            report["derived_columns"] = ["Computed Revenue"]
            if measure is None:
                measure = "Computed Revenue"
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        stem = Path(str(dataset["name"])).stem or "data"
        clean_name = f"Cleaned_{stem}_{stamp}.csv"
        report_name = f"Analysis_{stem}_{stamp}.xlsx"
        output_dir: Path = Path(app.config["OUTPUT_DIR"])
        output_dir.mkdir(parents=True, exist_ok=True)
        _csv_export_copy(clean, types).to_csv(output_dir / clean_name, index=False, encoding="utf-8-sig")
        build_excel_report(
            raw,
            clean,
            types,
            report,
            output_dir / report_name,
            measure=measure,
            dimensions=[dimension] if dimension else None,
        )
    except (OSError, ValueError, PermissionError) as error:
        return jsonify({"error": str(error)}), 400
    except Exception:
        app.logger.exception("Report generation failed")
        return jsonify({"error": "The report could not be generated. Check the server log for details."}), 500

    priority = {
        "repeated_header_rows_removed": 0,
        "empty_column_removed": 0,
        "conversion": 1,
        "out_of_range": 2,
        "invalid_email": 3,
        "negative": 4,
        "missing": 5,
        "outlier": 6,
    }
    review_issues = sorted(
        (issue for issue in report["quality_issues"] if issue["kind"] != "category_merge"),
        key=lambda issue: (priority.get(str(issue["kind"]), 5), -int(issue["count"])),
    )[:5]
    return jsonify(
        {
            "clean_rows": len(clean),
            "filtered_rows": len(raw),
            "filters_applied": filters,
            "validation_rules_applied": validation_rules,
            "columns": len(clean.columns),
            "blank_rows_removed": report["blank_rows_removed"],
            "duplicates_removed": report["duplicate_rows_removed"],
            "repeated_header_rows_removed": report["repeated_header_rows_removed"],
            "empty_columns_removed": report["empty_columns_removed"],
            "conversion_issues": sum(report["conversion_issues"].values()),
            "clean_csv": clean_name,
            "report": report_name,
            "preview": _preview(clean),
            "insights": [
                {
                    "column": column,
                    "type": types[column],
                    "missing": int(clean[column].isna().sum()),
                    "unique": int(clean[column].nunique(dropna=True)),
                }
                for column in clean.columns
            ],
            "quality_issue_count": len(report["quality_issues"]),
            "quality_issues": review_issues,
            "category_merges": report["category_merges"],
            "date_order_ambiguous": report["date_order_ambiguous"],
        }
    )


@app.get("/download/<path:filename>")
def download(filename: str):
    safe_name = Path(filename).name
    if safe_name != filename or not safe_name.startswith(("Analysis_", "Cleaned_")):
        return jsonify({"error": "File not found."}), 404
    return send_from_directory(Path(app.config["OUTPUT_DIR"]), safe_name, as_attachment=True)


@app.errorhandler(413)
def request_too_large(error):
    limit_mb = int(app.config["MAX_CONTENT_LENGTH"]) // (1024 * 1024)
    return jsonify({"error": f"Files must be at most {limit_mb} MB."}), error.code


def main() -> None:
    host = os.environ.get("CLEARVIEW_HOST", BIND_HOST)
    port = int(os.environ.get("CLEARVIEW_PORT", os.environ.get("PORT", str(LOCAL_PORT))))
    app.config["LOCAL_PORT"] = port
    display_host = "127.0.0.1" if host in {"0.0.0.0", "::"} else host
    url = f"http://{display_host}:{port}"

    def open_browser() -> None:
        import webbrowser

        webbrowser.open(url)

    if OPEN_BROWSER and RUN_MODE != "production":
        Timer(1.0, open_browser).start()
    print(f"ClearView is starting at {url}")
    print("Keep this window open while you use the app. Press Ctrl+C to stop.")
    if RUN_MODE == "production":
        try:
            from waitress import serve
        except ImportError as error:
            raise RuntimeError("Production mode requires Waitress. Install requirements.txt first.") from error
        serve(app, host=host, port=port, threads=SERVER_THREADS)
    else:
        app.run(host=host, port=port, debug=False, threaded=True, use_reloader=False)


atexit.register(_cleanup_uploads)


if __name__ == "__main__":
    main()
