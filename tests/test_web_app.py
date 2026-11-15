import csv
import base64
import io
import os
import re
import tempfile
import time
import unittest
from pathlib import Path

import client_app
from openpyxl import load_workbook


class WebAppTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        client_app.app.config.update(TESTING=True, OUTPUT_DIR=Path(self.temp_dir.name))
        self.client = client_app.app.test_client()
        self.session_id = None

    def tearDown(self):
        if self.session_id:
            dataset = client_app._datasets.pop(self.session_id, None)
            if dataset:
                Path(dataset["path"]).unlink(missing_ok=True)
        self.temp_dir.cleanup()

    def _upload_fixture(self, filename):
        fixture_path = Path(__file__).parent / "fixtures" / filename
        with fixture_path.open("rb") as source:
            response = self.client.post(
                "/api/upload",
                data={"file": (source, filename)},
                content_type="multipart/form-data",
            )
        self.assertEqual(response.status_code, 200, response.get_json())
        self.session_id = response.get_json()["session_id"]

    @staticmethod
    def _conversion_detail_rows(quality_sheet):
        rows = list(quality_sheet.iter_rows(values_only=True))
        header_index = next(
            index
            for index, row in enumerate(rows)
            if tuple(row[:4]) == ("column", "value", "row", "reason")
        )
        return [row[:4] for row in rows[header_index + 1:] if row[0] is not None]

    def test_browser_flow_upload_analyze_and_download_report(self):
        response = self.client.post(
            "/api/upload",
            data={
                "file": (
                    io.BytesIO(
                        b"Date,Team,Amount,Note,Formula Plus,Formula Minus,Formula At,Contact Phone,Dash\n"
                        b"2024-01-01,North,$12.50,=1+1,+SUM(1),-1+2,@SUM(A1:A2),+91 98765 43210,-\n"
                        b"2024-02-01,South,$24.00,second,plain,plain,plain,+91 99999 00000,-\n"
                        b"2024-03-01,North,$19.50,third,plain,plain,plain,+91 88888 11111,-\n"
                    ),
                    "sample.csv",
                )
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 200)
        uploaded = response.get_json()
        self.session_id = uploaded["session_id"]
        self.assertEqual(uploaded["rows"], 3)
        self.assertEqual(uploaded["columns"][2]["type"], "MONEY")
        self.assertNotIn("frame", client_app._datasets[self.session_id])

        result_response = self.client.post(
            "/api/analyze",
            json={
                "session_id": self.session_id,
                "types": {
                    "Date": "DATE",
                    "Team": "CATEGORY",
                    "Amount": "MONEY",
                    "Note": "TEXT",
                },
                "measure": "Amount",
                "dimension": "Team",
            },
        )
        self.assertEqual(result_response.status_code, 200)
        result = result_response.get_json()
        self.assertEqual(result["clean_rows"], 3)
        self.assertTrue((Path(self.temp_dir.name) / result["report"]).is_file())
        download_response = self.client.get(f"/download/{result['report']}")
        self.assertEqual(download_response.status_code, 200)
        report_book = load_workbook(io.BytesIO(download_response.get_data()))
        cleaned_sheet = report_book["Cleaned Data"]
        headers = [cell.value for cell in cleaned_sheet[1]]
        for column, value in {
            "Note": "=1+1",
            "Formula Plus": "+SUM(1)",
            "Formula Minus": "-1+2",
            "Formula At": "@SUM(A1:A2)",
        }.items():
            cell = cleaned_sheet.cell(2, headers.index(column) + 1)
            self.assertEqual(cell.value, value)
            self.assertEqual(cell.data_type, "s")
            self.assertTrue(cell.quotePrefix)
        for column, value in {"Contact Phone": "+91 98765 43210", "Dash": None}.items():
            cell = cleaned_sheet.cell(2, headers.index(column) + 1)
            self.assertEqual(cell.value, value)
            self.assertFalse(cell.quotePrefix)
        report_book.close()
        download_response.close()
        csv_response = self.client.get(f"/download/{result['clean_csv']}")
        self.assertEqual(csv_response.status_code, 200)
        csv_content = csv_response.get_data()
        csv_response.close()
        self.assertIn(b"=1+1", csv_content)
        self.assertNotIn(b"'=1+1", csv_content)
        self.assertIn(b"+91 98765 43210", csv_content)
        self.assertTrue(csv_content.splitlines()[1].endswith(b","))

    def test_inventory_upload_cleans_rows_dates_and_quality_report(self):
        self._upload_fixture("messy_test_inventory_semicolon.csv")
        response = self.client.post("/api/analyze", json={"session_id": self.session_id})

        self.assertEqual(response.status_code, 200, response.get_json())
        result = response.get_json()
        self.assertEqual(result["clean_rows"], 400)
        self.assertEqual(result["repeated_header_rows_removed"], 1)
        self.assertEqual(result["empty_columns_removed"], ["Column_13"])
        self.assertNotIn("Column_13", {item["column"] for item in result["insights"]})

        csv_response = self.client.get(f"/download/{result['clean_csv']}")
        clean_csv = csv_response.get_data(as_text=True).lstrip("\ufeff")
        csv_response.close()
        clean_rows = list(csv.DictReader(io.StringIO(clean_csv)))
        self.assertEqual(len(clean_rows), 400)
        self.assertNotIn("Column_13", clean_rows[0])
        self.assertEqual({row["Active?"] for row in clean_rows if row["Active?"]}, {"Yes", "No"})
        stock_values = [row["Stock On Hand"] for row in clean_rows if row["Stock On Hand"]]
        self.assertTrue(all(not value.endswith(".0") for value in stock_values))
        restock_dates = [row["Last Restock"] for row in clean_rows if row["Last Restock"]]
        self.assertTrue(all(re.fullmatch(r"\d{4}-\d{2}-\d{2}", value) for value in restock_dates))

        report_book = load_workbook(Path(self.temp_dir.name) / result["report"], read_only=True)
        quality_sheet = report_book["Data Quality"]
        summary = {
            (row[0], row[1]): row
            for row in quality_sheet.iter_rows(min_row=5, values_only=True)
            if row[0] is not None
        }
        self.assertEqual(summary[("conversion", "Last Restock")][2], 45)
        repeated_header_issue = next(
            row for (kind, _), row in summary.items()
            if kind == "repeated_header_rows_removed"
        )
        self.assertEqual(repeated_header_issue[2], 1)
        self.assertEqual(summary[("empty_column_removed", "Column_13")][2], 1)
        restock_failures = [
            row for row in self._conversion_detail_rows(quality_sheet)
            if row[0] == "Last Restock"
        ]
        self.assertEqual(len(restock_failures), 45)
        self.assertTrue(all(isinstance(row[2], int) for row in restock_failures))

        cleaned_sheet = report_book["Cleaned Data"]
        headers = [cell.value for cell in cleaned_sheet[1]]
        self.assertNotIn("Column_13", headers)
        stock_index = headers.index("Stock On Hand") + 1
        stock_cell = next(
            cleaned_sheet.cell(row, stock_index)
            for row in range(2, cleaned_sheet.max_row + 1)
            if cleaned_sheet.cell(row, stock_index).value is not None
        )
        self.assertEqual(stock_cell.number_format, "#,##0")
        self.assertIsInstance(stock_cell.value, int)
        report_book.close()

    def test_survey_upload_types_fields_and_reports_invalid_values(self):
        self._upload_fixture("messy_test_survey.txt")
        response = self.client.post("/api/analyze", json={"session_id": self.session_id})

        self.assertEqual(response.status_code, 200, response.get_json())
        result = response.get_json()
        types = {item["column"]: item["type"] for item in result["insights"]}
        self.assertEqual(types["Age"], "NUMBER")
        self.assertEqual(types["Income"], "MONEY")
        self.assertEqual(types["Submitted"], "DATE")

        report_book = load_workbook(Path(self.temp_dir.name) / result["report"], read_only=True)
        failures = self._conversion_detail_rows(report_book["Data Quality"])
        self.assertIn(("Submitted", "yesterday", 5, "Could not parse date"), failures)
        self.assertNotIn(("Age", "forty", 8, "Could not convert value"), failures)
        self.assertNotIn(("Income", "unknown", 10, "Could not convert value"), failures)
        report_book.close()

    def test_analyze_exports_iso_dates_and_honors_boolean_normalization_toggle(self):
        response = self.client.post(
            "/api/upload",
            data={
                "file": (
                    io.BytesIO(
                        b'Joined,Active,Salary\n'
                        b'2023-01-01,yes,85k\n'
                        b'2023-13-45,no,TBD\n'
                        b'2023-02-03,TRUE,"$95,000"\n'
                    ),
                    "employees.csv",
                )
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 200)
        self.session_id = response.get_json()["session_id"]

        normalized_response = self.client.post(
            "/api/analyze",
            json={"session_id": self.session_id},
        )
        self.assertEqual(normalized_response.status_code, 200)
        normalized = normalized_response.get_json()
        self.assertGreater(normalized["conversion_issues"], 0)
        report_book = load_workbook(Path(self.temp_dir.name) / normalized["report"], read_only=True)
        selected_dimensions = next(
            row[1].value
            for row in report_book["Overview"].iter_rows(min_col=1, max_col=2)
            if row[0].value == "Selected chart dimension(s)"
        )
        report_book.close()
        self.assertEqual(selected_dimensions, "Active")
        csv_response = self.client.get(f"/download/{normalized['clean_csv']}")
        normalized_csv = csv_response.get_data(as_text=True).lstrip("\ufeff")
        csv_response.close()
        self.assertIn("2023-01-01,Yes,85000.0", normalized_csv)
        self.assertIn(",No,", normalized_csv)
        self.assertIn("2023-02-03,Yes,95000.0", normalized_csv)
        self.assertNotIn("2023-13-45", normalized_csv)

        unnormalized_response = self.client.post(
            "/api/analyze",
            json={"session_id": self.session_id, "normalize_booleans": False},
        )
        self.assertEqual(unnormalized_response.status_code, 200)
        unnormalized_csv_response = self.client.get(
            f"/download/{unnormalized_response.get_json()['clean_csv']}"
        )
        unnormalized_csv = unnormalized_csv_response.get_data(as_text=True).lstrip("\ufeff")
        unnormalized_csv_response.close()
        self.assertIn("2023-01-01,yes,85000.0", unnormalized_csv)

    def test_homepage_serves_the_modern_interface(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Make your data", response.data)
        self.assertIn(b"/static/app.css", response.data)
        self.assertIn(b"Explore with sample data", response.data)
        self.assertIn(b'id="normalize-booleans"', response.data)

    def test_base64_upload_transport_accepts_waf_sensitive_text(self):
        content = b"ID,Notes\n1,drop table users\n2,<script>alert(1)</script>\n"
        response = self.client.post(
            "/api/upload-base64",
            json={
                "filename": "waf-sensitive.csv",
                "content": base64.b64encode(content).decode("ascii"),
            },
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        self.assertEqual(response.get_json()["rows"], 2)

    def test_sample_data_is_available_without_retaining_a_dataframe(self):
        response = self.client.post("/api/sample")
        self.assertEqual(response.status_code, 200)
        sample = response.get_json()
        self.session_id = sample["session_id"]

        self.assertEqual(sample["rows"], 10300)
        self.assertNotIn("frame", client_app._datasets[self.session_id])
        analyzed_response = self.client.post("/api/analyze", json={"session_id": self.session_id})
        self.assertEqual(analyzed_response.status_code, 200)
        analyzed = analyzed_response.get_json()
        self.assertEqual(analyzed["clean_rows"], 10000)
        quality_by_column = {
            (issue["kind"], issue["column"]): issue
            for issue in analyzed["quality_issues"]
        }
        self.assertEqual(quality_by_column[("negative", "Quantity")]["count"], 190)
        unique_by_column = {item["column"]: item["unique"] for item in analyzed["insights"]}
        self.assertEqual(unique_by_column["Product_Category"], 5)
        self.assertEqual(unique_by_column["Sales_Rep"], 10)

    def test_rejects_non_local_hosts_and_cross_origin_requests(self):
        bad_host = self.client.get("/", headers={"Host": "attacker.example"})
        bad_origin = self.client.get("/", headers={"Origin": "http://attacker.example"})

        self.assertEqual(bad_host.status_code, 403)
        self.assertEqual(bad_origin.status_code, 403)

    def test_idle_sessions_expire_and_remove_their_uploaded_file(self):
        response = self.client.post(
            "/api/upload",
            data={"file": (io.BytesIO(b"Name\nAda\n"), "small.csv")},
            content_type="multipart/form-data",
        )
        self.session_id = response.get_json()["session_id"]
        upload_path = Path(client_app._datasets[self.session_id]["path"])
        client_app._datasets[self.session_id]["last_access"] = 100

        client_app._expire_sessions(now=100 + client_app.SESSION_TTL_SECONDS + 1)

        self.assertNotIn(self.session_id, client_app._datasets)
        self.assertFalse(upload_path.exists())
        expired_response = self.client.post("/api/analyze", json={"session_id": self.session_id})
        self.assertEqual(expired_response.status_code, 404)

    def test_output_cleanup_only_removes_expired_generated_reports(self):
        output_dir = Path(self.temp_dir.name)
        generated = output_dir / "Cleaned_expired.csv"
        user_file = output_dir / "notes.txt"
        generated.write_text("generated", encoding="utf-8")
        user_file.write_text("user file", encoding="utf-8")
        old_time = time.time() - (client_app.OUTPUT_TTL_DAYS + 1) * 24 * 60 * 60
        os.utime(generated, (old_time, old_time))
        os.utime(user_file, (old_time, old_time))

        client_app._cleanup_old_outputs(now=time.time())

        self.assertFalse(generated.exists())
        self.assertTrue(user_file.exists())

    def test_computed_revenue_requires_explicit_opt_in(self):
        response = self.client.post(
            "/api/upload",
            data={
                "file": (
                    io.BytesIO(
                        b"Quantity,Unit Price,Discount,Department\n"
                        b"2,10,10%,A\n"
                        b"3,20,0%,B\n"
                    ),
                    "revenue.csv",
                )
            },
            content_type="multipart/form-data",
        )
        self.session_id = response.get_json()["session_id"]

        without_opt_in = self.client.post("/api/analyze", json={"session_id": self.session_id})
        with_opt_in = self.client.post(
            "/api/analyze",
            json={"session_id": self.session_id, "derive_revenue": True},
        )

        self.assertNotIn("Computed Revenue", {
            item["column"] for item in without_opt_in.get_json()["insights"]
        })
        computed_result = with_opt_in.get_json()
        self.assertEqual(computed_result["columns"], 5)
        self.assertEqual(computed_result["preview"][0]["Computed Revenue"], 18.0)


if __name__ == "__main__":
    unittest.main()
