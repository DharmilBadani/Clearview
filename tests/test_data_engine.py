import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from data_engine import (
    _parse_number,
    add_computed_revenue,
    clean_table,
    detect_column_types,
    filter_table,
    load_table,
    suggest_revenue_columns,
)


class DataEngineTests(unittest.TestCase):
    def test_detects_new_columns_without_a_fixed_business_schema(self):
        frame = pd.DataFrame(
            {
                "order_id": ["A-01", "A-02", "A-03"],
                "Start Date": ["2024-01-01", "2024-02-01", "2024-03-01"],
                "Amount": ["$1,200.50", "$25.00", "$8.75"],
                "Region": ["West", "East", "West"],
                "Client Notes": ["urgent", "follow-up requested", "standard"],
                "Client Email": ["person1@example.org", "person2@example.org", "person3@example.org"],
            }
        )

        types = detect_column_types(frame)

        self.assertEqual(types["order_id"], "IDENTIFIER")
        self.assertEqual(types["Start Date"], "DATE")
        self.assertEqual(types["Amount"], "MONEY")
        self.assertEqual(types["Region"], "CATEGORY")
        self.assertEqual(types["Client Notes"], "TEXT")
        self.assertEqual(types["Client Email"], "TEXT")

    def test_cleaning_preserves_partial_rows_and_audits_removed_records(self):
        raw = pd.DataFrame(
            [
                ["0001", "$1,200.50", "10%", "2024-01-02", "East", "=SUM(1,1)"],
                ["0002", "(€2,50)", "0.25", "not-a-date", "", "keep partial row"],
                ["0001", "$1,200.50", "10%", "2024-01-02", "East", "=SUM(1,1)"],
                [None, None, None, None, None, None],
            ],
            columns=["order_id", "Amount", "Share", "Date", "Region", "Notes"],
        )
        types = {
            "order_id": "IDENTIFIER",
            "Amount": "MONEY",
            "Share": "PERCENT",
            "Date": "DATE",
            "Region": "CATEGORY",
            "Notes": "TEXT",
        }

        clean, report = clean_table(raw, types)

        self.assertEqual(len(clean), 2)
        self.assertEqual(report["blank_rows_removed"], 1)
        self.assertEqual(report["duplicate_rows_removed"], 1)
        self.assertEqual(report["conversion_issues"], {"Date": 1})
        self.assertEqual(clean.loc[0, "order_id"], "0001")
        self.assertEqual(clean.loc[1, "Amount"], -2.5)
        self.assertEqual(clean.loc[0, "Share"], 0.1)
        self.assertTrue(pd.isna(clean.loc[1, "Date"]))

    def test_loads_delimited_files_and_makes_duplicate_headers_unique(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "input.csv"
            path.write_text("Name;Name;Amount\nAda;Lovelace;1,25\n", encoding="utf-8")

            frame, sheets = load_table(path)

        self.assertEqual(sheets, [])
        self.assertEqual(list(frame.columns), ["Name", "Name_2", "Amount"])
        self.assertEqual(frame.iloc[0, 0], "Ada")

    def test_messy_employee_workbook_infers_types_and_audits_failed_conversions(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "messy_test_workbook.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Employees"
            sheet.append(["Name", "Name", "Salary", "Joined", "Active"])
            sheet.append(["Ada", "Lovelace", "85k", "2023-01-15", "y"])
            sheet.append(["Alan", "Turing", "TBD", "2023-13-45", "No"])
            sheet.append(["Grace", "Hopper", "$95,000", "2023-02-02", "TRUE"])
            sheet.append(["Katherine", "Johnson", "\u20b910k", "TBD", "0"])
            workbook.save(path)

            frame, sheets = load_table(path, "Employees")
            types = detect_column_types(frame)
            clean, report = clean_table(frame, types)

        self.assertIn("Employees", sheets)
        self.assertEqual(list(frame.columns), ["Name", "Name_2", "Salary", "Joined", "Active"])
        self.assertEqual(types["Salary"], "MONEY")
        self.assertEqual(types["Joined"], "DATE")
        self.assertEqual(types["Active"], "BOOLEAN")
        self.assertEqual(clean.loc[0, "Salary"], 85000)
        self.assertTrue(pd.isna(clean.loc[1, "Salary"]))
        self.assertEqual(clean["Active"].tolist(), ["Yes", "No", "Yes", "No"])
        self.assertEqual(clean.loc[0, "Joined"].strftime("%Y-%m-%d"), "2023-01-15")
        self.assertGreater(sum(report["conversion_issues"].values()), 0)
        failures = report["conversion_failures"]
        invalid_date = next(item for item in failures if item["value"] == "2023-13-45")
        self.assertTrue(pd.isna(clean.loc[1, "Salary"]))
        self.assertEqual((invalid_date["column"], invalid_date["row"]), ("Joined", 3))
        self.assertEqual(_parse_number("(₹1,250.50)"), -1250.5)
        self.assertEqual(_parse_number("€1.5m"), 1_500_000)

    def test_inventory_fixture_removes_repeated_headers_and_empty_columns(self):
        path = Path(__file__).parent / "fixtures" / "messy_test_inventory_semicolon.csv"
        raw, _ = load_table(path)
        types = detect_column_types(raw)

        clean, report = clean_table(raw, types)

        self.assertEqual(len(clean), 400)
        self.assertNotIn("Column_13", clean.columns)
        self.assertEqual(types["Active?"], "BOOLEAN")
        self.assertEqual(set(clean["Active?"].dropna()), {"Yes", "No"})
        self.assertEqual(report["repeated_header_rows_removed"], 1)
        self.assertEqual(report["empty_columns_removed"], ["Column_13"])
        self.assertEqual(report["conversion_issues"]["Last Restock"], 45)
        quality_by_kind = {issue["kind"]: issue for issue in report["quality_issues"]}
        self.assertEqual(quality_by_kind["repeated_header_rows_removed"]["count"], 1)
        self.assertEqual(quality_by_kind["empty_column_removed"]["column"], "Column_13")

    def test_survey_fixture_detects_types_and_flags_invalid_values(self):
        path = Path(__file__).parent / "fixtures" / "messy_test_survey.txt"
        raw, _ = load_table(path)
        types = detect_column_types(raw)

        clean, report = clean_table(raw, types)

        self.assertEqual(types["Age"], "NUMBER")
        self.assertEqual(types["Income"], "MONEY")
        self.assertEqual(types["Submitted"], "DATE")
        self.assertEqual(len(clean), 150)
        failures = {(item["column"], item["value"], item["row"]) for item in report["conversion_failures"]}
        self.assertIn(("Submitted", "yesterday", 5), failures)
        self.assertNotIn(("Age", "forty", 8), failures)
        self.assertNotIn(("Income", "unknown", 10), failures)
        self.assertNotIn("Income", report["conversion_issues"])

    def test_money_parsing_supports_localized_numbers_and_lakh_suffix(self):
        self.assertEqual(_parse_number("\u20b91.234,56"), 1234.56)
        self.assertEqual(_parse_number("$1,234.56"), 1234.56)
        self.assertEqual(_parse_number("₹1.2L"), 120000)
        self.assertEqual(_parse_number("85k"), 85000)
        self.assertEqual(_parse_number("2m"), 2_000_000)

    def test_number_parser_does_not_drop_digits_or_unicode_minus_signs(self):
        self.assertEqual(_parse_number("1.5m"), 1_500_000)
        self.assertEqual(_parse_number("$1,200"), 1200)
        self.assertEqual(_parse_number("-42"), -42)

    def test_filters_rows_and_custom_validation_rules_are_reported(self):
        raw = pd.DataFrame({
            "Region": ["West", "East", "West"],
            "Amount": ["120", "80", "40"],
            "Status": ["Open", "Closed", "Open"],
        })
        types = {"Region": "CATEGORY", "Amount": "NUMBER", "Status": "CATEGORY"}
        filtered = filter_table(raw, types, [{"column": "Region", "operator": "equals", "value": "west"}])
        self.assertEqual(len(filtered), 2)

        clean, report = clean_table(
            raw,
            types,
            validation_rules=[
                {"column": "Amount", "operator": "min", "value": "100"},
                {"column": "Status", "operator": "one_of", "value": "Open,Pending"},
            ],
        )
        custom = {(issue["column"], issue["count"]) for issue in report["quality_issues"] if issue["kind"] == "custom_validation"}
        self.assertIn(("Amount", 2), custom)
        self.assertIn(("Status", 1), custom)

    def test_invalid_date_and_number_values_are_blank_and_keep_source_rows(self):
        raw = pd.DataFrame(
            {
                "Submitted": ["2024-01-01", "yesterday", "00/00/0000", "31/02/2024", "TBD"],
                "Income": ["10", "unknown", "forty", "20", "30"],
            }
        )

        clean, report = clean_table(raw, {"Submitted": "DATE", "Income": "NUMBER"})

        self.assertEqual(len(clean), len(raw))
        self.assertTrue(clean["Submitted"].iloc[1:].isna().all())
        self.assertTrue(pd.isna(clean["Income"].iloc[1]))
        self.assertEqual(clean["Income"].iloc[2], 40)
        failures = {(item["column"], item["value"], item["row"]) for item in report["conversion_failures"]}
        self.assertEqual(
            failures,
            {
                ("Submitted", "yesterday", 3),
                ("Submitted", "31/02/2024", 5),
            },
        )

    def test_range_flags_preserve_out_of_range_values_and_title_case_categories(self):
        raw = pd.DataFrame(
            {
                "Tax Rate": ["50%", "150%", "75%", None, None, None],
                "Quantity": ["1", "-2", "3", "4", "5", "6"],
                "Satisfaction": ["1", "2", "3", "4", "5", "90%"],
                "Region": ["rajkot", "RAJKOT", "rajkot", "RAJKOT", "new  delhi", "new delhi"],
            }
        )
        types = {
            "Tax Rate": "PERCENT",
            "Quantity": "INTEGER",
            "Satisfaction": "INTEGER",
            "Region": "CATEGORY",
        }

        clean, report = clean_table(raw, types)

        issues = {(issue["kind"], issue["column"]): issue for issue in report["quality_issues"]}
        self.assertEqual(len(clean), len(raw))
        self.assertEqual(clean.loc[1, "Quantity"], -2)
        self.assertEqual(clean.loc[5, "Satisfaction"], 90)
        self.assertEqual(issues[("negative", "Quantity")]["rows"], [3])
        self.assertEqual(issues[("out_of_range", "Tax Rate")]["rows"], [3])
        self.assertEqual(issues[("out_of_range", "Satisfaction")]["rows"], [7])
        self.assertEqual(issues[("outlier", "Satisfaction")]["rows"], [7])
        self.assertEqual(clean["Region"].tolist(), ["Rajkot", "Rajkot", "Rajkot", "Rajkot", "New Delhi", "New Delhi"])

    def test_boolean_normalization_can_be_disabled(self):
        raw = pd.DataFrame({"Active": ["true", "f", "no"]})

        clean, _ = clean_table(raw, {"Active": "BOOLEAN"}, normalize_booleans=False)

        self.assertEqual(clean["Active"].tolist(), ["true", "f", "no"])

    def test_empty_input_is_reported_instead_of_returning_a_success_shaped_result(self):
        with self.assertRaisesRegex(ValueError, "no non-empty records"):
            clean_table(pd.DataFrame({"value": [None, ""]}))

    def test_percentages_integer_issues_and_category_merges_are_audited_once(self):
        raw = pd.DataFrame(
            {
                "Share": ["1%", "0.5%", "25", "0.25"],
                "Count": ["1.5", "not-a-number", "3", "4"],
                "Product_Category": [" Home  Goods ", "home goods", "HOME GOODS", "Office"],
            }
        )
        clean, report = clean_table(
            raw,
            {"Share": "PERCENT", "Count": "INTEGER", "Product_Category": "CATEGORY"},
        )
        without_normalization, _ = clean_table(
            raw,
            {"Share": "PERCENT", "Count": "INTEGER", "Product_Category": "CATEGORY"},
            normalize_categories=False,
        )

        self.assertEqual(clean["Share"].tolist(), [0.01, 0.005, 0.25, 0.25])
        self.assertEqual(report["conversion_issues"], {"Count": 2})
        self.assertEqual(
            [issue["row"] for issue in report["conversion_failures"]],
            [2, 3],
        )
        self.assertEqual(clean["Product_Category"].nunique(), 2)
        self.assertEqual(len(report["category_merges"]["Product_Category"]), 1)
        self.assertEqual(without_normalization["Product_Category"].nunique(), 4)

    def test_date_order_flags_ambiguous_values_and_accepts_explicit_choice(self):
        raw = pd.DataFrame({"Order Date": ["03/04/2024", "05/06/2024"]})

        clean, report = clean_table(raw, {"Order Date": "DATE"})
        day_first, day_report = clean_table(
            raw,
            {"Order Date": "DATE"},
            date_orders={"Order Date": "day_first"},
        )

        self.assertTrue(report["date_order_ambiguous"]["Order Date"])
        self.assertEqual(clean.loc[0, "Order Date"].month, 3)
        self.assertEqual(day_first.loc[0, "Order Date"].month, 4)
        self.assertFalse(day_report["conversion_issues"])

    def test_quality_flags_are_advisory_and_include_source_rows(self):
        raw = pd.DataFrame(
            {
                "Quantity": ["1", "-2", "3", "4", "5"],
                "Client Email": ["valid@example.org", "invalid", "other@example.org", None, "more@example.org"],
            }
        )

        clean, report = clean_table(raw, {"Quantity": "INTEGER", "Client Email": "TEXT"})
        by_kind = {issue["kind"]: issue for issue in report["quality_issues"]}

        self.assertEqual(len(clean), len(raw))
        self.assertEqual(by_kind["negative"]["count"], 1)
        self.assertEqual(by_kind["negative"]["rows"], [3])
        self.assertEqual(by_kind["invalid_email"]["rows"], [3])
        self.assertEqual(by_kind["missing"]["count"], 1)

    def test_computed_revenue_is_suggested_and_added_only_when_requested(self):
        frame = pd.DataFrame(
            {
                "Quantity": [2, 3],
                "Unit_Price": [10.0, 20.0],
                "Discount": [0.1, 0.0],
            }
        )
        types = {"Quantity": "INTEGER", "Unit_Price": "MONEY", "Discount": "PERCENT"}
        suggestion = suggest_revenue_columns(frame, types)

        self.assertEqual(suggestion, {
            "quantity": "Quantity",
            "price": "Unit_Price",
            "discount": "Discount",
        })
        computed = add_computed_revenue(frame, suggestion)
        self.assertEqual(computed["Computed Revenue"].tolist(), [18.0, 60.0])
        self.assertNotIn("Computed Revenue", frame.columns)


if __name__ == "__main__":
    unittest.main()
