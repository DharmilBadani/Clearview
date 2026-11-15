import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from data_engine import clean_table, load_table
from report_builder import build_excel_report


class ReportBuilderTests(unittest.TestCase):
    def test_builds_portable_dynamic_workbook_with_safe_values_and_charts(self):
        raw = pd.DataFrame(
            {
                "Date": ["2024-01-03", "2024-02-02", "2024-02-15", "2024-03-10"],
                "Department": ["Sales", "Support", "Sales", "Support"],
                "Gross Amount": ["$100", "$80", "$120", "$70"],
                "Comment": ["=1+1", "regular", "regular", "regular"],
            }
        )
        types = {
            "Date": "DATE",
            "Department": "CATEGORY",
            "Gross Amount": "MONEY",
            "Comment": "TEXT",
        }
        clean, report = clean_table(raw, types)

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "dynamic_report.xlsx"
            build_excel_report(
                raw, clean, types, report, path,
                measure="Gross Amount",
                dimensions=["Department"],
            )
            workbook = load_workbook(path)
            round_trip, sheets = load_table(path, "Cleaned Data")

        self.assertIn("Column Insights", workbook.sheetnames)
        self.assertIn("Cleaned Data", workbook.sheetnames)
        self.assertIn("Category Insights", workbook.sheetnames)
        self.assertIn("Measure Insights", workbook.sheetnames)
        self.assertIn("Data Quality", workbook.sheetnames)
        self.assertIn("Charts", workbook.sheetnames)
        self.assertIn("Chart Data", workbook.sheetnames)
        self.assertEqual(workbook["Overview"]["B10"].value, "Gross Amount")
        self.assertEqual(workbook["Overview"]["B11"].value, "Department")
        self.assertEqual(workbook["Measure Insights"]["C6"].value, 220)
        self.assertEqual(workbook["Cleaned Data"]["D2"].value, "=1+1")
        self.assertEqual(workbook["Cleaned Data"]["D2"].data_type, "s")
        self.assertTrue(workbook["Cleaned Data"]["D2"].quotePrefix)
        self.assertGreater(len(workbook["Charts"]._charts), 0)
        self.assertEqual(workbook["Cleaned Data"].max_row, 5)
        self.assertIn("Overview", sheets)
        self.assertEqual(len(round_trip), len(clean))

    def test_numeric_only_files_still_receive_a_distribution_chart(self):
        raw = pd.DataFrame({"Reading": ["10", "20", "30", "40"]})
        types = {"Reading": "NUMBER"}
        clean, report = clean_table(raw, types)

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "numeric_report.xlsx"
            build_excel_report(raw, clean, types, report, path)
            workbook = load_workbook(path)

        self.assertEqual(len(workbook["Charts"]._charts), 1)
        self.assertEqual(workbook["Charts"]._charts[0].title.tx.rich.p[0].r[0].t, "Distribution of Reading")

    def test_unit_price_is_averaged_and_data_quality_sheet_preserves_flags(self):
        raw = pd.DataFrame(
            {
                "Department": ["A", "A", "B", "B"],
                "Unit Price": ["10", "20", "30", "50"],
                "Quantity": ["1", "-2", "3", "4"],
            }
        )
        types = {"Department": "CATEGORY", "Unit Price": "MONEY", "Quantity": "INTEGER"}
        clean, report = clean_table(raw, types)

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "quality_report.xlsx"
            build_excel_report(raw, clean, types, report, path, measure="Unit Price", dimensions=["Department"])
            workbook = load_workbook(path)

        self.assertEqual(workbook["Measure Insights"]["C6"].value, 40)
        self.assertEqual(workbook["Measure Insights"]["C7"].value, 15)
        self.assertEqual(workbook["Data Quality"]["A5"].value, "negative")
        self.assertEqual(workbook["Data Quality"]["D5"].value, "3")


if __name__ == "__main__":
    unittest.main()
