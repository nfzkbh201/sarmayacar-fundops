import unittest
from unittest.mock import patch

from quarterly_report import _quarterly_workbook_scan


class QuarterlyScanTests(unittest.TestCase):
    def setUp(self):
        self.session = {}
        self.months = ((2026, 4), (2026, 5), (2026, 6))
        self.patches = [
            patch("quarterly_report.st.session_state", self.session),
            patch("quarterly_report.st.spinner"),
            patch("quarterly_report.parse_monthly_reporting_workbook", return_value=()),
            patch("quarterly_report._commentary_rows_from_workbook", return_value={}),
            patch("quarterly_report.build_monthly_workbook_with_commentary_sheet", return_value=b"commentary"),
            patch("quarterly_report.parse_financial_exhibits_workbook", return_value=({"period": "Jun-26"}, [])),
        ]
        self.mocks = [p.start() for p in self.patches]
        for p in self.patches:
            self.addCleanup(p.stop)

    def scan(self, monthly=b"monthly", financial=b"financial", months=None):
        period = months or self.months
        return _quarterly_workbook_scan(monthly, financial, period, period)

    def test_repeated_handoff_actions_reuse_scan_and_commentary_workbook(self):
        first = self.scan()
        self.assertIs(self.scan(), first)
        self.mocks[2].assert_called_once()
        self.mocks[4].assert_called_once()
        self.mocks[5].assert_called_once()

    def test_changes_to_either_upload_period_or_mapping_refresh_the_scan(self):
        self.scan()
        self.scan(monthly=b"new monthly")
        self.scan(monthly=b"new monthly", financial=b"new financial")
        self.scan(months=((2026, 7), (2026, 8), (2026, 9)))
        with patch("quarterly_report.PORTFOLIO_MAPPING_VERSION", "new-version"):
            self.scan(months=((2026, 7), (2026, 8), (2026, 9)))
        self.assertEqual(self.mocks[2].call_count, 5)
        self.assertEqual(len(self.session), 1)

    def test_removing_financial_upload_does_not_reuse_its_values(self):
        self.scan()
        result = self.scan(financial=None)
        self.assertIsNone(result["financial_exhibits"])
        self.assertEqual(result["financial_warnings"], [])
        self.mocks[5].assert_called_once()

    def test_failed_scan_does_not_replace_successful_inputs(self):
        first = self.scan()
        with patch("quarterly_report.parse_monthly_reporting_workbook", side_effect=ValueError("Invalid workbook")):
            with self.assertRaisesRegex(ValueError, "Invalid workbook"):
                self.scan(monthly=b"broken")
        self.assertIs(self.scan(), first)


if __name__ == "__main__":
    unittest.main()
