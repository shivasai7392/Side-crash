"""

    _extended_summary_
"""

import unittest
from src.generate_reports import side_crash_report
class SideCrashReportTests(unittest.TestCase):
    def test_case_closest(self, ):
        list_of_values = [1,4,5,6,7]
        value = 2
        test_case_closest_func = side_crash_report.SideCrashReport.closest(list_of_values, value)
        self.assertEqual(test_case_closest_func, 1)
