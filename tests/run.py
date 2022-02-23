#PYTHON SCRIPT

import unittest

from .unit_test import user_input_test
from .unit_test.generate_report_tests import side_crash_report_test

def main():
    suite = unittest.TestLoader().loadTestsFromModule(user_input_test)
    unittest.TextTestRunner(verbosity=2,buffer=True).run(suite)

    side_crash_report_suite = unittest.TestLoader().loadTestsFromModule(side_crash_report_test)
    unittest.TextTestRunner(verbosity=2, buffer=True).run(side_crash_report_suite)

    return 0