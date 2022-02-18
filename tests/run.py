#PYTHON SCRIPT

import unittest

from .unit_test import user_input_test

def main(*args):
    suite = unittest.TestLoader().loadTestsFromModule(user_input_test)
    unittest.TextTestRunner(verbosity=2,buffer=True).run(suite)
    return 0