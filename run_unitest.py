# PYTHON script

import sys

# NOTE: ONLY FOR DEBUGGING
DEL_ITEMS = [
    "tests",
    "test.run"
    ]

for item in DEL_ITEMS:
   if item in sys.modules:
       del sys.modules[item]

from tests.run import main
