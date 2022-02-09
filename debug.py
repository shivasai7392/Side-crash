# PYTHON script
"""
##################################################
#      Copyright BETA CAE Systems USA Inc.,      #
#      2020 All Rights Reserved                  #
#      UNPUBLISHED, LICENSED SOFTWARE.           #
##################################################


Debug script to run the side crash report generation.


Developer   : Naresh Medipalli
Date        : Dec 31, 2021

"""

import os
import sys

from ansa import base


def quick_check():
    """ Quickly test a check
    """

    CustomChecksGUI()

    return 0



def main():
    """ Calling the Gui functionality.
    """

    side_crash_gui()


if __name__ == "__main__":
    main()
