# PYTHON script
"""
##################################################
#      Copyright BETA CAE Systems USA Inc.,      #
#      2020 All Rights Reserved                  #
#      UNPUBLISHED, LICENSED SOFTWARE.           #
##################################################


Gui for running the side crash report generation.

Developer   : Naresh Medipalli and Shiva Sai
Date        : Dec 31, 2021

"""

import sys

# NOTE: ONLY FOR DEBUGGING
DEL_ITEMS = [
    "src",
    "src.run",
    "src.metadb_info",
    "src.user_input",
    "src.meta_utilities",
    "src.general_utilities",
    "src.generate_reports.side_crash_ppt_report",
    "src.generate_reports.reporter",
    "src.generate_reports.report_slides.title_slide",
    "src.generate_reports.report_slides.cae_quality_slide",
    "src.generate_reports.report_slides.executive_slide",
    "src.generate_reports.report_slides.cbu_and_barrier_position_slide",
    "src.generate_reports.report_slides.biw_kinematics_slide",
    "src.generate_reports.report_slides.biw_cbu_deformation_slide",
    "src.generate_reports.report_slides.bom_f21_upb_slide",
    "src.generate_reports.report_slides.biw_stiff_ring_deformation_slide",
    ]

for item in DEL_ITEMS:
   if item in sys.modules:
       del sys.modules[item]

from src.run import main
