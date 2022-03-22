# PYTHON script
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""

import os
import logging

from meta import utils

from src.meta_utilities import visualize_3d_critical_section
from src.meta_utilities import visualize_annotation
from src.generate_reports.bom_excel_generator import ExcelBomGeneration


class ThreeDDataReporter():
    """
    __init__ _summary_

    _extended_summary_

    Args:
        metadb_3d_input (_type_): _description_
        threed_images_report_folder (_type_): _description_
        thrred_videos_report_folder (_type_): _description_
    """

    def __init__(self,
                threed_window_name,
                metadb_3d_input,
                threed_images_report_folder,
                threed_videos_report_folder,
                excel_bom_report_folder) -> None:

        self.threed_window_name = threed_window_name
        self.metadb_3d_input = metadb_3d_input
        self.critical_sections = metadb_3d_input.critical_sections
        self.threed_images_report_folder = threed_images_report_folder
        self.threed_videos_report_folder = threed_videos_report_folder
        self.excel_bom_report_folder = excel_bom_report_folder
        self.logger = logging.getLogger("side_crash_logger")

    def run_process(self):
        """
        This method is used to generate the excel BOM files.

        Returns:
            Int: 0 for Success 1 for failure
        """
        # Maximizing the threed window and calling the ExcelBomGeneration class and executing the excel_bom_generation function
        utils.MetaCommand('window maximize {}'.format(self.threed_window_name))
        self.logger.info("--- 3D MODEL BOM GENERATOR")
        excel_bom_report = ExcelBomGeneration(self.metadb_3d_input, self.excel_bom_report_folder)
        excel_bom_report.excel_bom_generation()

        return 0
