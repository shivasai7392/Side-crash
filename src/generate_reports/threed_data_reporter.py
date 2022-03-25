# PYTHON script
"""
This Script is used to generate the threeDData Reporting.
"""

import logging

from meta import utils

from src.generate_reports.bom_excel_generator import ExcelBomGeneration


class ThreeDDataReporter():
    """
    This Class is used to Automate the threed Data Reporting to generate the Excel Bill Of Material Excel Files.

    Args:
        metadb_3d_input (Meta3DInfo): Meta3DInfo class object.
        threed_images_report_folder (str): threed images reporting folder
        thrred_videos_report_folder (str): threed videos reporting folder
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
            Int: 0 Always
        """
        # Maximizing the threed window and calling the ExcelBomGeneration class and executing the excel_bom_generation function
        utils.MetaCommand('window maximize {}'.format(self.threed_window_name))
        self.logger.info("")
        self.logger.info("--- 3D MODEL BOM GENERATOR")
        excel_bom_report = ExcelBomGeneration(self.metadb_3d_input, self.excel_bom_report_folder)
        excel_bom_report.excel_bom_generation()

        return 0
