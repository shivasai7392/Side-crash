# PYTHON SCRIPT
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""
import os
import logging
from datetime import datetime

from src.generate_reports.ppt_report_generator import SideCrashPPTReportGenerator
from src.generate_reports.threed_data_reporter import ThreeDDataReporter

class Reporter():
    """
    This Class is used to Automate the generating the thesis report and threed data reporting

        Args:
            windows (object): window object
            general_input (GeneralInfo): GeneralInfo class object.
            metadb_2d_input (Meta2DInfo): Meta2DInfo class object.
            metadb_3d_input (Meta3DInfo): Meta3DInfo class object.
            config_folder (str): folder path of config.
    """
    def __init__(self,
                 windows,
                 general_input,
                 metadb_2d_input,
                 metadb_3d_input,
                 config_folder) -> None:

        self.windows = windows
        self.general_input = general_input
        self.metadb_2d_input = metadb_2d_input
        self.metadb_3d_input = metadb_3d_input
        self.config_folder = config_folder
        self.template_file = os.path.join(self.config_folder,"res",self.general_input.source_template_file_directory.replace("/","",1),self.general_input.source_template_file_name).replace("\\",os.sep)
        self.logger = logging.getLogger("side_crash_logger")
        self.get_reporting_folders()
        self.make_reporting_folders()

    def make_reporting_folders(self):
        """
        This method is used to create the Reporting Folders which we should report

        Returns:
            int: 0 Always
        """
        # all the reporting folders
        reporting_folders = [self.twod_images_report_folder, self.threed_images_report_folder, self.threed_videos_report_folder, self.excel_bom_report_folder,self.ppt_report_folder,self.log_report_folder]
        # Iterating the reporting folders if that folder not present the make the directory of that folder
        for folder_path in reporting_folders:
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
        return 0

    def get_reporting_folders(self):
        """
        This method is used to get the Reporting folders

        Returns:
            Int : 0 Always
        """
        # Path of the 2d,3d images, 3d videos,excel BOM and Reports
        self.twod_images_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.report_directory).replace("/","",1),"2d-data-images").replace("\\",os.sep)
        self.threed_images_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.report_directory).replace("/","",1),"3d-data-images").replace("\\",os.sep)
        self.threed_videos_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.report_directory).replace("/","",1),"3d-data-videos").replace("\\",os.sep)
        self.excel_bom_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.report_directory).replace("/","",1),"excel-bom").replace("\\",os.sep)
        self.ppt_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.report_directory).replace("/","",1),"reports").replace("\\",os.sep)
        self.log_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.log_file_directory).replace("/","",1)).replace("\\",os.sep)

        return 0

    def run_process(self):
        """
        This method is used to Calling the thesis report function and threed data reporting function

        Returns:
            Int: 0 Always
        """

        self.thesis_report_generation()
        self.threed_data_reporting()

        return 0

    def thesis_report_generation(self):
        """
        This method is used to generating the slides of thesis report

        Returns:
            [Int]: 0 Always
        """
        self.logger.info("--- Executive and Thesis Report Generation Started")
        starttime = datetime.now()
        side_crash_report_ppt = SideCrashPPTReportGenerator(self.windows,
                                                   self.general_input,
                                                   self.metadb_2d_input,
                                                   self.metadb_3d_input,
                                                   self.template_file,
                                                   self.twod_images_report_folder,
                                                   self.threed_images_report_folder,
                                                   self.ppt_report_folder)
        side_crash_report_ppt.generate_ppt()
        endtime = datetime.now()
        self.logger.info("--- Executive and Thesis Report Generation Completed")
        self.logger.info("Time Taken : {}".format(endtime - starttime))
        self.logger.info("")

        return 0


    def threed_data_reporting(self):
        """
        This methos is used to generate the threed data reporting

        Returns:
            [Int]: 0 Always
        """
        self.logger.info("--- 3D Data Reporting - Excel BOM Generation for critical part sets is started")
        starttime = datetime.now()
        threed_data_report = ThreeDDataReporter(self.general_input.threed_window_name,
                                            self.metadb_3d_input,
                                            self.threed_images_report_folder,
                                            self.threed_videos_report_folder,
                                            self.excel_bom_report_folder)
        threed_data_report.run_process()
        endtime = datetime.now()
        self.logger.info("--- 3D Data Reporting - Excel BOM Generation for critical part sets is compelted")
        self.logger.info("Time Taken : {}".format(endtime - starttime))
        self.logger.info("")

        return 0
