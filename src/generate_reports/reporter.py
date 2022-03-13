# PYTHON SCRIPT
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""
import os

from src.generate_reports.ppt_report_generator import SideCrashPPTReportGenerator
from src.generate_reports.threed_data_reporter import ThreeDDataReporter

class Reporter():
    """
        __init__ _summary_

        _extended_summary_

        Args:
            windows (_type_): _description_
            general_input (_type_): _description_
            metadb_2d_input (_type_): _description_
            metadb_3d_input (_type_): _description_
            config_folder (_type_): _description_
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
        self.get_reporting_folders()
        self.make_reporting_folders()

    def make_reporting_folders(self):
        """
        make_reporting_folders _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """

        reporting_folders = [self.twod_images_report_folder, self.threed_images_report_folder, self.threed_videos_report_folder, self.excel_bom_report_folder,self.ppt_report_folder,self.log_report_folder]
        for folder_path in reporting_folders:
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
        return 0

    def get_reporting_folders(self):
        """
        get_reporting_folders _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """

        self.twod_images_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.report_directory).replace("/","",1),"2d-data-images").replace("\\",os.sep)
        self.threed_images_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.report_directory).replace("/","",1),"3d-data-images").replace("\\",os.sep)
        self.threed_videos_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.report_directory).replace("/","",1),"3d-data-videos").replace("\\",os.sep)
        self.excel_bom_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.report_directory).replace("/","",1),"excel-bom").replace("\\",os.sep)
        self.ppt_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.report_directory).replace("/","",1),"reports").replace("\\",os.sep)
        self.log_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.log_file_directory).replace("/","",1)).replace("\\",os.sep)

        return 0

    def run_process(self):
        """
        run_process _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """

        self.thesis_report_generation()
        #self.threed_data_reporting()

        return 0

    def thesis_report_generation(self):
        """
        thesis_report_generation [summary]

        [extended_summary]

        Returns:
            [type]: [description]
        """
        side_crash_report_ppt = SideCrashPPTReportGenerator(self.windows,
                                                   self.general_input,
                                                   self.metadb_2d_input,
                                                   self.metadb_3d_input,
                                                   self.template_file,
                                                   self.twod_images_report_folder,
                                                   self.threed_images_report_folder,
                                                   self.ppt_report_folder)
        side_crash_report_ppt.generate_ppt()

        return 0


    def threed_data_reporting(self):
        """
        threed_data_reporting [summary]

        [extended_summary]

        Returns:
            [type]: [description]
        """
        threed_data_report = ThreeDDataReporter(self.general_input.threed_window_name,
                                            self.metadb_3d_input,
                                            self.threed_images_report_folder,
                                            self.threed_videos_report_folder,
                                            self.excel_bom_report_folder)
        threed_data_report.run_process()

        return 0
