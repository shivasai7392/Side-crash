# PYTHON SCRIPT
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""
import os
import datetime

from meta import utils

from src.generate_reports.side_crash_ppt_report import SideCrashPPTReport
from src.generate_reports.threed_data_report import ThreeDDataReport
from src.logger import SideCrashLogger

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
    def __init__(self,windows,general_input,metadb_2d_input,metadb_3d_input,config_folder) -> None:

        self.windows = windows
        self.general_input = general_input
        self.metadb_2d_input = metadb_2d_input
        self.metadb_3d_input = metadb_3d_input
        self.config_folder = config_folder
        self.template_file = os.path.join(self.config_folder,"res",self.general_input.source_template_file_directory.replace("/","",1),self.general_input.source_template_file_name).replace("\\",os.sep)
        self.get_reporting_folders()
        self.make_reporting_folders()

    @SideCrashLogger.excel_resource_log_decorator(Description = "MAIN METHOD")
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

    @SideCrashLogger.excel_resource_log_decorator(Description = "MAIN METHOD")
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

    @SideCrashLogger.excel_resource_log_decorator(Description = "MAIN METHOD")
    def run_process(self):
        """
        run_process _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """

        self.twod_data_reporting()
        self.threed_data_reporting()
        self.thesis_report_generation()
        self.log_report_generation()

        return 0

    def thesis_report_generation(self):
        """
        thesis_report_generation [summary]

        [extended_summary]

        Returns:
            [type]: [description]
        """
        side_crash_report_ppt = SideCrashPPTReport(self.windows,
                                                   self.general_input,
                                                   self.metadb_2d_input,
                                                   self.metadb_3d_input,
                                                   self.template_file,
                                                   self.twod_images_report_folder,
                                                   self.threed_images_report_folder,
                                                   self.ppt_report_folder
                                                   )
        side_crash_report_ppt.generate_ppt()

        return 0

    @SideCrashLogger.excel_resource_log_decorator(Description = "MAIN METHOD")
    def log_report_generation(self):
        """
        log_report_generation _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """
        current_datetime = datetime.datetime.now()
        file_path = os.path.join(self.log_report_folder,"2TN_MP_log_{}.xlsx".format(current_datetime.strftime('%Y-%d-%m-%H-%M-%S')))
        SideCrashLogger.save_workbook(file_path)

        return 0

    def threed_data_reporting(self):
        """
        threed_data_reporting [summary]

        [extended_summary]

        Returns:
            [type]: [description]
        """
        threed_data_report = ThreeDDataReport(self.general_input.threed_window_name,
                                            self.metadb_3d_input,
                                            self.threed_images_report_folder,
                                            self.threed_videos_report_folder,
                                            self.excel_bom_report_folder)
        threed_data_report.run_process()

        return 0

    def twod_data_reporting(self):
        """
        twod_data_reporting [summary]

        [extended_summary]

        Returns:
            [type]: [description]
        """
        from PIL import ImageFile
        ImageFile.LOAD_TRUNCATED_IMAGES = True

        window_2d_objects = self.metadb_2d_input.window_objects
        for window in window_2d_objects:
            window_name = window.name
            window_layout = window.meta_obj.get_plot_layout()
            plot = window.plot
            curve = plot.curve
            utils.MetaCommand('window active "{}"'.format(window_name))
            utils.MetaCommand('window maximize "{}"'.format(window_name))
            utils.MetaCommand('xyplot plotdeactive "{}" all'.format(window_name))
            curve.meta_obj.show()
            utils.MetaCommand('xyplot plotactive "{}" {}'.format(window_name, plot.id))
            utils.MetaCommand('xyplot curve visible and "{}" selected'.format(window_name))
            utils.MetaCommand('xyplot rlayout "{}" 1'.format(window_name))
            image_path = os.path.join(self.twod_images_report_folder,window_name+"_"+curve.name.lower()+".png")

            utils.MetaCommand('write png "{}"'.format(image_path))

            utils.MetaCommand('xyplot rlayout "{}" {}'.format(window_name,window_layout))

        return 0
