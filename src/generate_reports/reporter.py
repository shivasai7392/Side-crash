# PYTHON SCRIPT
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""
import os

from meta import utils

from src.generate_reports.side_crash_ppt_report import SideCrashPPTReport
from src.generate_reports.excel_generator import ExcelBomGeneration

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

        return 0

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

    def threed_data_reporting(self):
        """
        threed_data_reporting [summary]

        [extended_summary]

        Returns:
            [type]: [description]
        """
        _critical_sections_data = self.metadb_3d_input.critical_sections
        # for section,value in critical_sections_data.items():
        #     for key,vvalue in value.items():
        #         if key == "hes":
        excel_bom_report = ExcelBomGeneration(self.metadb_3d_input, self.excel_bom_report_folder)
        excel_bom_report.excel_bom_generation()

        return 0

    def twod_data_reporting(self):
        """
        twod_data_reporting [summary]

        [extended_summary]

        Returns:
            [type]: [description]
        """
        from PIL import Image,ImageFile
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

            if not os.path.exists(os.path.dirname(image_path)):
                os.makedirs(os.path.dirname(image_path))
            utils.MetaCommand('write png "{}"'.format(image_path))

            utils.MetaCommand('xyplot rlayout "{}" {}'.format(window_name,window_layout))

        return 0
