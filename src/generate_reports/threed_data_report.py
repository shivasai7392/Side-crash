# PYTHON script
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""

import os

from meta import utils

from src.meta_utilities import visualize_3d_critical_section,annotation
from src.generate_reports.excel_generator import ExcelBomGeneration


class ThreeDDataReport():
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
                excel_bom_report_folder,
                logger) -> None:

        self.threed_window_name = threed_window_name
        self.metadb_3d_input = metadb_3d_input
        self.critical_sections = metadb_3d_input.critical_sections
        self.threed_images_report_folder = threed_images_report_folder
        self.threed_videos_report_folder = threed_videos_report_folder
        self.excel_bom_report_folder = excel_bom_report_folder
        self.logger = logger

    def run_process(self):
        """
        run_process _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """
        utils.MetaCommand('window maximize {}'.format(self.threed_window_name))
        self.logger.log.info("--- 3D MODEL IMAGE GENERATOR")
        self.logger.log.info("")
        self.get_initial_state_images()
        self.get_peak_state_images()
        excel_bom_report = ExcelBomGeneration(self.metadb_3d_input, self.excel_bom_report_folder)
        self.logger.log.info("--- 3D MODEL BOM GENERATOR")
        excel_bom_report = ExcelBomGeneration(self.metadb_3d_input, self.excel_bom_report_folder,self.logger)
        excel_bom_report.excel_bom_generation()

        return 0

    def get_initial_state_images(self):
        """
        get_initial_state_images _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """
        from PIL import Image

        for section,value in self.critical_sections.items():
            if "hes" in value.keys() and value["hes"] != "null":
                self.logger.log.info("SOURCE WINDOW : {} ".format(self.threed_window_name))
                self.logger.log.info("SOURCE MODEL : 0")
                self.logger.log.info("STATE : ORIGINAL STATE")
                image_path = os.path.join(self.threed_images_report_folder,self.threed_window_name+"_"+section.lower()+"_Image"+".png")
                titled_image_path = os.path.join(self.threed_images_report_folder,self.threed_window_name+"_"+section.lower()+"_Titled_Image"+".png")
                self.logger.log.info("PID NAME SHOW FILTER : {} ".format(value["hes"] if "hes" in value.keys() else "null"))
                self.logger.log.info("ADDITIONAL PID'S SHOWN : {} ".format(value["hes_exceptions"] if "hes_exceptions" in value.keys() else "null"))
                self.logger.log.info("PID NAME ERASE FILTER : {} ".format(value["hes_exceptions"] if "hes_exceptions" in value.keys() else "null"))
                self.logger.log.info("PID'S TO ERASE : {} ".format(value["erase_pids"] if "erase_pids" in value.keys() else "null"))
                self.logger.log.info("ERASE BOX : {} ".format(value["erase_box"] if "erase_box" in value.keys() else "null"))
                self.logger.log.info("IMAGE VIEW : {} ".format(value["view"] if "view" in value.keys() else "null"))
                self.logger.log.info("TRANSPARENCY LEVEL : 50" )
                self.logger.log.info("TRANSPARENT PID'S : {} ".format(value["transparent_pids"] if "transparent_pids" in value.keys() else "null"))
                self.logger.log.info("COMP NAME : {} ".format(value["name"] if "name" in value.keys() else "null"))

                visualize_3d_critical_section(value)
                utils.MetaCommand('write png "{}"'.format(image_path))
                utils.MetaCommand('options title on')
                utils.MetaCommand('write png "{}"'.format(titled_image_path))

                image = Image.open(image_path)
                width,height = image.size
                self.logger.log.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(width,height))
                self.logger.log.info("OUTPUT MODEL IMAGES :")
                self.logger.log.info("{}".format(image_path))
                self.logger.log.info("{}".format(titled_image_path))
                self.logger.log.info("")
                self.logger.log.info("")

        return 0

    def get_peak_state_images(self):
        """
        get_initial_state_images _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """
        from PIL import Image

        utils.MetaCommand('0:options state variable "serial=1"')
        for section,value in self.critical_sections.items():
            if "hes" in value.keys() and value["hes"] != "null":
                self.logger.log.info("SOURCE WINDOW : {} ".format(self.threed_window_name))
                self.logger.log.info("SOURCE MODEL : 0")
                self.logger.log.info("STATE : PEAK STATE")
                image_path = os.path.join(self.threed_images_report_folder,self.threed_window_name+"_"+section.lower()+"_Image"+".png")
                spotweld_failure_image_path = os.path.join(self.threed_images_report_folder,self.threed_window_name+"_"+section.lower()+"_SpotWeld_Failure_Image"+".png")
                contour_image_path = os.path.join(self.threed_images_report_folder,self.threed_window_name+"_"+section.lower()+"_Contour_Image"+".png")
                contour_with_out_deform_image_path = os.path.join(self.threed_images_report_folder,self.threed_window_name+"_"+section.lower()+"_Contour_Without_Deformation_Image"+".png")
                model_color_image_path =  os.path.join(self.threed_images_report_folder,self.threed_window_name+"_"+section .lower()+"_Model_Color_Image"+".png")
                self.logger.log.info("PID NAME SHOW FILTER : {} ".format(value["hes"] if "hes" in value.keys() else "null"))
                self.logger.log.info("ADDITIONAL PID'S SHOWN : {} ".format(value["hes_exceptions"] if "hes_exceptions" in value.keys() else "null"))
                self.logger.log.info("PID NAME ERASE FILTER : {} ".format(value["hes_exceptions"] if "hes_exceptions" in value.keys() else "null"))
                self.logger.log.info("PID'S TO ERASE : {} ".format(value["erase_pids"] if "erase_pids" in value.keys() else "null"))
                self.logger.log.info("ERASE BOX : {} ".format(value["erase_box"] if "erase_box" in value.keys() else "null"))
                self.logger.log.info("IMAGE VIEW : {} ".format(value["view"] if "view" in value.keys() else "null"))
                self.logger.log.info("TRANSPARENCY LEVEL : 50" )
                self.logger.log.info("TRANSPARENT PID'S : {} ".format(value["transparent_pids"] if "transparent_pids" in value.keys() else "null"))
                self.logger.log.info("COMP NAME : {} ".format(value["name"] if "name" in value.keys() else "null"))
                visualize_3d_critical_section(value)

                utils.MetaCommand('options fringebar on')
                utils.MetaCommand('grstyle deform on')
                utils.MetaCommand('write png "{}"'.format(contour_image_path))

                utils.MetaCommand('grstyle deform off')
                utils.MetaCommand('write png "{}"'.format(contour_with_out_deform_image_path))

                utils.MetaCommand('grstyle scalarfringe disable')
                utils.MetaCommand('write png "{}"'.format(image_path))

                utils.MetaCommand('grstyle deform on')
                utils.MetaCommand('color pid Gray act')
                utils.MetaCommand('write png "{}"'.format(model_color_image_path))

                annotation()
                utils.MetaCommand('write png "{}"'.format(spotweld_failure_image_path))

                utils.MetaCommand('color pid reset act')

                image = Image.open(image_path)
                width,height = image.size
                self.logger.log.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(width,height))
                self.logger.log.info("OUTPUT MODEL IMAGES :")
                self.logger.log.info("{}".format(image_path))
                self.logger.log.info("{}".format(model_color_image_path))
                self.logger.log.info("{}".format(contour_image_path))
                self.logger.log.info("{}".format(contour_with_out_deform_image_path))
                self.logger.log.info("{}".format(spotweld_failure_image_path))
                self.logger.log.info("")
                self.logger.log.info("")

        return 0
