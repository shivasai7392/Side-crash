# PYTHON script
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""

import os

from meta import utils

from src.meta_utilities import visualize_3d_critical_section
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
        self.get_initial_state_images()
        self.get_peak_state_images()
        # self.get_spotweld_failure_images()
        excel_bom_report = ExcelBomGeneration(self.metadb_3d_input, self.excel_bom_report_folder)
        excel_bom_report.excel_bom_generation()

        return 0

    def get_initial_state_images(self):
        """
        get_initial_state_images _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """
        import PIL
        self.logger.log.info("--- 3D MODEL IMAGE GENERATOR")
        self.logger.log.info("")
        self.logger.log.info("SOURCE WINDOW:", "Metapost")
        self.logger.log.info("SOURCE MODEL:", "0")
        self.logger.log.info("STATE:", "ORIGINAL STATE")
        for section,value in self.critical_sections.items():
            if "hes" in value.keys() and value["hes"] != "null":
                image_path = os.path.join(self.threed_images_report_folder,self.threed_window_name+"_"+section.lower()+".png")
                self.logger.log.info("PID NAME SHOW FILTER:", )
                self.logger.log.info("ADDITIONAL PID'S SHOWN:", )
                self.logger.log.info("PID NAME ERASE FILTER:", )
                self.logger.log.info("PID'S TO ERASE:", )
                self.logger.log.info("ERASE BOX:", )
                self.logger.log.info("IMAGE VIEW:", )
                self.logger.log.info("TRANSPARENCY LEVEL:", )
                self.logger.log.info("TRANSPARENT PID'S:", )
                self.logger.log.info("COMP NAME:", )

                visualize_3d_critical_section(value)
                utils.MetaCommand('write png "{}"'.format(image_path))

                image = PIL.Image.open(image_path)
                width,height = image.size
                self.logger.log.info("OUTPUT IMAGE SIZE (PIXELS) :", width,height)
                self.logger.log.info("OUTPUT MODEL IMAGES:", image_path)


        return 0

    def get_peak_state_images(self):
        """
        get_initial_state_images _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """
        utils.MetaCommand('0:options state variable "serial=1"')
        for section,value in self.critical_sections.items():
            if "hes" in value.keys() and value["hes"] != "null":
                image_path = os.path.join(self.threed_images_report_folder,self.threed_window_name+"_"+section.lower()+".png")
                contour_image_path = os.path.join(self.threed_images_report_folder,self.threed_window_name+"_"+section.lower()+"_contour"+".png")
                contour_with_out_deform_image_path = os.path.join(self.threed_images_report_folder,self.threed_window_name+"_"+section.lower()+"_contour_without_deformation"+".png")
                model_color_image_path =  os.path.join(self.threed_images_report_folder,self.threed_window_name+"_"+section .lower()+"_model_color"+".png")
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
                utils.MetaCommand('color pid reset act')

        return 0
