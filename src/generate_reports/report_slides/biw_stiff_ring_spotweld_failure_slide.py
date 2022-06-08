# PYTHON script
"""
This script is used for all the automation process of Body In White stiff ring spotweld failure slide of thesis report.
"""

import os
import logging
from datetime import datetime

from meta import utils
from meta import models

from src.meta_utilities import capture_image_and_resize
from src.meta_utilities import visualize_3d_critical_section
from src.meta_utilities import visualize_annotation
from src.metadb_info import GeneralVarInfo

class BIWStiffRingSpotWeldFailureSlide():
    """
       This class is used to automate the biw stiff ring spotweld failure slide of thesis report.

        Args:
            slide (object): biw stiff ring spotweld failure pptx slide object.
            general_input (GeneralInfo): GeneralInfo class object.
            metadb_3d_input (Meta3DInfo): Meta3DInfo class object.
            twod_images_report_folder (str): folder path to save twod data images.
            threed_images_report_folder (str): folder path to save threed data images.
        """
    def __init__(self,
                slide,
                general_input,
                metadb_3d_input,
                threed_images_report_folder) -> None:
        self.shapes = slide.shapes
        self.general_input = general_input
        self.metadb_3d_input = metadb_3d_input
        self.threed_images_report_folder = threed_images_report_folder
        self.logger = logging.getLogger("side_crash_logger")

    def edit(self):
        """
        This method is used to iterate all the shapes of the biw stiff ring spotweld failure slide and insert respective data.

        Returns:
            int: 0 Always for Sucess,1 for Failure.
        """
        try:
            self.logger.info("Started seeding data into  biw stiff ring spotweld failure slide")
            self.logger.info("")
            starttime = datetime.now()
            #iterating through the shapes of the  biw stiff ring spotweld failure slide
            for shape in self.shapes:
                #image insertion for the shape named "Image 1"
                if shape.name == "Image 1":
                    #visualising and capturing image of "f21_upb_outer" critical part set with spotweld failure
                    utils.MetaCommand('window maximize {}'.format(self.general_input.threed_window_name))
                    utils.MetaCommand('0:options state original')
                    utils.MetaCommand('options fringebar off')
                    data = self.metadb_3d_input.critical_sections["f21_upb_outer"]
                    visualize_3d_critical_section(data,name = "f21_upb_outer")
                    self.logger.info("--- HDMA VISIBLE SPOTWELD ANALYSIS")
                    self.logger.info("START TIME : {}".format(datetime.now().strftime("%H:%M:%S")))
                    self.logger.info("THRESHOLD : {} | SOURCE MODEL ID : 0 | SOURCE WINDOW NAME : MetaPost | OUTPUT WINDOW NAME : MetaPost".format("0.7"))
                    self.logger.info("")
                    self.logger.info("SOURCE FILE FOR SPOTWELD ID'S : {}".format(self.general_input.d3hsp_file_path))
                    visualize_annotation(self.metadb_3d_input.spotweld_clusters,self.general_input.binout_directory)
                    #utils.MetaCommand('read session {}'.format(os.path.join(os.path.dirname(__file__),"annotation_positioning_session_files","f21_upb_outer.ses")))
                    if self.threed_images_report_folder is not None:
                        image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_"+"F21_UPB_OUTER_SPOTWELD_FAILURE"+".png").replace(" ","_")
                        utils.MetaCommand('window maximize {}'.format(self.general_input.threed_window_name))
                        capture_image_and_resize(image_path,shape.width,shape.height,resize = False)
                        self.logger.info("--- 3D MODEL IMAGE GENERATOR")
                        self.logger.info("")
                        self.logger.info("SOURCE WINDOW : {} ".format(self.general_input.threed_window_name))
                        self.logger.info("SOURCE MODEL : 0")
                        self.logger.info("STATE : ORIGINAL STATE")
                        self.logger.info("PID NAME SHOW FILTER : {} ".format(data["hes"] if "hes" in data.keys() else "null"))
                        self.logger.info("ADDITIONAL PID'S SHOWN : {} ".format(data["hes_exceptions"] if "hes_exceptions" in data.keys() else "null"))
                        self.logger.info("PID NAME ERASE FILTER : {} ".format(data["hes_exceptions"] if "hes_exceptions" in data.keys() else "null"))
                        self.logger.info("PID'S TO ERASE : {} ".format(data["erase_pids"] if "erase_pids" in data.keys() else "null"))
                        self.logger.info("ERASE BOX : {} ".format(data["erase_box"] if "erase_box" in data.keys() else "null"))
                        self.logger.info("IMAGE VIEW : {} ".format(data["view"] if "view" in data.keys() else "null"))
                        self.logger.info("TRANSPARENCY LEVEL : 50" )
                        self.logger.info("TRANSPARENT PID'S : {} ".format(data["transparent_pids"] if "transparent_pids" in data.keys() else "null"))
                        self.logger.info("COMP NAME : {} ".format(data["name"] if "name" in data.keys() else "null"))
                        self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                        self.logger.info("OUTPUT MODEL IMAGES :")
                        self.logger.info(image_path)
                        self.logger.info("")
                        #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                        picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                        picture.crop_left = 0
                        picture.crop_right = 0
                        #reverting visual settings
                        utils.MetaCommand('annotation delete all')
                        utils.MetaCommand('color pid transparency reset act')
                    else:
                        self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                        self.logger.info("")
                #image insertion for the shape named "Image 2"
                elif shape.name == "Image 2":
                    #visualising and capturing image of "f21_upb_inner" critical part set with spotweld failure
                    utils.MetaCommand('window maximize {}'.format(self.general_input.threed_window_name))
                    utils.MetaCommand('0:options state original')
                    utils.MetaCommand('options fringebar off')
                    data = self.metadb_3d_input.critical_sections["f21_upb_inner"]
                    visualize_3d_critical_section(data,name = "f21_upb_inner")
                    self.logger.info("--- HDMA VISIBLE SPOTWELD ANALYSIS")
                    self.logger.info("START TIME : {}".format(datetime.now().strftime("%H:%M:%S")))
                    self.logger.info("THRESHOLD : {} | SOURCE MODEL ID : 0 | SOURCE WINDOW NAME : MetaPost | OUTPUT WINDOW NAME : MetaPost".format("0.7"))
                    self.logger.info("")
                    self.logger.info("SOURCE FILE FOR SPOTWELD ID'S : {}".format(self.general_input.d3hsp_file_path))
                    visualize_annotation(self.metadb_3d_input.spotweld_clusters,self.general_input.binout_directory)
                    #utils.MetaCommand('read session {}'.format(os.path.join(os.path.dirname(__file__),"annotation_positioning_session_files","f21_upb_inner.ses")))
                    if self.threed_images_report_folder is not None:
                        image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_"+"F21_UPB_INNER_SPOTWELD_FAILURE"+".png").replace(" ","_")
                        utils.MetaCommand('window maximize {}'.format(self.general_input.threed_window_name))
                        capture_image_and_resize(image_path,shape.width,shape.height,resize = False)
                        self.logger.info("--- 3D MODEL IMAGE GENERATOR")
                        self.logger.info("")
                        self.logger.info("SOURCE WINDOW : {} ".format(self.general_input.threed_window_name))
                        self.logger.info("SOURCE MODEL : 0")
                        self.logger.info("STATE : ORIGINAL STATE")
                        self.logger.info("PID NAME SHOW FILTER : {} ".format(data["hes"] if "hes" in data.keys() else "null"))
                        self.logger.info("ADDITIONAL PID'S SHOWN : {} ".format(data["hes_exceptions"] if "hes_exceptions" in data.keys() else "null"))
                        self.logger.info("PID NAME ERASE FILTER : {} ".format(data["hes_exceptions"] if "hes_exceptions" in data.keys() else "null"))
                        self.logger.info("PID'S TO ERASE : {} ".format(data["erase_pids"] if "erase_pids" in data.keys() else "null"))
                        self.logger.info("ERASE BOX : {} ".format(data["erase_box"] if "erase_box" in data.keys() else "null"))
                        self.logger.info("IMAGE VIEW : {} ".format(data["view"] if "view" in data.keys() else "null"))
                        self.logger.info("TRANSPARENCY LEVEL : 50" )
                        self.logger.info("TRANSPARENT PID'S : {} ".format(data["transparent_pids"] if "transparent_pids" in data.keys() else "null"))
                        self.logger.info("COMP NAME : {} ".format(data["name"] if "name" in data.keys() else "null"))
                        self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                        self.logger.info("OUTPUT MODEL IMAGES :")
                        self.logger.info(image_path)
                        self.logger.info("")
                        #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                        picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                        picture.crop_left = 0
                        picture.crop_right = 0
                        utils.MetaCommand('color pid transparency reset act')
                        #reverting visual settings
                        utils.MetaCommand('annotation delete all')
                    else:
                        self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                        self.logger.info("")
            endtime = datetime.now()
        except Exception as e:
            self.logger.exception("Error while seeding data into  biw stiff ring spotweld failure slide:\n{}".format(e))
            self.logger.info("")
            return 1
        self.logger.info("Completed seeding data into  biw stiff ring spotweld failure slide")
        self.logger.info("Time Taken : {}".format(endtime - starttime))
        self.logger.info("")

        return 0
