# PYTHON script
"""
This script is used for all the automation process of Body In White CBU deformation slide of thesis report.
"""

import os
from datetime import datetime
import logging

from meta import utils

from src.meta_utilities import capture_image
from src.meta_utilities import visualize_3d_critical_section
from src.metadb_info import GeneralVarInfo

class BIWCBUDeformationSlide():
    """
       This class is used to automate the biw cbu deformation slide of thesis report.

        Args:
            slide (object): biw deformation pptx slide object.
            general_input (GeneralInfo): GeneralInfo class object.
            metadb_3d_input (Meta3DInfo): Meta3DInfo class object.
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
        This method is used to iterate all the shapes of the biw cbu deformation slide and insert respective data.

        Returns:
            int: 0 Always for Sucess,1 for Failure.
        """
        try:
            self.logger.info("Started seeding data into biw cbu deformation slide")
            self.logger.info("")
            starttime = datetime.now()
            #maximizing the MetaPost window
            utils.MetaCommand('window maximize "{}"'.format(self.general_input.threed_window_name))
            #iterating through the shapes of the biw deformation slide
            for shape in self.shapes:
                #image insertion for the shape named "Image 1"
                if shape.name == "Image 1":
                    #visualizing all critical parts hes instances
                    utils.MetaCommand('0:options state variable "serial=1"')
                    utils.MetaCommand('grstyle scalarfringe disable')
                    data = self.metadb_3d_input.critical_sections["cbu"]
                    visualize_3d_critical_section(data,name = "cbu")
                    # entities = list()
                    # list_of_prop_names = list()
                    # for _key,value in data.items():
                    #     if 'hes' in value.keys():
                    #         prop_names = value['hes']
                    #         list_of_prop_names.append(prop_names)
                    #         re_props = prop_names.split(",")
                    #         for re_prop in re_props:
                    #             entities.extend(self.metadb_3d_input.get_props(re_prop))
                    # self.metadb_3d_input.hide_all()
                    # self.metadb_3d_input.show_only_props(entities)
                    #utils.MetaCommand('color pid transparency reset act')
                    utils.MetaCommand('view default isometric')
                    utils.MetaCommand('options fringebar off')
                    if self.threed_images_report_folder is not None:
                        #capturing cbu image
                        image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_CBU_AT_PEAK_STATE_WITHOUT_PLASTIC_STRAIN"+".png").replace(" ","_")
                        capture_image(image_path,self.general_input.threed_window_name,shape.width,shape.height,transparent=True)
                        self.logger.info("--- 3D MODEL IMAGE GENERATOR")
                        self.logger.info("")
                        self.logger.info("SOURCE WINDOW : {} ".format(self.general_input.threed_window_name))
                        self.logger.info("SOURCE MODEL : 0")
                        self.logger.info("STATE : PEAK STATE WITHOUT PLASTIC STRAIN")
                        self.logger.info("PID NAME SHOW FILTER : {} ".format(",".join(list_of_prop_names)))
                        self.logger.info("ADDITIONAL PID'S SHOWN : null ")
                        self.logger.info("PID NAME ERASE FILTER : null ")
                        self.logger.info("PID'S TO ERASE : null ")
                        self.logger.info("ERASE BOX : null ")
                        self.logger.info("IMAGE VIEW : null ")
                        self.logger.info("TRANSPARENCY LEVEL : null" )
                        self.logger.info("TRANSPARENT PID'S : null ")
                        self.logger.info("COMP NAME : CBU ")
                        self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                        self.logger.info("OUTPUT MODEL IMAGES :")
                        self.logger.info(image_path)
                        self.logger.info("")
                        #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                        transparent_image_path = image_path.replace(".png","")+"_transparent.png"
                        picture = self.shapes.add_picture(transparent_image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                        picture.crop_left = 0
                        picture.crop_right = 0
                        #removing transparent image
                        os.remove(transparent_image_path)
                        utils.MetaCommand('grstyle scalarfringe enable')
                    else:
                        self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                        self.logger.info("")
                #image insertion for the shape named "Image 1"
                elif shape.name == "Image 2":
                    #visualizing all critical parts hes instances
                    utils.MetaCommand('0:options state variable "serial=1"')
                    utils.MetaCommand('grstyle scalarfringe enable')
                    data = self.metadb_3d_input.critical_sections
                    entities = list()
                    list_of_prop_names = list()
                    for _key,value in data.items():
                        if 'hes' in value.keys():
                            prop_names = value['hes']
                            list_of_prop_names.append(prop_names)
                            re_props = prop_names.split(",")
                            for re_prop in re_props:
                                entities.extend(self.metadb_3d_input.get_props(re_prop))
                    self.metadb_3d_input.hide_all()
                    self.metadb_3d_input.show_only_props(entities)
                    utils.MetaCommand('color pid transparency reset act')
                    utils.MetaCommand('view default isometric')
                    utils.MetaCommand('options fringebar off')
                    if self.threed_images_report_folder is not None:
                        #capturing cbu image with plastic strain
                        image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_CBU_AT_PEAK_STATE_WITH_PLASTIC_STRAIN"+".png").replace(" ","_")
                        capture_image(image_path,self.general_input.threed_window_name,shape.width,shape.height,transparent=True)
                        self.logger.info("--- 3D MODEL IMAGE GENERATOR")
                        self.logger.info("")
                        self.logger.info("SOURCE WINDOW : {} ".format(self.general_input.threed_window_name))
                        self.logger.info("SOURCE MODEL : 0")
                        self.logger.info("STATE : PEAK STATE WITH PLASTIC STRAIN")
                        self.logger.info("PID NAME SHOW FILTER : {} ".format(",".join(list_of_prop_names)))
                        self.logger.info("ADDITIONAL PID'S SHOWN : null ")
                        self.logger.info("PID NAME ERASE FILTER : null ")
                        self.logger.info("PID'S TO ERASE : null ")
                        self.logger.info("ERASE BOX : null ")
                        self.logger.info("IMAGE VIEW : null ")
                        self.logger.info("TRANSPARENCY LEVEL : null" )
                        self.logger.info("TRANSPARENT PID'S : null ")
                        self.logger.info("COMP NAME : CBU ")
                        self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                        self.logger.info("OUTPUT MODEL IMAGES :")
                        self.logger.info(image_path)
                        self.logger.info("")
                        #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                        transparent_image_path = image_path.replace(".png","")+"_transparent.png"
                        picture = self.shapes.add_picture(transparent_image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                        picture.crop_left = 0
                        picture.crop_right = 0
                        #removing transparent image
                        os.remove(transparent_image_path)
                        utils.MetaCommand('0:options state variable "serial=0"')
                    else:
                        self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                        self.logger.info("")
                #image insertion for the shape named "Image 3"
                elif shape.name == "Image 3":
                    #capturing fringe bar of metapost window
                    utils.MetaCommand('add all')
                    utils.MetaCommand('add invert')
                    utils.MetaCommand('options fringebar on')
                    if self.threed_images_report_folder is not None:
                        image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_FRINGE_BAR"+".png").replace(" ","_")
                        utils.MetaCommand('write scalarfringebar png {} '.format(image_path))
                        self.logger.info("--- 3D FRINGE BAR IMAGE GENERATOR")
                        self.logger.info("")
                        self.logger.info("SOURCE WINDOW : {} ".format(self.general_input.threed_window_name))
                        self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                        self.logger.info("OUTPUT MODEL IMAGES :")
                        self.logger.info(image_path)
                        self.logger.info("")
                        #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                        picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                        picture.crop_left = 0
                        picture.crop_right = 0
                        utils.MetaCommand('options fringebar off')
                    else:
                        self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                        self.logger.info("")
            endtime = datetime.now()
        except Exception as e:
            self.logger.exception("Error while seeding data into biw cbu deformation slide:\n{}".format(e))
            self.logger.info("")
            return 1
        self.logger.info("Completed seeding data into biw cbu deformation slide")
        self.logger.info("Time Taken : {}".format(endtime - starttime))
        self.logger.info("")

        return 0
