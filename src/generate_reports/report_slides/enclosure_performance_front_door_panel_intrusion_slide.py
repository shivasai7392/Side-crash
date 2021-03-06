# PYTHON script
"""
This script is used for all the automation process of Enclosure performance front door panel intrusion slide of thesis report.
"""

import os
import logging
from datetime import datetime

from meta import utils
from meta import plot2d
from meta import windows

from src.meta_utilities import capture_image_and_resize
from src.general_utilities import add_row
from src.metadb_info import GeneralVarInfo

class EnclosurePerformanceFrontDoorPanelIntrusionSlide():
    """
       This class is used to automate the enclosure performance front door panel intrusion slide of thesis report.

        Args:
            slide (object): enclosure performance front door panel intrsusion pptx slide object.
            general_input (GeneralInfo): GeneralInfo class object.
            twod_images_report_folder (str): folder path to save twod data images.
        """
    def __init__(self,
                slide,
                general_input,
                twod_images_report_folder) -> None:
        self.shapes = slide.shapes
        self.general_input = general_input
        self.twod_images_report_folder = twod_images_report_folder
        self.intrusion_areas = {"ROW 1":{}}
        self.logger = logging.getLogger("side_crash_logger")

    def intrusion_curve_format(self,source_window,curve,temporary_window_name,curve_name):
        """
        This method is used for formatting axis,title options and attributes of curve of source window.

        Args:
            source_window (str): source plot window name.
            curve (object): source curve object.
            temporary_window_name (str): window name.
            curve_name (str): Curve name to set.

        Returns:
            int: 0 Always for Sucess.1 for Failure.
        """
        try:
            utils.MetaCommand('xyplot curve copy "{}" {}'.format(source_window,curve.id))
            utils.MetaCommand('xyplot create "{}"'.format(temporary_window_name))
            utils.MetaCommand('xyplot curve paste "{}" 0 {}'.format(temporary_window_name,curve.id))
            win = windows.Window(temporary_window_name, page_id=0)
            curve = win.get_curves('all')[0]
            y_values = []
            for x in [0.03,0.04,0.05,0.06,0.07,0.08]:
                y_values.append(round(curve.get_y_values_from_x(specifier = 'first', xvalue =x)[0]))
            self.intrusion_areas[curve_name.rsplit(" ",1)[0]][curve_name.rsplit(" ",1)[1]] = y_values
            utils.MetaCommand('xyplot gridoptions line major style "{}" 0 2'.format(temporary_window_name))
            utils.MetaCommand('xyplot axisoptions yaxis active "{}" 0 0'.format(temporary_window_name))
            utils.MetaCommand('xyplot axisoptions labels yposition "{}" 0 left'.format(temporary_window_name))
            utils.MetaCommand('xyplot axisoptions labels yalign "{}" 0 left'.format(temporary_window_name))
            utils.MetaCommand('xyplot axisoptions axyrange "{}" 0 0 0 600'.format(temporary_window_name))
            utils.MetaCommand('xyplot gridoptions yspace "{}" 0 40'.format(temporary_window_name))
            utils.MetaCommand('xyplot plotoptions title set "{}" 0 "{}"'.format(temporary_window_name,curve_name))
            utils.MetaCommand('xyplot axisoptions ylabel set "{}" 0 "Intrusion [mm]"'.format(temporary_window_name))
            utils.MetaCommand('xyplot curve select "{}" all'.format(temporary_window_name))
            utils.MetaCommand('xyplot curve set style "{}" selected 0'.format(temporary_window_name))
            utils.MetaCommand('xyplot curve set linewidth "{}" selected 9'.format(temporary_window_name))
            utils.MetaCommand('xyplot axisoptions ylabel font "{}" 0 "Arial,44,-1,5,75,0,0,0,0,0"'.format(temporary_window_name))
            utils.MetaCommand('xyplot axisoptions labels yfont "{}" 0 "Arial,44,-1,5,75,0,0,0,0,0"'.format(temporary_window_name))
            utils.MetaCommand('xyplot axisoptions yaxis deactive "{}" 0 0'.format(temporary_window_name))
            utils.MetaCommand('xyplot axisoptions xaxis active "{}" 0 0'.format(temporary_window_name))
            utils.MetaCommand('xyplot axisoptions xlabel font "{}" 0 "Arial,44,-1,5,75,0,0,0,0,0"'.format(temporary_window_name))
            utils.MetaCommand('xyplot axisoptions labels xfont "{}" 0 "Arial,44,-1,5,75,0,0,0,0,0"'.format(temporary_window_name))
            utils.MetaCommand('xyplot plotoptions title font "{}" 0 "Arial,44,-1,5,75,0,0,0,0,0"'.format(temporary_window_name))
            utils.MetaCommand('xyplot axisoptions xaxis deactive "{}" 0 0'.format(temporary_window_name))
        except:
            return 1

        return 0

    def edit(self):
        """
        This method is used to iterate all the shapes of the enclosure performance front door panel intrusion slide and insert respective data.

        Returns:
            int: 0 Always for Sucess,1 for Failure.
        """
        from pptx.util import Pt

        try:
            self.logger.info("Started seeding data into enclosure performance front door panel intrusion slide")
            self.logger.info("")
            starttime = datetime.now()
            #iterating through the Image shapes of the enclosure performance front door panel intrusion slide
            if self.general_input.front_door_accel_window_name not in ["null","none",""]:
                front_door_accel_window_name = self.general_input.front_door_accel_window_name
                front_door_accel_window_obj = windows.WindowByName(front_door_accel_window_name)
                if front_door_accel_window_obj:
                    utils.MetaCommand('window maximize {}'.format(front_door_accel_window_name))
                    for shape in self.shapes:
                        #image insertion for the shape named "Image 1"
                        if shape.name == "Image 1":
                            #getting "Front Door Accel" window and "Front shoulder intrusion curve" curve objects to activate and format the curve axisoptions and title options
                            temporary_window_name = "Temporary"
                            if self.general_input.front_shoulder_intrusion_curve_name not in ["null","none",""]:
                                front_shoulder_intrusion_curve_name = self.general_input.front_shoulder_intrusion_curve_name
                                curves = plot2d.CurvesByName(front_door_accel_window_name, front_shoulder_intrusion_curve_name, 1)
                                if curves:
                                    curves[0].show()
                                    self.intrusion_curve_format(front_door_accel_window_name,curves[0],temporary_window_name,"ROW 1 SHOULDER")
                                else:
                                    self.logger.info("ERROR : Front Door Accel window does not contain '{}' curve from META 2D variable {}. Please update.".format(front_shoulder_intrusion_curve_name,GeneralVarInfo.front_shoulder_intrusion_curve_key))
                                    self.logger.info("")
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.front_shoulder_intrusion_curve_key))
                                self.logger.info("")
                            #capturing image of the formatted intrusion curve
                            if self.twod_images_report_folder:
                                image_path = os.path.join(self.twod_images_report_folder,front_door_accel_window_name+"_"+"ROW 1 SHOULDER"+".png").replace(" ","_")
                                capture_image_and_resize(image_path,shape.width,shape.height)
                                self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                                self.logger.info("")
                                self.logger.info("CURVES : {} | FROM VARIABLE : {} | SOURCE WINDOW : {}".format(self.general_input.front_shoulder_intrusion_curve_name,self.general_input.front_shoulder_intrusion_curve_key,front_door_accel_window_name))
                                self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                                self.logger.info("OUTPUT CURVE IMAGES : ")
                                self.logger.info(image_path)
                                self.logger.info("")
                                #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                                picture.crop_left = 0
                                picture.crop_right = 0
                                utils.MetaCommand('window delete "{}"'.format(temporary_window_name))
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                                self.logger.info("")
                        #image insertion for the shape named "Image 2"
                        elif shape.name == "Image 2":
                            #getting "Front Door Accel" window and "Front abdomen intrusion curve" curve objects to activate and format the curve axisoptions and title options
                            temporary_window_name = "Temporary"
                            if self.general_input.front_abdomen_intrusion_curve_name not in ["null","none",""]:
                                front_abdomen_intrusion_curve_name = self.general_input.front_abdomen_intrusion_curve_name
                                curves = plot2d.CurvesByName(front_door_accel_window_name, front_abdomen_intrusion_curve_name, 1)
                                if curves:
                                    curves[0].show()
                                    self.intrusion_curve_format(front_door_accel_window_name,curves[0],temporary_window_name,"ROW 1 ABDOMEN")
                                else:
                                    self.logger.info("ERROR : Front Door Accel window does not contain '{}' curve from META 2D variable {}. Please update.".format(front_abdomen_intrusion_curve_name,GeneralVarInfo.front_abdomen_intrusion_curve_key))
                                    self.logger.info("")
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.front_abdomen_intrusion_curve_key))
                                self.logger.info("")
                            #capturing image of the formatted intrusion curve
                            if self.twod_images_report_folder is not None:
                                image_path = os.path.join(self.twod_images_report_folder,front_door_accel_window_name+"_"+"ROW 1 ABDOMEN"+".png").replace(" ","_")
                                capture_image_and_resize(image_path,shape.width,shape.height)
                                self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                                self.logger.info("")
                                self.logger.info("CURVES : {} | FROM VARIABLE : {} | SOURCE WINDOW : {}".format(self.general_input.front_abdomen_intrusion_curve_name,self.general_input.front_abdomen_intrusion_curve_key,front_door_accel_window_name))
                                self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                                self.logger.info("OUTPUT CURVE IMAGES : ")
                                self.logger.info(image_path)
                                self.logger.info("")
                                #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                                picture.crop_left = 0
                                picture.crop_right = 0
                                utils.MetaCommand('window delete "{}"'.format(temporary_window_name))
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                                self.logger.info("")
                        #image insertion for the shape named "Image 4"
                        elif shape.name == "Image 4":
                            #getting "Front Door Accel" window and "Front femur intrusion curve" curve objects to activate and format the curve axisoptions and title options
                            temporary_window_name = "Temporary"
                            if self.general_input.front_femur_intrusion_curve_name not in ["null","none",""]:
                                front_femur_intrusion_curve_name = self.general_input.front_femur_intrusion_curve_name
                                curves = plot2d.CurvesByName(front_door_accel_window_name, front_femur_intrusion_curve_name, 1)
                                if curves:
                                    curves[0].show()
                                    self.intrusion_curve_format(front_door_accel_window_name,curves[0],temporary_window_name,"ROW 1 FEMUR")
                                else:
                                    self.logger.info("ERROR : Front Door Accel window does not contain '{}' curve from META 2D variable {}. Please update.".format(front_femur_intrusion_curve_name,GeneralVarInfo.front_femur_intrusion_curve_key))
                                    self.logger.info("")
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.front_femur_intrusion_curve_key))
                                self.logger.info("")
                            #capturing image of the formatted intrusion curve
                            if self.twod_images_report_folder is not None:
                                image_path = os.path.join(self.twod_images_report_folder,front_door_accel_window_name+"_"+"ROW 1 FEMUR"+".jepg").replace(" ","_")
                                capture_image_and_resize(image_path,shape.width,shape.height)
                                self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                                self.logger.info("")
                                self.logger.info("CURVES : {} | FROM VARIABLE : {} | SOURCE WINDOW : {}".format(self.general_input.front_femur_intrusion_curve_name,self.general_input.front_femur_intrusion_curve_key,front_door_accel_window_name))
                                self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                                self.logger.info("OUTPUT CURVE IMAGES : ")
                                self.logger.info(image_path)
                                self.logger.info("")
                                #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                                picture.crop_left = 0
                                picture.crop_right = 0
                                utils.MetaCommand('window delete "{}"'.format(temporary_window_name))
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                                self.logger.info("")
                        #image insertion for the shape named "Image 3"
                        elif shape.name == "Image 3":
                            #getting "Front Door Accel" window and "Front pelvis intrusion curve" curve objects to activate and format the curve axisoptions and title options
                            temporary_window_name = "Temporary"
                            if self.general_input.front_pelvis_intrusion_curve_name not in ["null","none",""]:
                                front_pelvis_intrusion_curve_name = self.general_input.front_pelvis_intrusion_curve_name
                                curves = plot2d.CurvesByName(front_door_accel_window_name, front_pelvis_intrusion_curve_name, 1)
                                if curves:
                                    curves[0].show()
                                    self.intrusion_curve_format(front_door_accel_window_name,curves[0],temporary_window_name,"ROW 1 PELVIS")
                                else:
                                    self.logger.info("ERROR : Front Door Accel window does not contain '{}' curve from META 2D variable {}. Please update.".format(front_pelvis_intrusion_curve_name,GeneralVarInfo.front_pelvis_intrusion_curve_key))
                                    self.logger.info("")
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.front_pelvis_intrusion_curve_key))
                                self.logger.info("")
                            #capturing image of the formatted intrusion curve
                            if self.twod_images_report_folder is not None:
                                image_path = os.path.join(self.twod_images_report_folder,front_door_accel_window_name+"_"+"ROW 1 PELVIS"+".png").replace(" ","_")
                                capture_image_and_resize(image_path,shape.width,shape.height)
                                self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                                self.logger.info("")
                                self.logger.info("CURVES : {} | FROM VARIABLE : {} | SOURCE WINDOW : {}".format(self.general_input.front_pelvis_intrusion_curve_name,self.general_input.front_pelvis_intrusion_curve_key,front_door_accel_window_name))
                                self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                                self.logger.info("OUTPUT CURVE IMAGES : ")
                                self.logger.info(image_path)
                                self.logger.info("")
                                #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                                picture.crop_left = 0
                                picture.crop_right = 0
                                utils.MetaCommand('window delete "{}"'.format(temporary_window_name))
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                                self.logger.info("")
                else:
                    self.logger.info("ERROR : 2D METADB does not contain 'Front Door - Accel' window. Please update.")
            else:
                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.front_door_accel_window_key))
            #iterating through the Table shapes of the enclosure performance front door panel intrusion slide
            for shape in self.shapes:
                #table population for the table named as "Table 1"
                if shape.name == "Table 1":
                    #getting table object
                    table = shape.table
                    row_index = 2
                    node_index = 0
                    #adding new row to the table
                    add_row(table)
                    #iterating through intrusion curve data of "ROW 1"
                    for iindex,(key,values) in enumerate(self.intrusion_areas["ROW 1"].items()):
                        if iindex == 2:
                            #adding a new row
                            add_row(table)
                            row_index = row_index+1
                        if iindex == 1 or iindex == 3:
                            node_index = 7
                        else:
                            node_index = 0
                        #getting table row objects
                        rows = table.rows
                        #inserting heading into the cell
                        text_frame_1 = rows[row_index].cells[node_index].text_frame
                        font = text_frame_1.paragraphs[0].font
                        font.bold = True
                        font.size = Pt(11)
                        text_frame_1.paragraphs[0].text = key.capitalize()
                        #inserting y values in to the cells
                        for index,value in enumerate(values):
                            text_frame = rows[row_index].cells[node_index+index+1].text_frame
                            font = text_frame.paragraphs[0].font
                            font.size = Pt(11)
                            text_frame.paragraphs[0].text = str(value)
            endtime = datetime.now()
        except Exception as e:
            self.logger.exception("Error while seeding data into enclosure performance front door panel intrusion slide:\n{}".format(e))
            self.logger.info("")
            return 1
        self.logger.info("Completed seeding data into enclosure performance front door panel intrusion slide")
        self.logger.info("Time Taken : {}".format(endtime - starttime))
        self.logger.info("")

        return 0
