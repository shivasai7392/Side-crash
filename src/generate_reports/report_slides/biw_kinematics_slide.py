# PYTHON script
"""
This script is used for all the automation process of Body-In-White Kinematics slide of thesis report.
"""

import os
import logging
from datetime import datetime

from meta import utils
from meta import plot2d
from meta import windows

from src.meta_utilities import capture_image
from src.meta_utilities import capture_image_and_resize
from src.meta_utilities import visualize_3d_critical_section
from src.general_utilities import clone_shape
from src.metadb_info import GeneralVarInfo

class BIWKinematicsSlide():
    """
        This class is used to automate the biw kinematics slide of thesis report.

        Args:
            slide (object): executive report pptx slide object.
            general_input (GeneralInfo): GeneralInfo class object.
            metadb_3d_input (Meta3DInfo): Meta3DInfo class object.
            twod_images_report_folder (str): folder path to save twod data images.
            threed_images_report_folder (str): folder path to save threed data images.
        """
    def __init__(self,
                slide,
                general_input,
                metadb_3d_input,
                twod_images_report_folder,
                threed_images_report_folder) -> None:
        self.shapes = slide.shapes
        self.general_input = general_input
        self.metadb_3d_input = metadb_3d_input
        self.twod_images_report_folder = twod_images_report_folder
        self.threed_images_report_folder = threed_images_report_folder
        self.biw_accel_window_name = None
        self.biw_accel_window_obj = None
        self.biw_accel_window_layout = None
        self.activated_plot = None
        self.logger = logging.getLogger("side_crash_logger")

    def setup(self,format_type = "2d"):
        """
        This method is used to setup data for meta windows based on 3d and 2d formats

        Args:
            format_type (str, optional): format type. Defaults to "2d".

        Returns:
            int: 0 Always for Sucess.1 for Failure.
        """
        try:
            if format_type == "2d":
                self.biw_accel_window_name = self.general_input.biw_accel_window_name
                self.biw_accel_window_name = self.biw_accel_window_name.replace("\"","")
                utils.MetaCommand('window maximize "{}"'.format(self.biw_accel_window_name))
                self.biw_accel_window_obj = windows.Window(self.biw_accel_window_name, page_id = 0)
                self.biw_accel_window_layout = self.biw_accel_window_obj.get_plot_layout()
            else:
                utils.MetaCommand('window maximize "MetaPost"')
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('grstyle scalarfringe disable')
                utils.MetaCommand('options fringebar off')
                #visualizing cbu critical part set to capture whole cbu and barrier image at peak state in gray color
                data = self.metadb_3d_input.critical_sections["cbu"]
                visualize_3d_critical_section(data,name = "cbu")
                # critical_data = self.metadb_3d_input.critical_sections
                # for (index,(critical_section,value)) in enumerate(critical_data.items()):
                #     and_filter = False
                #     if index>0:
                #         and_filter = True
                #     visualize_3d_critical_section(value,and_filter = and_filter,name = critical_section)
                utils.MetaCommand('color pid Gray act')
        except Exception  as e:
            return 1

        return 0

    def kinematics_curve_format(self, biw_accel_window_name, plot_id,velocity_min_max_values, displacement_min_max_values,title = None):
        """
        This method is used to format kinematics curves.

        Args:
            biw_accel_window_name (str): biw accel window name.
            plot_id (int): id of the plot.
            velocity_min_max_values (list): list of velocity curve y min and max values.
            displacement_min_max_values (list): list of displacement curve y min nad max values.
            title (str, optional): title of the plot. Defaults to None.

        Returns:
            int: 0 Always for Sucess.1 for Failure.
        """
        try:
            # self.logger.info("Intrusion curve fomat : {}".format(curve_name))
            # self.logger.info("")
            # starttime = datetime.now()
            # self.logger.info("Moving the {} curve to a Temporary window for custom formatting of title,yaxis,xaxis options and attibutes".format(curve_name))
            # self.logger.info(".")
            # self.logger.info(".")
            # self.logger.info(".")
            #rounding the velocity,displacement curves y min and max values
            velocity_values = [round(each_velocity_min_max_value) for each_velocity_min_max_value in velocity_min_max_values]
            displacement_values = [round(each_displacement_min_max_value) for each_displacement_min_max_value in displacement_min_max_values]
            #applying custom style and size for plot title,xaxis,yaxis options and attributes
            utils.MetaCommand('xyplot plotactive "{}" {}'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot plotoptions value off "{}" {}'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions yaxis active "{}" {} 0'.format(biw_accel_window_name,plot_id))
            utils.MetaCommand('xyplot axisoptions yaxis hideaxis "{}" {} 0'.format(biw_accel_window_name,plot_id))
            utils.MetaCommand('xyplot curve select "{}" vis'.format(biw_accel_window_name))
            utils.MetaCommand('xyplot curve set linewidth "{}" selected 9'.format(biw_accel_window_name))
            if title:
                utils.MetaCommand('xyplot plotoptions title set "{}" {} "{}"'.format(biw_accel_window_name, plot_id, title))
            utils.MetaCommand('xyplot axisoptions xaxis active "{}" {} 0'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions xlabel font "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions labels xfont "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions xaxis deactive "{}" {} 0'.format(biw_accel_window_name, plot_id))

            utils.MetaCommand('xyplot axisoptions yaxis active "{}" {} 1'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions ylabel font "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions axyrange "{}" {} 1 {} {}'.format(biw_accel_window_name, plot_id,str(round((min(velocity_values)-1000)//1000)*1000), str(round(((max(velocity_values)+1000)//1000)*1000))))
            utils.MetaCommand('xyplot gridoptions yspace "{}" {} 500'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions labels yfont "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions yaxis deactive "{}" {} 1'.format(biw_accel_window_name, plot_id))

            utils.MetaCommand('xyplot axisoptions yaxis active "{}" {} 2'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions ylabel font "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions axyrange "{}" {} 2 {} {}'.format(biw_accel_window_name, plot_id,str(min(displacement_values)-100), str(max(displacement_values)+100)))
            utils.MetaCommand('xyplot gridoptions yspace "{}" {} 50'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions labels yfont "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions yaxis deactive "{}" {} 2'.format(biw_accel_window_name, plot_id))

            utils.MetaCommand('xyplot plotoptions title font "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
            #endtime = datetime.now()
        except Exception as e:
            self.logger.exception("Error while seeding data into executive report slide:\n{}".format(e))
            self.logger.info("")
            return 1
        # self.logger.info("Intrusion curve format complete")
        # self.logger.info("Time Taken : {}".format(endtime - starttime))
        # self.logger.info("")

        return 0

    def edit(self):
        """
        This method is used to iterate all the shapes of the biw kinematics slide and insert respective data.

        Returns:
            int: 0 Always for Sucess,1 for Failure.
        """
        from PIL import Image
        try:
            self.logger.info("Started seeding data into biw kinematics slide")
            self.logger.info("")
            starttime = datetime.now()
            oval_shapes = [shape for shape in self.shapes if "Oval" in shape.name]
            #setting 3d data settings
            self.setup(format_type="3d")
            #iterating through the shapes of the biw kinematis slide
            for shape in self.shapes:
                #image insertion for the shape named "Image 1"
                if shape.name == "Image 1":
                    if self.threed_images_report_folder is not None:
                        #capturing "MetaPost" window image
                        image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_"+"CBU_AND_BARRIER_FRONT_VIEW_AT_PEAK_STATE"+".png").replace(" ","_")
                        capture_image(image_path,self.general_input.threed_window_name,shape.width,shape.height,view = "front",transparent=True)
                        self.logger.info("--- 3D MODEL IMAGE GENERATOR")
                        self.logger.info("")
                        self.logger.info("SOURCE WINDOW : {} ".format(self.general_input.threed_window_name))
                        self.logger.info("SOURCE MODEL : 0")
                        self.logger.info("STATE : PEAK STATE")
                        self.logger.info("PID NAME SHOW FILTER : {} ".format("CBU"))
                        self.logger.info("ADDITIONAL PID'S SHOWN : {} ".format("CBU"))
                        self.logger.info("IMAGE VIEW : TOP ")
                        self.logger.info("COMP NAME : CBU AND BARRIER ")
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
                    else:
                        self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                        self.logger.info("")
                #image insertion for the shape named "Image 2"
                elif shape.name == "Image 2":
                    if self.threed_images_report_folder is not None:
                        #capturing "MetaPost" window image
                        image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_"+"CBU_AND_BARRIER_TOP_VIEW_AT_PEAK_STATE"+".png").replace(" ","_")
                        capture_image(image_path,self.general_input.threed_window_name,shape.width,shape.height,view = "top",rotate = Image.ROTATE_90,transparent=True)
                        self.logger.info("--- 3D MODEL IMAGE GENERATOR")
                        self.logger.info("")
                        self.logger.info("SOURCE WINDOW : {} ".format(self.general_input.threed_window_name))
                        self.logger.info("SOURCE MODEL : 0")
                        self.logger.info("STATE : PEAK STATE")
                        self.logger.info("PID NAME SHOW FILTER : {} ".format("CBU"))
                        self.logger.info("ADDITIONAL PID'S SHOWN : {} ".format("CBU"))
                        self.logger.info("IMAGE VIEW : TOP ")
                        self.logger.info("COMP NAME : CBU AND BARRIER ")
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
                    else:
                        self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                        self.logger.info("")
                #image insertion for the shape named "Image 3"
                elif shape.name == "Image 3":
                    if self.threed_images_report_folder is not None:
                        #capturing "MetaPost" window image
                        image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_"+"CBU_AND_BARRIER_FRONT_VIEW_AT_FINAL_STATE"+".png").replace(" ","_")
                        utils.MetaCommand('0:options state variable "serial=2"')
                        capture_image(image_path,self.general_input.threed_window_name,shape.width,shape.height,view = "front",transparent=True)
                        self.logger.info("--- 3D MODEL IMAGE GENERATOR")
                        self.logger.info("")
                        self.logger.info("SOURCE WINDOW : {} ".format(self.general_input.threed_window_name))
                        self.logger.info("SOURCE MODEL : 0")
                        self.logger.info("STATE : FINAL STATE")
                        self.logger.info("PID NAME SHOW FILTER : {} ".format("CBU"))
                        self.logger.info("ADDITIONAL PID'S SHOWN : {} ".format("CBU"))
                        self.logger.info("IMAGE VIEW : TOP ")
                        self.logger.info("COMP NAME : CBU AND BARRIER ")
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
                    else:
                        self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                        self.logger.info("")
                #image insertion for the shape named "Image 4"
                elif shape.name == "Image 4":
                    if self.threed_images_report_folder is not None:
                        utils.MetaCommand('color pid Gray act')
                        #capturing "MetaPost" window image
                        image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_"+"CBU_AND_BARRIER_TOP_VIEW_AT_FINAL_STATE"+".png").replace(" ","_")
                        utils.MetaCommand('0:options state variable "serial=2"')
                        capture_image(image_path,self.general_input.threed_window_name,shape.width,shape.height,view = "top",rotate = Image.ROTATE_90,transparent=True)
                        self.logger.info("--- 3D MODEL IMAGE GENERATOR")
                        self.logger.info("")
                        self.logger.info("SOURCE WINDOW : {} ".format(self.general_input.threed_window_name))
                        self.logger.info("SOURCE MODEL : 0")
                        self.logger.info("STATE : FINAL STATE")
                        self.logger.info("PID NAME SHOW FILTER : {} ".format("CBU"))
                        self.logger.info("ADDITIONAL PID'S SHOWN : {} ".format("CBU"))
                        self.logger.info("IMAGE VIEW : TOP ")
                        self.logger.info("COMP NAME : CBU AND BARRIER ")
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
                    else:
                        self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                        self.logger.info("")
            #reverting color and 3d data setup
            utils.MetaCommand('color pid reset act')
            self.revert(format_type="3d")

            if self.general_input.biw_accel_window_name not in ["null","none",""]:
                biw_accel_window_name = self.general_input.biw_accel_window_name.replace("\"","")
                biw_accel_window_obj = windows.WindowByName(biw_accel_window_name)
                if biw_accel_window_obj:
                    utils.MetaCommand('window maximize "{}"'.format(biw_accel_window_name))
                    #iterating through curve image shapes
                    for shape in self.shapes:
                        #image insertion for the shape named "Image 5"
                        if shape.name == "Image 5":
                            #calling default 2d data setup and getting the X velocity and X displacement curve from biw accel window of MDB plot id 3
                            self.setup()
                            plot_id = 3
                            page_id = 0
                            velocity_min_max_values = []
                            displacement_min_max_values = []
                            plot = plot2d.Plot(plot_id, biw_accel_window_name, page_id)
                            #collecting min max y values for setting the y axis range of the plot
                            biw_accel_mdb_x_velocity_curve = plot.get_curves('byname', name = '*X velocity')[0]
                            biw_accel_max_velocity = biw_accel_mdb_x_velocity_curve.get_limit_value_y(specifier = 'max')
                            velocity_min_max_values.append(biw_accel_max_velocity)
                            biw_accel_min_velocity = biw_accel_mdb_x_velocity_curve.get_limit_value_y(specifier = 'min')
                            velocity_min_max_values.append(biw_accel_min_velocity)
                            biw_accel_mdb_x_velocity_curve.show()
                            biw_accel_mdb_x_displacement_curve = plot.get_curves('byname', name = '*X displacement')[0]
                            biw_accel_max_curve_displacement = biw_accel_mdb_x_displacement_curve.get_limit_value_y(specifier = 'max')
                            displacement_min_max_values.append(biw_accel_max_curve_displacement)
                            biw_accel_min_displacement = biw_accel_mdb_x_displacement_curve.get_limit_value_y(specifier = 'min')
                            displacement_min_max_values.append(biw_accel_min_displacement)
                            biw_accel_mdb_x_displacement_curve.show()
                            #actiavting plot,getting plot title object and formatting the plot title,yaxis,xaxis options and attributes
                            title = plot2d.Title(plot_id, biw_accel_window_name, page_id)
                            plot.activate()
                            self.activated_plot = plot
                            utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, 1))
                            self.kinematics_curve_format(biw_accel_window_name, plot_id, velocity_min_max_values, displacement_min_max_values)
                            if self.twod_images_report_folder is not None:
                                #capturing "MDB" plot image
                                image_path = os.path.join(self.twod_images_report_folder,biw_accel_window_name+"_"+title.get_text()+".png").replace(" ","_")
                                capture_image_and_resize(image_path,shape.width,shape.height)
                                self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                                self.logger.info("")
                                self.logger.info("CURVES : {},{} | SOURCE PLOT : {} | SOURCE WINDOW : {}".format("*X velocity","*X displacement",title.get_text(),biw_accel_window_name))
                                self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                                self.logger.info("OUTPUT CURVE IMAGES : ")
                                self.logger.info(image_path)
                                self.logger.info("")
                                #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                                picture.crop_left = 0
                                picture.crop_right = 0
                                #reverting 2d data setup
                                biw_accel_mdb_x_displacement_curve.hide()
                                biw_accel_mdb_x_velocity_curve.hide()
                                self.revert()
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                                self.logger.info("")
                        #image insertion for the shape named "Image 6"
                        elif shape.name == "Image 6":
                            #calling default 2d data setup and getting the Y velocity and Y displacement curve from biw accel window of UNIT plot id 0
                            self.setup()
                            plot_id = 0
                            page_id = 0
                            plot = plot2d.Plot(plot_id, biw_accel_window_name, page_id)
                            velocity_min_max_values = []
                            displacement_min_max_values = []
                            #collecting min max y values for setting the y axis range of the plot
                            biw_accel_unit_y_velocity_curve = plot.get_curves('byname', name = '*UNIT*Y velocity')[0]
                            biw_accel_max_velocity = biw_accel_unit_y_velocity_curve.get_limit_value_y(specifier = 'max')
                            velocity_min_max_values.append(biw_accel_max_velocity)
                            biw_accel_min_velocity = biw_accel_unit_y_velocity_curve.get_limit_value_y(specifier = 'min')
                            velocity_min_max_values.append(biw_accel_min_velocity)
                            biw_accel_unit_y_velocity_curve.show()
                            biw_accel_unit_y_displacement_curve = plot.get_curves('byname', name = '*UNIT*Y displacement')[0]
                            biw_accel_max_curve_displacement = biw_accel_unit_y_displacement_curve.get_limit_value_y(specifier = 'max')
                            displacement_min_max_values.append(biw_accel_max_curve_displacement)
                            biw_accel_min_displacement = biw_accel_unit_y_displacement_curve.get_limit_value_y(specifier = 'min')
                            displacement_min_max_values.append(biw_accel_min_displacement)
                            biw_accel_unit_y_displacement_curve.show()
                            #actiavting plot,getting plot title object and formatting the plot title,yaxis,xaxis options and attributes
                            title = plot2d.Title(plot_id, biw_accel_window_name, page_id)
                            plot.activate()
                            self.activated_plot = plot
                            utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, 1))
                            self.kinematics_curve_format(biw_accel_window_name, plot_id,velocity_min_max_values, displacement_min_max_values, title = "UNIT")
                            if self.twod_images_report_folder is not None:
                                #capturing "UNIT" plot image
                                image_path = os.path.join(self.twod_images_report_folder,biw_accel_window_name+"_"+title.get_text()+".png").replace(" ","_")
                                capture_image_and_resize(image_path,shape.width,shape.height)
                                self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                                self.logger.info("")
                                self.logger.info("CURVES : {},{} | SOURCE PLOT : {} | SOURCE WINDOW : {}".format("*UNIT*Y velocity","*UNIT*Y displacement",title.get_text(),biw_accel_window_name))
                                self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                                self.logger.info("OUTPUT CURVE IMAGES : ")
                                self.logger.info(image_path)
                                self.logger.info("")
                                #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                                picture.crop_left = 0
                                picture.crop_right = 0
                                #reverting 2d data setup
                                biw_accel_unit_y_velocity_curve.hide()
                                biw_accel_unit_y_velocity_curve.hide()
                                self.revert()
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                                self.logger.info("")
                        #image insertion for the shape named "Image 7"
                        elif shape.name == "Image 7":
                            #calling default 2d data setup and getting the Y velocity and Y displacement curve from biw accel window of APLR R plot id 1
                            self.setup()
                            plot_id = 1
                            page_id = 0
                            plot = plot2d.Plot(plot_id, biw_accel_window_name, page_id)
                            velocity_min_max_values = []
                            displacement_min_max_values = []
                            biw_accel_aplr_y_velocity_curve = plot.get_curves('byname', name = '*APLR_R*Y velocity')[0]
                            #collecting min max y values for setting the y axis range of the plot
                            biw_accel_max_velocity = biw_accel_aplr_y_velocity_curve.get_limit_value_y(specifier = 'max')
                            velocity_min_max_values.append(biw_accel_max_velocity)
                            biw_accel_min_velocity = biw_accel_aplr_y_velocity_curve.get_limit_value_y(specifier = 'min')
                            velocity_min_max_values.append(biw_accel_min_velocity)
                            biw_accel_aplr_y_velocity_curve.show()
                            biw_accel_aplr_y_displacement_curve = plot.get_curves('byname', name = '*APLR_R*Y displacement')[0]
                            biw_accel_max_curve_displacement = biw_accel_aplr_y_displacement_curve.get_limit_value_y(specifier = 'max')
                            displacement_min_max_values.append(biw_accel_max_curve_displacement)
                            biw_accel_min_displacement = biw_accel_aplr_y_displacement_curve.get_limit_value_y(specifier = 'min')
                            displacement_min_max_values.append(biw_accel_min_displacement)
                            biw_accel_aplr_y_displacement_curve.show()
                            #actiavting plot,getting plot title object and formatting the plot title,yaxis,xaxis options and attributes
                            title = plot2d.Title(plot_id, biw_accel_window_name, page_id)
                            plot.activate()
                            self.activated_plot = plot
                            utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, 1))
                            self.kinematics_curve_format(biw_accel_window_name, plot_id, velocity_min_max_values, displacement_min_max_values,title="APLR_R")
                            if self.twod_images_report_folder is not None:
                                #capturing "APLR R" plot image
                                image_path = os.path.join(self.twod_images_report_folder,biw_accel_window_name+"_"+title.get_text()+".png").replace(" ","_")
                                capture_image_and_resize(image_path,shape.width,shape.height)
                                self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                                self.logger.info("")
                                self.logger.info("CURVES : {},{} | SOURCE PLOT : {} | SOURCE WINDOW : {}".format("*APLR_R*Y velocity","*APLR_R*Y displacement",title.get_text(),biw_accel_window_name))
                                self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                                self.logger.info("OUTPUT CURVE IMAGES : ")
                                self.logger.info(image_path)
                                self.logger.info("")
                                #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                                picture.crop_left = 0
                                picture.crop_right = 0
                                #reverting 2d data setup
                                biw_accel_aplr_y_displacement_curve.hide()
                                biw_accel_aplr_y_velocity_curve.hide()
                                self.revert()
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                                self.logger.info("")
                        #image insertion for the shape named "Image 8"
                        elif shape.name == "Image 8":
                            #calling default 2d data setup and getting the Y velocity and Y displacement curve from biw accel window of SIS ROW2 R plot id 1
                            self.setup()
                            plot_id = 1
                            page_id = 0
                            plot = plot2d.Plot(plot_id, biw_accel_window_name, page_id)
                            velocity_min_max_values = []
                            displacement_min_max_values = []
                            biw_accel_sis_row_2_y_velocity_curve = plot.get_curves('byname', name = '*SIS_ROW2_R*Y velocity')[0]
                            #collecting min max y values for setting the y axis range of the plot
                            biw_accel_max_velocity = biw_accel_sis_row_2_y_velocity_curve.get_limit_value_y(specifier = 'max')
                            velocity_min_max_values.append(biw_accel_max_velocity)
                            biw_accel_min_velocity = biw_accel_sis_row_2_y_velocity_curve.get_limit_value_y(specifier = 'min')
                            velocity_min_max_values.append(biw_accel_min_velocity)
                            biw_accel_sis_row_2_y_velocity_curve.show()
                            biw_accel_sis_row_2_y_displacement_curve = plot.get_curves('byname', name = '*SIS_ROW2_R*Y displacement')[0]
                            biw_accel_max_curve_displacement = biw_accel_sis_row_2_y_displacement_curve.get_limit_value_y(specifier = 'max')
                            displacement_min_max_values.append(biw_accel_max_curve_displacement)
                            biw_accel_min_displacement = biw_accel_sis_row_2_y_displacement_curve.get_limit_value_y(specifier = 'min')
                            displacement_min_max_values.append(biw_accel_min_displacement)
                            biw_accel_sis_row_2_y_displacement_curve.show()
                            title = plot2d.Title(plot_id, biw_accel_window_name, page_id)
                            #actiavting plot,getting plot title object and formatting the plot title,yaxis,xaxis options and attributes
                            plot.activate()
                            self.activated_plot = plot
                            utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, 1))
                            self.kinematics_curve_format(biw_accel_window_name, plot_id,velocity_min_max_values, displacement_min_max_values,title="SS_BP_R")
                            if self.twod_images_report_folder is not None:
                                #capturing "SIS ROW2 R" plot image
                                image_path = os.path.join(self.twod_images_report_folder,biw_accel_window_name+"_"+title.get_text()+".png").replace(" ","_")
                                capture_image_and_resize(image_path,shape.width,shape.height)
                                self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                                self.logger.info("")
                                self.logger.info("CURVES : {},{} | SOURCE PLOT : {} | SOURCE WINDOW : {}".format("*SIS_ROW2_R*Y velocity","*SIS_ROW2_R*Y displacement",title.get_text(),biw_accel_window_name))
                                self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                                self.logger.info("OUTPUT CURVE IMAGES : ")
                                self.logger.info(image_path)
                                self.logger.info("")
                                #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                                picture.crop_left = 0
                                picture.crop_right = 0
                                #reverting 2d data setup
                                biw_accel_sis_row_2_y_displacement_curve.hide()
                                biw_accel_sis_row_2_y_velocity_curve.hide()
                                self.revert()
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                                self.logger.info("")
                        #image insertion for the shape named "Image 9"
                        elif shape.name == "Image 9":
                            #calling default 2d data setup and getting the Y velocity and Y displacement curve from biw accel window of SS_R_RR_TOP_G_Y1 plot id 1
                            self.setup()
                            plot_id = 1
                            page_id = 0
                            plot = plot2d.Plot(plot_id, biw_accel_window_name, page_id)
                            velocity_min_max_values = []
                            displacement_min_max_values = []
                            biw_accel_ss_rr_top_y_velocity_curve = plot.get_curves('byname', name = '*SS_R_RR_TOP_G_Y1*Y velocity')[0]
                            #collecting min max y values for setting the y axis range of the plot
                            biw_accel_max_velocity = biw_accel_ss_rr_top_y_velocity_curve.get_limit_value_y(specifier = 'max')
                            velocity_min_max_values.append(biw_accel_max_velocity)
                            biw_accel_min_velocity = biw_accel_ss_rr_top_y_velocity_curve.get_limit_value_y(specifier = 'min')
                            velocity_min_max_values.append(biw_accel_min_velocity)
                            biw_accel_ss_rr_top_y_velocity_curve.show()
                            biw_accel_ss_rr_top_y_displacement_curve = plot.get_curves('byname', name = '*SS_R_RR_TOP_G_Y1*Y displacement')[0]
                            biw_accel_max_curve_displacement = biw_accel_ss_rr_top_y_displacement_curve.get_limit_value_y(specifier = 'max')
                            displacement_min_max_values.append(biw_accel_max_curve_displacement)
                            biw_accel_min_displacement = biw_accel_ss_rr_top_y_displacement_curve.get_limit_value_y(specifier = 'min')
                            displacement_min_max_values.append(biw_accel_min_displacement)
                            biw_accel_ss_rr_top_y_displacement_curve.show()
                            title = plot2d.Title(plot_id, biw_accel_window_name, page_id)
                            #actiavting plot,getting plot title object and formatting the plot title,yaxis,xaxis options and attributes
                            plot.activate()
                            self.activated_plot = plot
                            utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, 1))
                            self.kinematics_curve_format(biw_accel_window_name, plot_id,velocity_min_max_values, displacement_min_max_values, title="SS_R_RR_TOP")
                            if self.twod_images_report_folder is not None:
                                #capturing "SS_R_RR_TOP_G_Y1" plot image
                                image_path = os.path.join(self.twod_images_report_folder,biw_accel_window_name+"_"+title.get_text()+".png").replace(" ","_")
                                capture_image_and_resize(image_path,shape.width,shape.height)
                                self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                                self.logger.info("")
                                self.logger.info("CURVES : {},{} | SOURCE PLOT : {} | SOURCE WINDOW : {}".format("*SS_R_RR_TOP_G_Y1*Y velocity","*SS_R_RR_TOP_G_Y1*Y displacement",title.get_text(),biw_accel_window_name))
                                self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                                self.logger.info("OUTPUT CURVE IMAGES : ")
                                self.logger.info(image_path)
                                self.logger.info("")
                                #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                                picture.crop_left = 0
                                picture.crop_right = 0
                                #reverting 2d data setup
                                biw_accel_ss_rr_top_y_displacement_curve.hide()
                                biw_accel_ss_rr_top_y_velocity_curve.hide()
                                self.revert()
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                                self.logger.info("")
                        #image insertion for the shape named "Image 9"
                        elif shape.name == "Image 10":
                            #calling default 2d data setup and getting the Y velocity and Y displacement curve from biw accel window of UPR plot id 2
                            self.setup()
                            plot_id = 2
                            page_id = 0
                            plot = plot2d.Plot(plot_id, biw_accel_window_name, page_id)
                            velocity_min_max_values = []
                            displacement_min_max_values = []
                            biw_accel_upr_y_velocity_curve = plot.get_curves('byname', name = '*MID_UPR*Y velocity')[0]
                            #collecting min max y values for setting the y axis range of the plot
                            biw_accel_max_velocity = biw_accel_upr_y_velocity_curve.get_limit_value_y(specifier = 'max')
                            velocity_min_max_values.append(biw_accel_max_velocity)
                            biw_accel_min_velocity = biw_accel_upr_y_velocity_curve.get_limit_value_y(specifier = 'min')
                            velocity_min_max_values.append(biw_accel_min_velocity)
                            biw_accel_upr_y_velocity_curve.show()
                            biw_accel_upr_y_displacement_curve = plot.get_curves('byname', name = '*MID_UPR*Y displacement')[0]
                            biw_accel_max_curve_displacement = biw_accel_upr_y_displacement_curve.get_limit_value_y(specifier = 'max')
                            displacement_min_max_values.append(biw_accel_max_curve_displacement)
                            biw_accel_min_displacement = biw_accel_upr_y_displacement_curve.get_limit_value_y(specifier = 'min')
                            displacement_min_max_values.append(biw_accel_min_displacement)
                            biw_accel_upr_y_displacement_curve.show()
                            title = plot2d.Title(plot_id, biw_accel_window_name, page_id)
                            #actiavting plot,getting plot title object and formatting the plot title,yaxis,xaxis options and attributes
                            plot.activate()
                            self.activated_plot = plot
                            utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, 1))
                            self.kinematics_curve_format(biw_accel_window_name, plot_id,velocity_min_max_values, displacement_min_max_values, title="BPLR_MID_R")
                            if self.twod_images_report_folder is not None:
                                #capturing "UPR" plot image
                                image_path = os.path.join(self.twod_images_report_folder,biw_accel_window_name+"_"+title.get_text()+".png").replace(" ","_")
                                capture_image_and_resize(image_path,shape.width,shape.height)
                                self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                                self.logger.info("")
                                self.logger.info("CURVES : {},{} | SOURCE PLOT : {} | SOURCE WINDOW : {}".format("*MID_UPR*Y velocity","*MID_UPR*Y displacement",title.get_text(),biw_accel_window_name))
                                self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                                self.logger.info("OUTPUT CURVE IMAGES : ")
                                self.logger.info(image_path)
                                self.logger.info("")
                                #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                                picture.crop_left = 0
                                picture.crop_right = 0
                                #reverting 2d data setup
                                biw_accel_upr_y_displacement_curve.hide()
                                biw_accel_upr_y_velocity_curve.hide()
                                self.revert()
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                                self.logger.info("")
                else:
                    self.logger.info("ERROR : 2D METADB does not contain 'BIW - Accel' window. Please update.")
            else:
                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.biw_accel_window_key))
            # Duplicating Oval Shapes
            for shape in oval_shapes:
                clone_shape(shape)
            endtime = datetime.now()
        except Exception as e:
            self.logger.exception("Error while seeding data into biw kinematics slide:\n{}".format(e))
            self.logger.info("")
            return 1
        self.logger.info("Completed seeding data into biw kinematics slide")
        self.logger.info("Time Taken : {}".format(endtime - starttime))
        self.logger.info("")

        return 0

    def revert(self,format_type = "2d"):
        """
        This method is used to revert data for meta windows based on 3d and 2d formats

        Args:
            format_type (str, optional): format type. Defaults to "2d".

        Returns:
            int: 0 Always for Sucess.1 for Failure.
        """
        try:
            if format_type == "2d":
                self.activated_plot.deactivate()
                utils.MetaCommand('xyplot rlayout "{}" {}'.format(self.biw_accel_window_name, self.biw_accel_window_layout))
            else:
                utils.MetaCommand('0:options state original')
                utils.MetaCommand('options fringebar on')
                utils.MetaCommand('grstyle scalarfringe enable')
        except Exception as e:
            return 1

        return 0
