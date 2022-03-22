# PYTHON script
"""
This script is used for all the automation process of Body In White Floor Deformation and spotweld failure slide of thesis report.
"""

import os
import logging
from datetime import datetime

from meta import utils
from meta import windows
from meta import plot2d,models

from src.meta_utilities import capture_image
from src.meta_utilities import visualize_3d_critical_section
from src.meta_utilities import capture_image_and_resize
from src.meta_utilities import visualize_annotation
from src.general_utilities import closest

class BIWFloorDeformationAndSpotWeldFailureSlide():
    """
       This class is used to automate the biw floor deformation and spotweld failure slide of thesis report.

        Args:
            slide (object): biw floor deformation and spotweld failure pptx slide object.
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
        self.logger = logging.getLogger("side_crash_logger")
        self.visible_parts = None

    def edit(self):
        """
        This method is used to iterate all the shapes of the biw floor deformation and spotweld failure slide and insert respective data.

        Returns:
            int: 0 Always for Sucess,1 for Failure.
        """
        try:
            self.logger.info("Started seeding data into biw floor deformation spotweld failure slide")
            self.logger.info("")
            starttime = datetime.now()
            #iterating through the shapes of the biw floor deformation and spotweld failure slide
            for shape in self.shapes:
                #image insertion for the shape named "Image 3"
                if shape.name == "Image 3":
                    #visualizing and capturing image of "f21_front_floor" critical part set at peak state with deformation
                    utils.MetaCommand('window maximize {}'.format(self.general_input.threed_window_name))
                    utils.MetaCommand('0:options state variable "serial=1"')
                    utils.MetaCommand('grstyle scalarfringe enable')
                    utils.MetaCommand('options fringebar off')
                    data = self.metadb_3d_input.critical_sections["f21_front_floor"]
                    visualize_3d_critical_section(data)
                    utils.MetaCommand('color pid transparency reset act')
                    image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_"+"F21_FRONT_FLOOR_AT_PEAK_STATE_WITH_DEFORMATION"+".png").replace(" ","_")
                    capture_image(image_path,self.general_input.threed_window_name,shape.width,shape.height)
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
                    #reverting visual changes
                    utils.MetaCommand('options fringebar on')
                    utils.MetaCommand('0:options state original"')
                #image insertion for the shape named "Image 4"
                elif shape.name == "Image 4":
                    #visualizing and capturing image of "f21_front_floor" critical part set at peak state without deformation
                    utils.MetaCommand('window maximize {}'.format(self.general_input.threed_window_name))
                    utils.MetaCommand('0:options state variable "serial=1"')
                    utils.MetaCommand('grstyle scalarfringe enable')
                    utils.MetaCommand('options fringebar off')
                    utils.MetaCommand('grstyle deform off')
                    data = self.metadb_3d_input.critical_sections["f21_front_floor"]
                    visualize_3d_critical_section(data)
                    utils.MetaCommand('color pid transparency reset act')
                    image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_"+"F21_FRONT_FLOOR_AT_PEAK_STATE_WITHOUT_DEFORMATION"+".png").replace(" ","_")
                    capture_image(image_path,self.general_input.threed_window_name,shape.width,shape.height,view = "btm")
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
                    #reverting visual changes
                    utils.MetaCommand('options fringebar on')
                    utils.MetaCommand('grstyle deform on')
                    utils.MetaCommand('0:options state original"')
                #image insertion for the shape named "Image 2"
                elif shape.name == "Image 2":
                    #capturing fringe bar of metapost window
                    utils.MetaCommand('window maximize {}'.format(self.general_input.threed_window_name))
                    utils.MetaCommand('0:options state variable "serial=1"')
                    utils.MetaCommand('add all')
                    utils.MetaCommand('add invert')
                    utils.MetaCommand('options fringebar on')
                    image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_"+"FRINGE_BAR"+".png").replace(" ","_")
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
                #image insertion for the shape named "Image 1"
                elif shape.name == "Image 1":
                    #maximizing biw stiff ring deformation window and getting its plot layout number
                    biw_stiff_ring_def_window_name = self.general_input.biw_stiff_ring_deformation_name
                    win = windows.Window(str(biw_stiff_ring_def_window_name), page_id=0)
                    layout = win.get_plot_layout()
                    utils.MetaCommand('window maximize "{}"'.format(biw_stiff_ring_def_window_name))
                    final_time_curve_name = "SIDE_SILL_{}MS".format(round(float(self.general_input.survival_space_final_time)))
                    initial_time_curve_name = "SIDE_SILL_0MS"
                    #showing side sill initial,final and peak state curves
                    plot_id = 1
                    page_id = 0
                    plot = plot2d.Plot(plot_id, biw_stiff_ring_def_window_name, page_id)
                    title = plot.get_title()
                    plot.activate()
                    final_curve = plot2d.CurvesByName(biw_stiff_ring_def_window_name, final_time_curve_name, 0)[0]
                    final_curve.show()
                    initial_curve = plot2d.CurvesByName(biw_stiff_ring_def_window_name, initial_time_curve_name, 0)[0]
                    initial_curve.show()
                    peak_time_value = self.general_input.peak_time_display_value
                    peak_time_value = peak_time_value.split(".")[0]
                    roof_line_curves = plot.get_curves('all')
                    roof_line_cuves_list = list()
                    for each_roof_line_curves in roof_line_curves:
                        ms = each_roof_line_curves.name.split("_")[2]
                        if 'MS' in ms:
                            ms_replacing = ms.replace('MS',"")
                            roof_line_cuves_list.append(int(ms_replacing))
                    peak_curve_value = closest(roof_line_cuves_list, int(peak_time_value))
                    peak_curve_value = plot.get_curves('byname', name ="SIDE_SILL_"+str(peak_curve_value)+"MS")
                    peak_curve = plot2d.CurvesByName(biw_stiff_ring_def_window_name, peak_curve_value[0].name, 0)[0]
                    peak_curve.show()
                    #custom formating of visible initial,peak and final state curves
                    utils.MetaCommand('xyplot plotactive "{}" 1'.format(biw_stiff_ring_def_window_name))
                    utils.MetaCommand('xyplot rlayout "{}" 1'.format(biw_stiff_ring_def_window_name))
                    utils.MetaCommand('xyplot curve visible and "{}" {} {},{}'.format(biw_stiff_ring_def_window_name,initial_curve.id,peak_curve.id, final_curve.id))
                    utils.MetaCommand('xyplot curve set style "{}" {} 9'.format(biw_stiff_ring_def_window_name, initial_curve.id))
                    utils.MetaCommand('xyplot curve set style "{}" {} 5'.format(biw_stiff_ring_def_window_name,peak_curve.id))
                    utils.MetaCommand('xyplot curve select "{}" all'.format(biw_stiff_ring_def_window_name))
                    utils.MetaCommand('xyplot axisoptions yaxis active "{}" 1 0'.format(biw_stiff_ring_def_window_name))
                    utils.MetaCommand('xyplot axisoptions axyrange "{}" 1 0 -805 -770'.format(biw_stiff_ring_def_window_name))
                    utils.MetaCommand('xyplot axisoptions ylabel font "{}" 1 "Arial,25,-1,5,75,0,0,0,0,0"'.format(biw_stiff_ring_def_window_name))
                    utils.MetaCommand('xyplot axisoptions labels yfont "{}" 1 "Arial,25,-1,5,75,0,0,0,0,0"'.format(biw_stiff_ring_def_window_name))
                    utils.MetaCommand('xyplot axisoptions yaxis deactive "{}" 1 0'.format(biw_stiff_ring_def_window_name))
                    utils.MetaCommand('xyplot axisoptions xaxis active "{}" 1 0'.format(biw_stiff_ring_def_window_name))
                    utils.MetaCommand('xyplot axisoptions xlabel font "{}" 1 "Arial,25,-1,5,75,0,0,0,0,0"'.format(biw_stiff_ring_def_window_name))
                    utils.MetaCommand('xyplot axisoptions labels xfont "{}" 1 "Arial,25,-1,5,75,0,0,0,0,0"'.format(biw_stiff_ring_def_window_name))
                    utils.MetaCommand('xyplot plotoptions title font "{}" 1 "Arial,25,-1,5,75,0,0,0,0,0"'.format(biw_stiff_ring_def_window_name))
                    #capturing plot image
                    image_path = os.path.join(self.twod_images_report_folder,biw_stiff_ring_def_window_name+"_"+title.get_text()+".png").replace("&","and").replace(" ","_")
                    capture_image_and_resize(image_path,shape.width,shape.height)
                    self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                    self.logger.info("")
                    self.logger.info("CURVES : {} | SOURCE PLOT : {} | SOURCE WINDOW : {}".format("ROOF_LINE_0MS,ROOF_LINE_{}MS,ROOF_LINE_{}MS",self.general_input.front_abdomen_intrusion_curve_key,biw_stiff_ring_def_window_name))
                    self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                    self.logger.info("OUTPUT CURVE IMAGES : ")
                    self.logger.info(image_path)
                    self.logger.info("")
                    #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                    picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                    picture.crop_left = 0
                    picture.crop_right = 0
                    #reverting the layoout of the window
                    plot.deactivate()
                    utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_stiff_ring_def_window_name, layout))
                elif shape.name == "Image 5":
                    #visualizing and capturing "f21_front_floor" critical part set image with spotweld failure annotations
                    utils.MetaCommand('window maximize {}'.format(self.general_input.threed_window_name))
                    utils.MetaCommand('0:options state original')
                    utils.MetaCommand('options fringebar off')
                    data = self.metadb_3d_input.critical_sections["f21_front_floor"]
                    visualize_3d_critical_section(data)
                    visualize_annotation(self.metadb_3d_input.spotweld_clusters,self.general_input.binout_directory)
                    image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_"+"F21_FRONT_FLOOR_SPOTWELD_FAILURE"+".png").replace(" ","_")
                    capture_image(image_path,self.general_input.threed_window_name,shape.width,shape.height,view = "top")
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
            endtime = datetime.now()
        except Exception as e:
            self.logger.exception("Error while seeding data into biw floor deformation spotweld failure slide:\n{}".format(e))
            self.logger.info("")
            return 1
        self.logger.info("Completed seeding data into biw floor deformation spotweld failure slide")
        self.logger.info("Time Taken : {}".format(endtime - starttime))
        self.logger.info("")

        return 0
