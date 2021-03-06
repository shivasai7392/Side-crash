# PYTHON script
"""
This script is used for all the automation process of Body in white stiff ring deformation slide of thesis report.
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
from src.general_utilities import closest
from src.metadb_info import GeneralVarInfo

class BIWStiffRingDeformationSlide():
    """
       This class is used to automate the BIW stiff ring deformation slide of thesis report.

        Args:
            slide (object): biw stiff ring deformation slide pptx slide object.
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

    def edit(self):
        """
        This method is used to iterate all the shapes of the biw stiff ring deformation slide and insert respective data.

        Returns:
            int: 0 Always for Sucess,1 for Failure.
        """
        try:
            self.logger.info("Started seeding data into biw stiff ring deformation slide")
            self.logger.info("")
            starttime = datetime.now()
            #iterating through the shapes of the  biw stiff ring deformation slide
            for shape in self.shapes:
                #image insertion for the shape named "Image 6"
                if shape.name == "Image 6":
                    if self.general_input.biw_stiff_ring_deformation_name not in ["null","none",""]:
                        biw_stiff_ring_deformation_name = self.general_input.biw_stiff_ring_deformation_name
                        biw_stiff_ring_deformation_obj = windows.WindowByName(biw_stiff_ring_deformation_name)
                        if biw_stiff_ring_deformation_obj:
                            #maximizing biw stiff ring deformation window and getting its plot layout number
                            layout = biw_stiff_ring_deformation_obj.get_plot_layout()
                            utils.MetaCommand('window maximize "{}"'.format(biw_stiff_ring_deformation_name))
                            plot_id = 0
                            page_id=0
                            plot = plot2d.Plot(plot_id, biw_stiff_ring_deformation_name, page_id)
                            title = plot.get_title()
                            plot.activate()
                            if self.general_input.survival_space_final_time not in ["null","none",""]:
                                #showing roof line initial,final and peak state curves
                                final_time_curve_name = "ROOF_LINE_{}MS".format(round(float(self.general_input.survival_space_final_time)))
                                final_curves = plot2d.CurvesByName(biw_stiff_ring_deformation_name, final_time_curve_name, 0)
                                if final_curves:
                                    final_curves[0].show()
                                else:
                                    final_curves = None
                                    self.logger.info("ERROR : Side sill & Roof intrusion window does not contain '{}' curve from META 2D variable {}. Please update.".format(final_time_curve_name,GeneralVarInfo.survival_space_final_time_key))
                                    self.logger.info("")
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.survival_space_final_time_key))
                                self.logger.info("")
                            initial_time_curve_name = "ROOF_LINE_0MS"
                            initial_curve = plot2d.CurvesByName(biw_stiff_ring_deformation_name, initial_time_curve_name, 0)[0]
                            initial_curve.show()
                            roof_line_curves = plot.get_curves('all')
                            roof_line_cuves_list = list()
                            for each_roof_line_curves in roof_line_curves:
                                ms = each_roof_line_curves.name.split("_")[2]
                                if 'MS' in ms:
                                    ms_replacing = ms.replace('MS',"")
                                    roof_line_cuves_list.append(int(ms_replacing))
                            if self.general_input.peak_time_display_value not in ["null","none",""]:
                                peak_time_value = self.general_input.peak_time_display_value
                                peak_time_value = peak_time_value.split(".")[0]
                                peak_curve_value = closest(roof_line_cuves_list, int(peak_time_value))
                                peak_curves = plot2d.CurvesByName(biw_stiff_ring_deformation_name, "ROOF_LINE_"+str(peak_curve_value)+"MS", 0)
                                if peak_curves:
                                    peak_curves[0].show()
                                else:
                                    peak_curves = None
                                    self.logger.info("ERROR : Side sill & Roof intrusion window does not contain '{}' curve from META 2D variable {}. Please update.".format("ROOF_LINE_"+str(peak_curve_value[0])+"MS",GeneralVarInfo.peak_time_display_key))
                                    self.logger.info("")
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.peak_time_display_key))
                                self.logger.info("")
                            #custom formating of visible initial,peak and final state curves
                            utils.MetaCommand('xyplot plotactive "{}" 0'.format(biw_stiff_ring_deformation_name))
                            utils.MetaCommand('xyplot rlayout "{}" 1'.format(biw_stiff_ring_deformation_name))
                            if peak_curves and final_curves:
                                utils.MetaCommand('xyplot curve visible and "{}" {} {},{}'.format(biw_stiff_ring_deformation_name,initial_curve.id,peak_curves[0].id, final_curves[0].id))
                            utils.MetaCommand('xyplot curve set style "{}" {} 9'.format(biw_stiff_ring_deformation_name, initial_curve.id))
                            if peak_curves:
                                utils.MetaCommand('xyplot curve set style "{}" {} 5'.format(biw_stiff_ring_deformation_name,peak_curves[0].id))
                            utils.MetaCommand('xyplot curve select "{}" all'.format(biw_stiff_ring_deformation_name))
                            utils.MetaCommand('xyplot axisoptions yaxis active "{}" 0 0'.format(biw_stiff_ring_deformation_name))
                            utils.MetaCommand('xyplot axisoptions ylabel font "{}" 0 "Arial,25,-1,5,75,0,0,0,0,0"'.format(biw_stiff_ring_deformation_name))
                            utils.MetaCommand('xyplot axisoptions labels yfont "{}" 0 "Arial,25,-1,5,75,0,0,0,0,0"'.format(biw_stiff_ring_deformation_name))
                            utils.MetaCommand('xyplot axisoptions yaxis deactive "{}" 0 0'.format(biw_stiff_ring_deformation_name))
                            utils.MetaCommand('xyplot axisoptions xaxis active "{}" 0 0'.format(biw_stiff_ring_deformation_name))
                            utils.MetaCommand('xyplot axisoptions xlabel font "{}" 0 "Arial,25,-1,5,75,0,0,0,0,0"'.format(biw_stiff_ring_deformation_name))
                            utils.MetaCommand('xyplot axisoptions labels xfont "{}" 0 "Arial,25,-1,5,75,0,0,0,0,0"'.format(biw_stiff_ring_deformation_name))
                            utils.MetaCommand('xyplot plotoptions title font "{}" 0 "Arial,25,-1,5,75,0,0,0,0,0"'.format(biw_stiff_ring_deformation_name))
                            if self.twod_images_report_folder is not None:
                                #capturing plot image
                                image_path = os.path.join(self.twod_images_report_folder,biw_stiff_ring_deformation_name+"_"+title.get_text()+".png").replace("&","and").replace(" ","_")
                                capture_image_and_resize(image_path,shape.width,shape.height)
                                self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                                self.logger.info("")
                                self.logger.info("CURVES : {} | SOURCE PLOT : {} | SOURCE WINDOW : {}".format("ROOF_LINE_0MS,ROOF_LINE_{}MS,ROOF_LINE_{}MS",self.general_input.front_abdomen_intrusion_curve_key,biw_stiff_ring_deformation_name))
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
                                utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_stiff_ring_deformation_name, layout))
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                                self.logger.info("")
                        else:
                            self.logger.info("ERROR : 2D METADB does not contain 'Side sill & Roof intrusion'. Please update.")
                    else:
                        self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.biw_stiff_ring_deformation_key))
                #image insertion for the shape named "Image 7"
                elif shape.name == "Image 7":
                    if self.general_input.biw_stiff_ring_deformation_name not in ["null","none",""]:
                        biw_stiff_ring_def_window_name = self.general_input.biw_stiff_ring_deformation_name
                        biw_stiff_ring_def_window_obj = windows.Window(str(biw_stiff_ring_def_window_name), page_id=0)
                        if biw_stiff_ring_def_window_obj:
                            layout = biw_stiff_ring_def_window_obj.get_plot_layout()
                            utils.MetaCommand('window maximize "{}"'.format(biw_stiff_ring_def_window_name))
                            plot_id = 1
                            page_id = 0
                            plot = plot2d.Plot(plot_id, biw_stiff_ring_def_window_name, page_id)
                            title = plot.get_title()
                            plot.activate()
                            if self.general_input.survival_space_final_time not in ["null","none",""]:
                                final_time_curve_name = "SIDE_SILL_{}MS".format(round(float(self.general_input.survival_space_final_time)))
                                final_curves = plot2d.CurvesByName(biw_stiff_ring_def_window_name, final_time_curve_name, 0)
                                if final_curves:
                                    final_curves[0].show()
                                else:
                                    self.logger.info("ERROR : Side sill & Roof intrusion window does not contain '{}' curve from META 2D variable {}. Please update.".format(final_time_curve_name,GeneralVarInfo.survival_space_final_time_key))
                                    self.logger.info("")
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.survival_space_final_time_key))
                                self.logger.info("")
                            initial_time_curve_name = "SIDE_SILL_0MS"
                            #showing side sill initial,final and peak state curves
                            initial_curve = plot2d.CurvesByName(biw_stiff_ring_def_window_name, initial_time_curve_name, 0)[0]
                            initial_curve.show()
                            roof_line_curves = plot.get_curves('all')
                            roof_line_cuves_list = list()
                            for each_roof_line_curves in roof_line_curves:
                                ms = each_roof_line_curves.name.split("_")[2]
                                if 'MS' in ms:
                                    ms_replacing = ms.replace('MS',"")
                                    roof_line_cuves_list.append(int(ms_replacing))
                            if self.general_input.peak_time_display_value not in ["null","none",""]:
                                peak_time_value = self.general_input.peak_time_display_value
                                peak_time_value = peak_time_value.split(".")[0]
                                peak_curve_value = closest(roof_line_cuves_list, int(peak_time_value))
                                peak_curves = plot2d.CurvesByName(biw_stiff_ring_def_window_name, "SIDE_SILL_"+str(peak_curve_value)+"MS", 0)
                                if peak_curves:
                                    peak_curves[0].show()
                                else:
                                    peak_curves = None
                                    self.logger.info("ERROR : Side sill & Roof intrusion window does not contain '{}' curve from META 2D variable {}. Please update.".format("ROOF_LINE_"+str(peak_curve_value[0])+"MS",GeneralVarInfo.peak_time_display_key))
                                    self.logger.info("")

                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.peak_time_display_key))
                                self.logger.info("")
                            #custom formating of visible initial,peak and final state curves
                            utils.MetaCommand('xyplot plotactive "{}" 1'.format(biw_stiff_ring_def_window_name))
                            utils.MetaCommand('xyplot rlayout "{}" 1'.format(biw_stiff_ring_def_window_name))
                            if peak_curves and final_curves:
                                utils.MetaCommand('xyplot curve visible and "{}" {} {},{}'.format(biw_stiff_ring_def_window_name,initial_curve.id,peak_curves[0].id, final_curves[0].id))
                                utils.MetaCommand('xyplot curve set style "{}" {} 5'.format(biw_stiff_ring_def_window_name,peak_curves[0].id))
                            utils.MetaCommand('xyplot curve set style "{}" {} 9'.format(biw_stiff_ring_def_window_name, initial_curve.id))
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
                            if self.twod_images_report_folder is not None:
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
                            else:
                                self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                                self.logger.info("")
                        else:
                            self.logger.info("ERROR : 2D METADB does not contain 'Side sill & Roof intrusion' window. Please update.")
                    else:
                        self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.biw_stiff_ring_deformation_key))
                #image insertion for the shape named "Image 7"
                elif shape.name == "Image 4":
                    #visualizing "f21_upb_inner" critical part set to capture image at peak state with deformation
                    data = self.metadb_3d_input.critical_sections["f21_upb_inner"]
                    visualize_3d_critical_section(data,name = "f21_upb_inner")
                    utils.MetaCommand('color pid transparency reset act')
                    utils.MetaCommand('grstyle scalarfringe enable')
                    utils.MetaCommand('0:options state variable "serial=1"')
                    utils.MetaCommand('options fringebar off')
                    if self.threed_images_report_folder is not None:
                        image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_"+"F21_UPB_INNER_AT_PEAK_STATE_WITH_DEFORMATION"+".png").replace(" ","_")
                        capture_image(image_path,self.general_input.threed_window_name,shape.width,shape.height, view = "right",transparent=True)
                        self.logger.info("--- 3D MODEL IMAGE GENERATOR")
                        self.logger.info("")
                        self.logger.info("SOURCE WINDOW : {} ".format(self.general_input.threed_window_name))
                        self.logger.info("SOURCE MODEL : 0")
                        self.logger.info("STATE : PEAK STATE")
                        self.logger.info("PID NAME SHOW FILTER : {} ".format(data["hes"] if "hes" in data.keys() else "null"))
                        self.logger.info("ADDITIONAL PID'S SHOWN : {} ".format(data["hes_exceptions"] if "hes_exceptions" in data.keys() else "null"))
                        #self.logger.info("PID NAME ERASE FILTER : {} ".format(data["hes_exceptions"] if "hes_exceptions" in data.keys() else "null"))
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
                        transparent_image_path = image_path.replace(".png","")+"_transparent.png"
                        picture = self.shapes.add_picture(transparent_image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                        picture.crop_left = 0
                        picture.crop_right = 0
                        #removing transparent image
                        os.remove(transparent_image_path)
                    else:
                        self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                        self.logger.info("")
                #image insertion for the shape named "Image 7"
                elif shape.name == "Image 5":
                    #visualizing "f21_upb_inner" critical part set to capture image at peak state without deformation
                    data = self.metadb_3d_input.critical_sections["f21_upb_inner"]
                    visualize_3d_critical_section(data,name = "f21_upb_inner")
                    utils.MetaCommand('color pid transparency reset act')
                    utils.MetaCommand('grstyle scalarfringe enable')
                    utils.MetaCommand('0:options state variable "serial=1"')
                    utils.MetaCommand('options fringebar off')
                    if self.threed_images_report_folder is not None:
                        image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_"+"F21_UPB_INNER_AT_PEAK_STATE_WITHOUT_DEFORMATION"+".png").replace(" ","_")
                        capture_image(image_path,self.general_input.threed_window_name,shape.width,shape.height, view = "right",transparent=True)
                        self.logger.info("--- 3D MODEL IMAGE GENERATOR")
                        self.logger.info("")
                        self.logger.info("SOURCE WINDOW : {} ".format(self.general_input.threed_window_name))
                        self.logger.info("SOURCE MODEL : 0")
                        self.logger.info("STATE : PEAK STATE")
                        self.logger.info("PID NAME SHOW FILTER : {} ".format(data["hes"] if "hes" in data.keys() else "null"))
                        self.logger.info("ADDITIONAL PID'S SHOWN : {} ".format(data["hes_exceptions"] if "hes_exceptions" in data.keys() else "null"))
                        #self.logger.info("PID NAME ERASE FILTER : {} ".format(data["hes_exceptions"] if "hes_exceptions" in data.keys() else "null"))
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
                        transparent_image_path = image_path.replace(".png","")+"_transparent.png"
                        picture = self.shapes.add_picture(transparent_image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                        picture.crop_left = 0
                        picture.crop_right = 0
                        #removing transparent image
                        os.remove(transparent_image_path)
                        #reverting back deformation
                        utils.MetaCommand('grstyle deform on')
                    else:
                        self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                        self.logger.info("")
                #image insertion for the shape named "Image 2"
                elif shape.name == "Image 2":
                    #visualizing "f21_upb_outer" critical part set to capture image at peak state with deformation
                    data = self.metadb_3d_input.critical_sections["f21_upb_outer"]
                    visualize_3d_critical_section(data,name = "f21_upb_outer")
                    utils.MetaCommand('color pid transparency reset act')
                    utils.MetaCommand('grstyle scalarfringe enable')
                    utils.MetaCommand('0:options state variable "serial=1"')
                    utils.MetaCommand('options fringebar off')
                    utils.MetaCommand('grstyle deform on')
                    if self.threed_images_report_folder is not None:
                        image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_"+"F21_UPB_OUTER_AT_PEAK_STATE_WITH_DEFORMATION"+".png").replace(" ","_")
                        capture_image(image_path,self.general_input.threed_window_name,shape.width,shape.height, view = "right",transparent=True)
                        self.logger.info("--- 3D MODEL IMAGE GENERATOR")
                        self.logger.info("")
                        self.logger.info("SOURCE WINDOW : {} ".format(self.general_input.threed_window_name))
                        self.logger.info("SOURCE MODEL : 0")
                        self.logger.info("STATE : PEAK STATE")
                        self.logger.info("PID NAME SHOW FILTER : {} ".format(data["hes"] if "hes" in data.keys() else "null"))
                        self.logger.info("ADDITIONAL PID'S SHOWN : {} ".format(data["hes_exceptions"] if "hes_exceptions" in data.keys() else "null"))
                        #self.logger.info("PID NAME ERASE FILTER : {} ".format(data["hes_exceptions"] if "hes_exceptions" in data.keys() else "null"))
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
                    #visualizing "f21_upb_outer" critical part set to capture image at peak state without deformation
                    data = self.metadb_3d_input.critical_sections["f21_upb_outer"]
                    visualize_3d_critical_section(data,name = "f21_upb_outer")
                    utils.MetaCommand('color pid transparency reset act')
                    utils.MetaCommand('grstyle scalarfringe enable')
                    utils.MetaCommand('0:options state variable "serial=1"')
                    utils.MetaCommand('options fringebar off')
                    utils.MetaCommand('grstyle deform off')
                    if self.threed_images_report_folder is not None:
                        image_path = os.path.join(self.threed_images_report_folder,self.general_input.threed_window_name+"_"+"F21_UPB_OUTER_AT_PEAK_STATE_WITHOUT_DEFORMATION"+".png").replace(" ","_")
                        capture_image(image_path,self.general_input.threed_window_name,shape.width,shape.height, view = "right",transparent=True)
                        self.logger.info("--- 3D MODEL IMAGE GENERATOR")
                        self.logger.info("")
                        self.logger.info("SOURCE WINDOW : {} ".format(self.general_input.threed_window_name))
                        self.logger.info("SOURCE MODEL : 0")
                        self.logger.info("STATE : PEAK STATE")
                        self.logger.info("PID NAME SHOW FILTER : {} ".format(data["hes"] if "hes" in data.keys() else "null"))
                        self.logger.info("ADDITIONAL PID'S SHOWN : {} ".format(data["hes_exceptions"] if "hes_exceptions" in data.keys() else "null"))
                        #self.logger.info("PID NAME ERASE FILTER : {} ".format(data["hes_exceptions"] if "hes_exceptions" in data.keys() else "null"))
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
                        transparent_image_path = image_path.replace(".png","")+"_transparent.png"
                        picture = self.shapes.add_picture(transparent_image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                        picture.crop_left = 0
                        picture.crop_right = 0
                        #removing transparent image
                        os.remove(transparent_image_path)
                        #reverting back deformation
                        utils.MetaCommand('grstyle deform on')
                    else:
                        self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                        self.logger.info("")
                #image insertion for the shape named "Image 1"
                elif shape.name == "Image 1":
                    #capturing fringe bar of metapost window
                    utils.MetaCommand('add all')
                    utils.MetaCommand('add invert')
                    utils.MetaCommand('options fringebar on')
                    if self.threed_images_report_folder is not None:
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
                    else:
                        self.logger.info("ERROR : META 2D variable '{}' is not available or invalid. Please update.".format(GeneralVarInfo.report_directory_key))
                        self.logger.info("")
            endtime = datetime.now()
        except Exception as e:
            self.logger.exception("Error while seeding data into biw stiff ring deformation slide slide:\n{}".format(e))
            self.logger.info("")
            return 1
        self.logger.info("Completed seeding data into biw stiff ring deformation slide slide")
        self.logger.info("Time Taken : {}".format(endtime - starttime))
        self.logger.info("")

        return 0
