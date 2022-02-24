# PYTHON script
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""

import os

from meta import utils
from meta import plot2d
from meta import windows

from src.meta_utilities import capture_image
from src.meta_utilities import capture_resized_image

class BIWKinematicsSlide():

    def __init__(self,
                slide,
                windows,
                general_input,
                metadb_2d_input,
                metadb_3d_input,
                template_file,
                twod_images_report_folder,
                threed_images_report_folder,
                ppt_report_folder) -> None:
        self.shapes = slide.shapes
        self.windows = windows
        self.general_input = general_input
        self.metadb_2d_input = metadb_2d_input
        self.metadb_3d_input = metadb_3d_input
        self.template_file = template_file
        self.twod_images_report_folder = twod_images_report_folder
        self.threed_images_report_folder = threed_images_report_folder
        self.ppt_report_folder = ppt_report_folder

    def setup(self):
        """
        setup _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """

        return 0

    def kinematics_curve_format(self, biw_accel_window_name, plot_id,title = None):
        """
        kinematics_curve_format _summary_

        _extended_summary_

        Args:
            biw_accel_window_name (_type_): _description_
            plot_id (_type_): _description_
            title (_type_, optional): _description_. Defaults to None.

        Returns:
            _type_: _description_
        """

        utils.MetaCommand('xyplot axisoptions yaxis active "BIW - Accel" {} 0'.format(plot_id))
        utils.MetaCommand('xyplot axisoptions yaxis hideaxis "BIW - Accel" {} 0'.format(plot_id))
        utils.MetaCommand('xyplot gridoptions line major style "BIW - Accel" {} 0'.format(plot_id))
        if title:
            utils.MetaCommand('xyplot plotoptions title set "{}" {} "{}"'.format(biw_accel_window_name, plot_id, title))
        utils.MetaCommand('xyplot axisoptions xaxis active "{}" {} 0'.format(biw_accel_window_name, plot_id))
        utils.MetaCommand('xyplot axisoptions xlabel font "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
        utils.MetaCommand('xyplot axisoptions labels xfont "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
        utils.MetaCommand('xyplot axisoptions xaxis deactive "{}" {} 0'.format(biw_accel_window_name, plot_id))

        utils.MetaCommand('xyplot axisoptions yaxis active "{}" {} 1'.format(biw_accel_window_name, plot_id))
        utils.MetaCommand('xyplot axisoptions ylabel font "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
        utils.MetaCommand('xyplot axisoptions labels yfont "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
        utils.MetaCommand('xyplot axisoptions yaxis deactive "{}" {} 1'.format(biw_accel_window_name, plot_id))

        utils.MetaCommand('xyplot axisoptions yaxis active "{}" {} 2'.format(biw_accel_window_name, plot_id))
        utils.MetaCommand('xyplot axisoptions ylabel font "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
        utils.MetaCommand('xyplot axisoptions labels yfont "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
        utils.MetaCommand('xyplot axisoptions yaxis deactive "{}" {} 2'.format(biw_accel_window_name, plot_id))

        utils.MetaCommand('xyplot plotoptions title font "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))

        return 0

    def edit(self):
        """
        edit _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """

        self.setup()

        biw_accel_window_name = self.general_input.biw_accel_window_name
        biw_accel_window_name = biw_accel_window_name.replace("\"","")
        utils.MetaCommand('window maximize "{}"'.format(biw_accel_window_name))
        window_biw_accel = windows.Window(biw_accel_window_name, page_id = 0)
        layout = window_biw_accel.get_plot_layout()

        for shape in self.shapes:
            if shape.name == "Image 1":
                utils.MetaCommand('window maximize "MetaPost"')
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('options fringebar off')
                self.metadb_3d_input.show_all()
                self.metadb_3d_input.hide_floor()
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"model_front".lower()+".png")
                capture_image("MetaPost",shape.width,shape.height,image_path,view = "front")
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('0:options state original')
                utils.MetaCommand('options fringebar off')
            elif shape.name == "Image 2":
                utils.MetaCommand('window maximize "MetaPost"')
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('options fringebar off')
                self.metadb_3d_input.show_all()
                self.metadb_3d_input.hide_floor()
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"model_top".lower()+".png")
                capture_image("MetaPost",shape.width,shape.height,image_path,view = "top")
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('0:options state original')
                utils.MetaCommand('options fringebar off')
            elif shape.name == "Image 3":
                utils.MetaCommand('window maximize "MetaPost"')
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('options fringebar off')
                self.metadb_3d_input.show_all()
                self.metadb_3d_input.hide_floor()
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"model_front".lower()+".png")
                capture_image("MetaPost",shape.width,shape.height,image_path,view = "front")
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('0:options state original')
                utils.MetaCommand('options fringebar off')
            elif shape.name == "Image 4":
                utils.MetaCommand('window maximize "MetaPost"')
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('options fringebar off')
                self.metadb_3d_input.show_all()
                self.metadb_3d_input.hide_floor()
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"model_top".lower()+".png")
                capture_image("MetaPost",shape.width,shape.height,image_path,view = "top")
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('0:options state original')
                utils.MetaCommand('options fringebar off')
            elif shape.name == "Image 5":
                plot_id = 3
                page_id = 0
                plot = plot2d.Plot(plot_id, biw_accel_window_name, page_id)
                biw_accel_curves = plot.get_curves('all')
                for each_biw_accel_curve in biw_accel_curves:
                    if str(each_biw_accel_curve.name).endswith(("X velocity", "X displacement")):
                        each_biw_accel_curve.show()
                title = plot2d.Title(plot_id, biw_accel_window_name, page_id)
                plot.activate()
                utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, 1))
                self.kinematics_curve_format(biw_accel_window_name, plot_id)

                image_path = os.path.join(self.twod_images_report_folder,biw_accel_window_name+"_"+title.get_text().lower()+".png")
                capture_resized_image(biw_accel_window_name,shape.width,shape.height,image_path,plot_id=0)
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, layout))
            elif shape.name == "Image 6":
                plot_id = 0
                page_id = 0
                plot = plot2d.Plot(plot_id, biw_accel_window_name, page_id)
                biw_accel_curves = plot.get_curves('all')
                for each_biw_accel_curve in biw_accel_curves:
                    if (str(each_biw_accel_curve.name).__contains__("UNIT")) and (str(each_biw_accel_curve.name).endswith(("Y velocity", "Y displacement"))):
                        each_biw_accel_curve.show()
                title = plot2d.Title(plot_id, biw_accel_window_name, page_id)
                plot.activate()
                utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, plot.id))
                self.kinematics_curve_format(biw_accel_window_name, plot_id, title = "UNIT")

                image_path = os.path.join(self.twod_images_report_folder,biw_accel_window_name+"_"+title.get_text().lower()+".png")
                capture_resized_image(biw_accel_window_name,shape.width,shape.height,image_path,plot_id=0)
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, layout))
            elif shape.name == "Image 7":
                plot_id = 1
                page_id = 0
                plot = plot2d.Plot(plot_id, biw_accel_window_name, page_id)
                biw_accel_curves = plot.get_curves('all')
                for each_biw_accel_curve in biw_accel_curves:
                    if (str(each_biw_accel_curve.name).__contains__("APLR_R")) and (str(each_biw_accel_curve.name).endswith(("Y velocity", "Y displacement"))):
                        each_biw_accel_curve.show()
                title = plot2d.Title(plot_id, biw_accel_window_name, page_id)
                plot.activate()
                utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, plot.id))
                self.kinematics_curve_format(biw_accel_window_name, plot_id, title="APLR_R")

                image_path = os.path.join(self.twod_images_report_folder,biw_accel_window_name+"_"+title.get_text().lower()+".png")
                capture_resized_image(biw_accel_window_name,shape.width,shape.height,image_path,plot_id=0)
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                plot.deactivate()
                utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, layout))

            elif shape.name == "Image 8":
                plot_id = 1
                page_id = 0
                plot = plot2d.Plot(plot_id, biw_accel_window_name, page_id)
                biw_accel_curves = plot.get_curves('all')
                biw_accel_window_curves_all = window_biw_accel.get_curves('all')
                window_biw_accel.hide_curves(biw_accel_window_curves_all)
                for each_biw_accel_curve in biw_accel_curves:
                    if (str(each_biw_accel_curve.name).__contains__("SIS_ROW2_R")) and (str(each_biw_accel_curve.name).endswith(("Y velocity", "Y displacement"))):

                        each_biw_accel_curve.show()
                title = plot2d.Title(plot_id, biw_accel_window_name, page_id)
                plot.activate()
                utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, plot.id))
                self.kinematics_curve_format(biw_accel_window_name, plot_id,title="SS_BP_R")

                image_path = os.path.join(self.twod_images_report_folder,biw_accel_window_name+"_"+title.get_text().lower()+".png")
                capture_resized_image(biw_accel_window_name,shape.width,shape.height,image_path,plot_id=0)
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                plot.deactivate()
                utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, layout))
            elif shape.name == "Image 9":
                plot_id = 1
                page_id = 0
                plot = plot2d.Plot(plot_id, biw_accel_window_name, page_id)
                biw_accel_curves = plot.get_curves('all')
                biw_accel_window_curves_all = window_biw_accel.get_curves('all')
                window_biw_accel.hide_curves(biw_accel_window_curves_all)
                for each_biw_accel_curve in biw_accel_curves:
                    if (str(each_biw_accel_curve.name).__contains__("SS_R_RR_TOP_G_Y1")) and (str(each_biw_accel_curve.name).endswith(("Y velocity", "Y displacement"))):

                        each_biw_accel_curve.show()
                title = plot2d.Title(plot_id, biw_accel_window_name, page_id)
                plot.activate()
                utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, plot.id))
                self.kinematics_curve_format(biw_accel_window_name, plot_id, title="SS_R_RR_TOP")

                image_path = os.path.join(self.twod_images_report_folder,biw_accel_window_name+"_"+title.get_text().lower()+".png")
                capture_resized_image(biw_accel_window_name,shape.width,shape.height,image_path,plot_id=0)
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                plot.deactivate()
                utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, layout))
            elif shape.name == "Image 10":
                plot_id = 2
                page_id = 0
                plot = plot2d.Plot(plot_id, biw_accel_window_name, page_id)
                biw_accel_curves = plot.get_curves('all')
                for each_biw_accel_curve in biw_accel_curves:
                    if (str(each_biw_accel_curve.name).__contains__("UPR")) and (str(each_biw_accel_curve.name).endswith(("Y velocity", "Y displacement"))):

                        each_biw_accel_curve.show()
                title = plot2d.Title(plot_id, biw_accel_window_name, page_id)
                plot.activate()
                utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, plot.id))
                self.kinematics_curve_format(biw_accel_window_name, plot_id, title="BPLR_MID_R")

                image_path = os.path.join(self.twod_images_report_folder,biw_accel_window_name+"_"+title.get_text().lower()+".png")
                capture_resized_image(biw_accel_window_name,shape.width,shape.height,image_path,plot_id=0)
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                plot.deactivate()
                utils.MetaCommand('xyplot rlayout "{}" {}'.format(biw_accel_window_name, layout))

        self.revert()

        return 0

    def revert(self):
        """
        revert _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """

        return 0
