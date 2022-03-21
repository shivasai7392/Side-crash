# PYTHON script
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""
import os
from meta import utils,windows,plot2d,models

from src.meta_utilities import capture_resized_image,visualize_3d_critical_section, visualize_annotation
from src.general_utilities import closest

class BIWROOFDeformationAndSpotWeldFailure():
    def __init__(self,
                slide,
                windows,
                general_input,
                metadb_3d_input,
                metadb_2d_input,
                template_file,
                twod_images_report_folder,
                threed_images_report_folder,
                ppt_report_folder) -> None:
        self.shapes = slide.shapes
        self.windows = windows
        self.general_input = general_input
        self.metadb_3d_input = metadb_3d_input
        self.metadb_2d_input = metadb_2d_input
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
    # def annotation(self):
    #     """
    #     annotation

    #     _extended_summary_
    #     """
    #     utils.MetaCommand('window maximize {}'.format(self.general_input.threed_window_name))
    #     utils.MetaCommand('0:options state original')
    #     utils.MetaCommand('options fringebar off')
    #     data = self.metadb_3d_input.critical_sections["f21_roof"]
    #     visualize_3d_critical_section(data)
    #     m = models.Model(0)
    #     self.visible_parts = m.get_parts('visible')
    #     utils.MetaCommand('add element connected')

    #     return 0
    def edit(self, ):
        from PIL import Image
        self.setup()
        for shape in self.shapes:
            if shape.name == "Image 1":
                utils.MetaCommand('add all')
                utils.MetaCommand('add invert')
                utils.MetaCommand('options fringebar on')
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"Fringe_bar".lower()+".png")
                utils.MetaCommand('write scalarfringebar png {} '.format(image_path))
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('options fringebar off')
            elif shape.name == "Image 2":
                data = self.metadb_3d_input.critical_sections["f21_roof"]
                visualize_3d_critical_section(data)
                utils.MetaCommand('color pid transparency reset act')
                utils.MetaCommand('window maximize "MetaPost"')
                utils.MetaCommand('options fringebar off')
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('grstyle deform on')
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"f21_roof_with_deform".lower()+".png")
                capture_resized_image("MetaPost",shape.width,shape.height,image_path,rotate = Image.ROTATE_270,view = "btm")
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
            elif shape.name == "Image 3":
                data = self.metadb_3d_input.critical_sections["f21_roof"]
                visualize_3d_critical_section(data)
                utils.MetaCommand('color pid transparency reset act')
                utils.MetaCommand('window maximize "MetaPost"')
                utils.MetaCommand('options fringebar off')
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('grstyle deform off')
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"f21_roof_without_deform".lower()+".png")
                capture_resized_image("MetaPost",shape.width,shape.height,image_path,rotate = Image.ROTATE_270,view = "btm")
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
            elif shape.name == "Image 4":
                window_name = self.general_input.biw_stiff_ring_deformation_name
                win = windows.Window(str(window_name), page_id=0)
                layout = win.get_plot_layout()
                utils.MetaCommand('window maximize "{}"'.format(window_name))
                final_time_curve_name = "ROOF_LINE_{}MS".format(round(float(self.general_input.survival_space_final_time)))
                initial_time_curve_name = "ROOF_LINE_0MS"
                plot_id = 0
                page_id=0
                plot = plot2d.Plot(plot_id, window_name, page_id)
                title = plot.get_title()
                plot.activate()
                final_curve = plot2d.CurvesByName(window_name, final_time_curve_name, 0)[0]
                final_curve.show()
                initial_curve = plot2d.CurvesByName(window_name, initial_time_curve_name, 0)[0]
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
                peak_curve_value = plot.get_curves('byname', name ="ROOF_LINE_"+str(peak_curve_value)+"MS")
                peak_curve = plot2d.CurvesByName(window_name, peak_curve_value[0].name, 0)[0]
                peak_curve.show()
                utils.MetaCommand('xyplot plotactive "{}" 0'.format(window_name))
                utils.MetaCommand('xyplot rlayout "{}" 1'.format(window_name))
                utils.MetaCommand('xyplot curve visible and "{}" {} {},{}'.format(window_name,initial_curve.id,peak_curve.id, final_curve.id))
                utils.MetaCommand('xyplot curve set style "{}" {} 9'.format(window_name, initial_curve.id))
                utils.MetaCommand('xyplot curve set style "{}" {} 5'.format(window_name,peak_curve.id))
                utils.MetaCommand('xyplot curve select "{}" all'.format(window_name))
                utils.MetaCommand('xyplot axisoptions yaxis active "{}" 0 0'.format(window_name))
                utils.MetaCommand('xyplot axisoptions ylabel font "{}" 0 "Arial,25,-1,5,75,0,0,0,0,0"'.format(window_name))
                utils.MetaCommand('xyplot axisoptions labels yfont "{}" 0 "Arial,25,-1,5,75,0,0,0,0,0"'.format(window_name))
                utils.MetaCommand('xyplot axisoptions yaxis deactive "{}" 0 0'.format(window_name))
                utils.MetaCommand('xyplot axisoptions xaxis active "{}" 0 0'.format(window_name))
                utils.MetaCommand('xyplot axisoptions xlabel font "{}" 0 "Arial,25,-1,5,75,0,0,0,0,0"'.format(window_name))
                utils.MetaCommand('xyplot axisoptions labels xfont "{}" 0 "Arial,25,-1,5,75,0,0,0,0,0"'.format(window_name))
                utils.MetaCommand('xyplot plotoptions title font "{}" 0 "Arial,25,-1,5,75,0,0,0,0,0"'.format(window_name))
                image_path = os.path.join(self.twod_images_report_folder,window_name+"_"+title.get_text().lower()+".png").replace("&","and")
                capture_resized_image(window_name,shape.width,shape.height,image_path,plot_id=plot.id)
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                plot.deactivate()
                utils.MetaCommand('xyplot rlayout "{}" {}'.format(window_name, layout))
            elif shape.name == "Image 5":
                utils.MetaCommand('window maximize {}'.format(self.general_input.threed_window_name))
                utils.MetaCommand('0:options state original')
                utils.MetaCommand('options fringebar off')
                data = self.metadb_3d_input.critical_sections["f21_roof"]
                visualize_3d_critical_section(data)
                m = models.Model(0)
                visible_parts = m.get_parts('visible')
                visualize_annotation(visible_parts,self.metadb_3d_input.spotweld_clusters,self.general_input.binout_directory)

                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"f21_roof_spotweld_failure".lower()+".png")
                capture_resized_image("MetaPost",shape.width,shape.height,image_path,view = "top")
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                self.revert()

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