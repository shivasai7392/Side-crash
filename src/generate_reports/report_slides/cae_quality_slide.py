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

from src.general_utilities import add_row
from src.meta_utilities import capture_resized_image

class CAEQualitySlide():

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
        self.slide = slide
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

    def edit(self):
        """
        edit _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """

        self.setup()

        from PIL import Image
        from pptx.util import Pt

        window_name = self.general_input.cae_quality_window_name
        utils.MetaCommand('window maximize "{}"'.format(window_name))

        for shape in self.shapes:

            if shape.name == "Image 2":
                plot_id = 0
                page_id=0
                plot = plot2d.Plot(plot_id, window_name, page_id)
                title = plot.get_title()
                plot.activate()
                utils.MetaCommand('xyplot rlayout "{}" 1'.format(window_name))
                utils.MetaCommand('xyplot curve select "{}" all'.format(window_name))
                utils.MetaCommand('xyplot curve visible and "{}" sel'.format(window_name))
                image_path = os.path.join(self.twod_images_report_folder,window_name+"_"+title.get_text().lower()+".jpeg")
                capture_resized_image(window_name,shape.width,shape.height,image_path,plot_id=plot.id)
                picture = self.slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('xyplot plotoptions legend on "CAE quality" 0')
                utils.MetaCommand('xyplot legend hook left "CAE quality" 0')
                utils.MetaCommand('xyplot legend hook hout "CAE quality" 0')
                utils.MetaCommand('xyplot legend ymove "CAE quality" 0 1.060000')
                image2_path = os.path.join(self.twod_images_report_folder,window_name+"_"+title.get_text().lower()+"_Legend"+".jpeg")
                utils.MetaCommand('write jpeg "{}" 100'.format(image2_path))
                img_2 = Image.open(image2_path)
                img_2.save(image2_path, 'PNG')
                img_2 = Image.open(image2_path)
                legend = plot2d.Legend(plot_id, window_name, page_id)
                left,top = legend.get_position()
                width = legend.get_width()
                height = legend.get_height()
                img_2 = img_2.crop((left,top,width+8,height+8))
                img_2.save(image2_path,"PNG")
                shape2 = [shape for shape in self.slide.shapes if shape.name == "Image 1"][0]
                picture = self.slide.shapes.add_picture(image2_path,shape2.left,shape2.top,width = shape2.width,height = shape2.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('xyplot rlayout "{}" 2'.format(window_name))
                plot.deactivate()
            elif shape.name == "Image 3":
                plot_id = 1
                page_id=0
                plot = plot2d.Plot(plot_id, window_name, page_id)
                title = plot.get_title()
                plot.activate()
                utils.MetaCommand('xyplot rlayout "{}" 1'.format(window_name))
                utils.MetaCommand('xyplot curve select "{}" all'.format(window_name))
                utils.MetaCommand('xyplot curve visible and "{}" sel'.format(window_name))
                image_path = os.path.join(self.twod_images_report_folder,window_name+"_"+title.get_text().lower()+".jpeg")
                capture_resized_image(window_name,shape.width,shape.height,image_path,plot_id=plot.id)
                picture = self.slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                picture = self.slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                plot.deactivate()
                utils.MetaCommand('xyplot rlayout "{}" 2'.format(window_name))
            elif shape.name == "Table 1":
                plot_id = 0
                page_id=0
                plot = plot2d.Plot(plot_id, window_name, page_id)
                curvelist = plot.get_curves('all')
                index = 0
                for curve in curvelist:
                    if curve.id == 5:
                        continue
                    index = index+1
                    min_y = curve.get_limit_value_y(specifier = 'min')
                    max_y = curve.get_limit_value_y(specifier = 'max')
                    table_obj = shape.table
                    add_row(table_obj)
                    row = table_obj.rows[index]
                    text_frame_1 = row.cells[0].text_frame
                    font_1 = text_frame_1.paragraphs[0].font
                    font_1.size = Pt(12)
                    text_frame_1.paragraphs[0].text = str(curve.name).replace(" energy","")

                    text_frame_2 = row.cells[1].text_frame
                    font_2 = text_frame_2.paragraphs[0].font
                    font_2.size = Pt(12)
                    text_frame_2.paragraphs[0].text = "{:.2e}".format(max_y)

                    text_frame_3 = row.cells[2].text_frame
                    font_3 = text_frame_3.paragraphs[0].font
                    font_3.size = Pt(12)
                    text_frame_3.paragraphs[0].text = "{:.2e}".format(min_y)
            elif shape.name == "Table 2":
                table_obj = shape.table
                table_value_dict ={"Termination type":self.general_input.termination_type,
                                    "Computation time":self.general_input.computation_time,
                                    "Core count":self.general_input.core_count,
                                    "Verification mode":self.general_input.verification_mode,
                                    "Compute cluster":self.general_input.compute_cluster}
                index = 0
                for item,value in table_value_dict.items():
                    index = index+1
                    add_row(table_obj)
                    row = table_obj.rows[index]
                    text_frame_1 = row.cells[0].text_frame
                    font_1 = text_frame_1.paragraphs[0].font
                    font_1.size = Pt(12)
                    text_frame_1.paragraphs[0].text = item

                    if item == "Core count":
                        value = value.split("with")[1].rstrip()
                    text_frame_2 = row.cells[1].text_frame
                    font_2 = text_frame_2.paragraphs[0].font
                    font_2.size = Pt(12)
                    text_frame_2.paragraphs[0].text = value

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
