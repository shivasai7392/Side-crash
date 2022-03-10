# PYTHON script
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""

import os

from meta import utils, models
from src.meta_utilities import capture_image,visualize_3d_critical_section, annotation

class BIWStiffRingSpotWeldFailureSlide():
    def __init__(self,
                slide,
                windows,
                general_input,
                metadb_3d_input,
                threed_images_report_folder,
                template_file,
                ppt_report_folder) -> None:
        self.shapes = slide.shapes
        self.windows = windows
        self.general_input = general_input
        self.metadb_3d_input = metadb_3d_input
        self.threed_images_report_folder = threed_images_report_folder
        self.template_file = template_file
        self.ppt_report_folder = ppt_report_folder
    def setup(self,):
        """
        setup _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """


        return 0
    def edit(self, ):
        """
        edit _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """
        for shape in self.shapes:
            if shape.name == "Image 1":
                utils.MetaCommand('window maximize {}'.format(self.general_input.threed_window_name))
                utils.MetaCommand('0:options state original')
                utils.MetaCommand('options fringebar off')
                data = self.metadb_3d_input.critical_sections["f21_upb_outer"]
                visualize_3d_critical_section(data)
                m = models.Model(0)
                visible_parts = m.get_parts('visible')
                annotation(visible_parts)

                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"f21_upb_outer_stiff_ring_spotweld_failure".lower()+".png")
                capture_image("MetaPost",shape.width,shape.height,image_path,view = "left")
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                self.revert()
            elif shape.name == "Image 2":
                utils.MetaCommand('window maximize {}'.format(self.general_input.threed_window_name))
                utils.MetaCommand('0:options state original')
                utils.MetaCommand('options fringebar off')
                data = self.metadb_3d_input.critical_sections["f21_upb_inner"]
                visualize_3d_critical_section(data)
                m = models.Model(0)
                visible_parts = m.get_parts('visible')
                annotation(visible_parts)

                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"f21_upb_inner_stiff_ring_spotweld_failure".lower()+".png")
                capture_image("MetaPost",shape.width,shape.height,image_path,view = "right")
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                self.revert()


        return 0
    def revert(self, ):
        """
        revert _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """
        utils.MetaCommand('color pid transparency reset act')
        return 0