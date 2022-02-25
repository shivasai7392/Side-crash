# PYTHON script
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""

import os

from meta import utils

from src.meta_utilities import capture_image

class CBUAndBarrierPositionSlide():

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

    def edit(self):
        """
        edit _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """

        self.setup()

        from PIL import Image
        utils.MetaCommand('window maximize "MetaPost"')
        utils.MetaCommand('0:options state original')
        utils.MetaCommand('options fringebar off')
        for shape in self.shapes:
            if shape.name == "Image 4":
                self.metadb_3d_input.show_all()
                self.metadb_3d_input.hide_floor()
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"cbu_and_barrier".lower()+".png")
                capture_image("MetaPost",shape.width,shape.height,image_path,rotate = Image.ROTATE_90,view = "top")
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
            elif shape.name == "Image 3":
                self.metadb_3d_input.show_all()
                self.metadb_3d_input.hide_floor()
                utils.MetaCommand('color pid Gray act')
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"cbu_critical".lower()+".png")
                capture_image("MetaPost",shape.width,shape.height,image_path,view = "left")
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('color pid reset act')
        utils.MetaCommand('options fringebar on')

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
