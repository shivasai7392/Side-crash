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

from src.meta_utilities import capture_image

class BIWCBUDeformationSlide():

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

        utils.MetaCommand('window maximize "MetaPost"')
        for shape in self.shapes:
            if shape.name == "Image 1":
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('grstyle scalarfringe disable')
                data = self.metadb_3d_input.critical_sections
                entities = list()
                for _key,value in data.items():
                    if 'hes' in value.keys():
                        prop_names = value['hes']
                        re_props = prop_names.split(",")
                        for re_prop in re_props:
                            entities.extend(self.metadb_3d_input.get_props(re_prop))
                self.metadb_3d_input.hide_all()
                self.metadb_3d_input.show_only_props(entities)
                utils.MetaCommand('view default isometric')
                utils.MetaCommand('options fringebar off')

                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"cbu_without_plastic_strain".lower()+".png")
                capture_image("MetaPost",shape.width,shape.height,image_path)
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('grstyle scalarfringe enable')

            elif shape.name == "Image 2":
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('grstyle scalarfringe enable')
                data = self.metadb_3d_input.critical_sections
                entities = list()
                for _key,value in data.items():
                    if 'hes' in value.keys():
                        prop_names = value['hes']
                        re_props = prop_names.split(",")
                        for re_prop in re_props:
                            entities.extend(self.metadb_3d_input.get_props(re_prop))
                self.metadb_3d_input.hide_all()
                self.metadb_3d_input.show_only_props(entities)
                utils.MetaCommand('view default isometric')
                utils.MetaCommand('options fringebar off')
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"cbu_with_plastic_strain".lower()+".png")
                capture_image("MetaPost",shape.width,shape.height,image_path)
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('0:options state variable "serial=0"')
            elif shape.name == "Image 3":
                utils.MetaCommand('add all')
                utils.MetaCommand('add invert')
                utils.MetaCommand('options fringebar on')
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"fringe_bar".lower()+".png")
                utils.MetaCommand('write scalarfringebar png {} '.format(image_path))
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('options fringebar off')

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
