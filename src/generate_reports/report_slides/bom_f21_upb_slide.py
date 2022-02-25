# PYTHON script
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""

import os

from meta import utils
from meta import parts
from meta import constants

from src.meta_utilities import capture_image,visualize_3d_critical_section
from src.general_utilities import add_row

class BOMF21UPBSlide():

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

        from pptx.util import Pt

        utils.MetaCommand('0:options state original')
        for shape in self.shapes:
            if shape.name == "Image 2":
                data = self.metadb_3d_input.critical_sections["f21_upb_inner"]
                visualize_3d_critical_section(data)
                utils.MetaCommand('window maximize "MetaPost"')
                utils.MetaCommand('options fringebar off')
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"f21_upb_inner".lower()+".png")
                capture_image("MetaPost",shape.width,shape.height,image_path,view = "right")
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
            if shape.name == "Image 1":
                data = self.metadb_3d_input.critical_sections["f21_upb_outer"]
                visualize_3d_critical_section(data)
                utils.MetaCommand('window maximize "MetaPost"')
                utils.MetaCommand('options fringebar off')
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"f21_upb_inner".lower()+".png")
                capture_image("MetaPost",shape.width,shape.height,image_path,view = "left")
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
            elif shape.name == "Table 1":
                data = self.metadb_3d_input.critical_sections["f21_upb_outer"]
                prop_names = data["hes"]
                re_props = prop_names.split(",")
                entities = []
                for re_prop in re_props:
                    utils.MetaCommand('window maximize "MetaPost"')
                    entities.extend(self.metadb_3d_input.get_props(re_prop))
                table_obj = shape.table
                for id,prop in enumerate(entities[:15]):
                    part = parts.Part(id=prop.id, type=constants.PSHELL, model_id=0)
                    materials = part.get_materials('all')

                    add_row(table_obj)
                    prop_row = table_obj.rows[id+1]

                    text_frame = prop_row.cells[0].text_frame
                    font = text_frame.paragraphs[0].font
                    font.size = Pt(8)
                    text_frame.paragraphs[0].text = str(prop.id)

                    text_frame_name = prop_row.cells[1].text_frame
                    font_name = text_frame_name.paragraphs[0].font
                    font_name.size = Pt(8)
                    text_frame_name.paragraphs[0].text = str(prop.name)

                    text_frame_material = prop_row.cells[2].text_frame
                    font_material = text_frame_material.paragraphs[0].font
                    font_material.size = Pt(8)
                    for each_material in materials:
                        text_frame_material.paragraphs[0].text = str(each_material.name)

                    text_frame_thickness = prop_row.cells[3].text_frame
                    font_thickness = text_frame_thickness.paragraphs[0].font
                    font_thickness.size = Pt(8)
                    thickness = round(prop.shell_thick,1)
                    text_frame_thickness.paragraphs[0].text = str(thickness)


            elif shape.name == "Table 2":
                data = self.metadb_3d_input.critical_sections["f21_upb_outer"]
                prop_names = data["hes"]
                re_props = prop_names.split(",")
                entities = []
                for re_prop in re_props:
                    utils.MetaCommand('window maximize "MetaPost"')
                    entities.extend(self.metadb_3d_input.get_props(re_prop))

                entities_all = []
                for each_entity in entities[15:]:
                    if str(each_entity.id).startswith("2"):
                        entities_all.append(each_entity)

                table_obj = shape.table
                for id,prop in enumerate(entities_all):

                    part = parts.Part(id=prop.id, type=constants.PSHELL, model_id=0)
                    materials = part.get_materials('all')

                    add_row(table_obj)
                    prop_row = table_obj.rows[id+1]
                    text_frame = prop_row.cells[0].text_frame
                    font = text_frame.paragraphs[0].font
                    font.size = Pt(8)
                    text_frame.paragraphs[0].text = str(prop.id)

                    text_frame_name = prop_row.cells[1].text_frame
                    font_name = text_frame_name.paragraphs[0].font
                    font_name.size = Pt(8)
                    text_frame_name.paragraphs[0].text = str(prop.name)

                    text_frame_material = prop_row.cells[2].text_frame
                    font_material = text_frame_material.paragraphs[0].font
                    font_material.size = Pt(8)
                    for each_material in materials:
                        text_frame_material.paragraphs[0].text = str(each_material.name)

                    text_frame_thickness = prop_row.cells[3].text_frame
                    font_thickness = text_frame_thickness.paragraphs[0].font
                    font_thickness.size = Pt(8)
                    thickness = round(prop.shell_thick,1)
                    text_frame_thickness.paragraphs[0].text = str(thickness)
            elif shape.name == "Table 3":
                data = self.metadb_3d_input.critical_sections["f21_upb_inner"]
                prop_names = data["hes"]
                re_props = prop_names.split(",")
                entities = []
                for re_prop in re_props:
                    utils.MetaCommand('window maximize "MetaPost"')
                    entities.extend(self.metadb_3d_input.get_props(re_prop))

                entities_all = []
                for each_entity in entities:
                    if str(each_entity.id).startswith("2"):
                        entities_all.append(each_entity)

                table_obj = shape.table
                for id,prop in enumerate(entities_all):

                    part = parts.Part(id=prop.id, type=constants.PSHELL, model_id=0)
                    materials = part.get_materials('all')

                    add_row(table_obj)
                    prop_row = table_obj.rows[id+1]

                    text_frame = prop_row.cells[0].text_frame
                    font = text_frame.paragraphs[0].font
                    font.size = Pt(8)
                    text_frame.paragraphs[0].text = str(prop.id)

                    text_frame_name = prop_row.cells[1].text_frame
                    font_name = text_frame_name.paragraphs[0].font
                    font_name.size = Pt(8)
                    text_frame_name.paragraphs[0].text = str(prop.name)

                    text_frame_material = prop_row.cells[2].text_frame
                    font_material = text_frame_material.paragraphs[0].font
                    font_material.size = Pt(8)
                    for each_material in materials:
                        text_frame_material.paragraphs[0].text = str(each_material.name)

                    text_frame_thickness = prop_row.cells[3].text_frame
                    font_thickness = text_frame_thickness.paragraphs[0].font
                    font_thickness.size = Pt(8)
                    thickness = round(prop.shell_thick,1)
                    text_frame_thickness.paragraphs[0].text = str(thickness)

            elif shape.name == "TextBox 1":
                text_frame = shape.text_frame
                text_frame.clear()
                p = text_frame.paragraphs[0]
                run = p.add_run()
                run.text = "OUTER"
            elif shape.name == "TextBox 2":
                text_frame = shape.text_frame
                text_frame.clear()
                p = text_frame.paragraphs[0]
                run = p.add_run()
                run.text = "INNER"

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
