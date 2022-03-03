# PYTHON script
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""

import os

from meta import utils
from meta import nodes
from meta import models

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
        from pptx.util import Pt

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
            elif shape.name == "Table 4":
                table_obj = shape.table
                table_row = table_obj.rows[1]
                text_frame = table_row.cells[1].text_frame
                font = text_frame.paragraphs[0].font
                font.name = 'Arial'
                font.size = Pt(11)
                font.bold = True
                font.underline = True
                text_frame.paragraphs[0].text = str(round(float(self.general_input.test_mass_value)*1000,2))

                table_row = table_obj.rows[2]
                text_frame = table_row.cells[1].text_frame
                font = text_frame.paragraphs[0].font
                font.name = 'Arial'
                font.size = Pt(11)
                text_frame.paragraphs[0].text = str(round(float(self.general_input.physical_mass_value)*1000,2))

                table_row = table_obj.rows[3]
                text_frame = table_row.cells[1].text_frame
                font = text_frame.paragraphs[0].font
                font.name = 'Arial'
                font.size = Pt(11)
                text_frame.paragraphs[0].text = str(round(float(self.general_input.added_mass_value)*1000,2))

                table_row = table_obj.rows[6]
                text_frame = table_row.cells[1].text_frame
                font = text_frame.paragraphs[0].font
                font.name = 'Arial'
                font.size = Pt(11)
                text_frame.paragraphs[0].text = str(round(float(self.general_input.total_mass_value)*1000,2))

                table_row = table_obj.rows[7]
                text_frame = table_row.cells[1].text_frame
                font = text_frame.paragraphs[0].font
                font.name = 'Arial'
                font.size = Pt(11)
                text_frame.paragraphs[0].text = str(round((float(self.general_input.test_mass_value)-float(self.general_input.total_mass_value))*1000,2))

            elif shape.name == "Table 1":
                MDB_fr_node_id = int(self.general_input.MDB_fr_node_id)
                MDB_fr_node = nodes.Node(id=MDB_fr_node_id, model_id=0)
                table_obj = shape.table
                table_row = table_obj.rows[3]
                text_frame = table_row.cells[1].text_frame
                font = text_frame.paragraphs[0].font
                font.name = 'Arial'
                font.size = Pt(11)
                text_frame.paragraphs[0].text = str(round(MDB_fr_node.x))

                target_row = table_obj.rows[2]
                text_frame = target_row.cells[1].text_frame
                target_z_value = int(text_frame.paragraphs[0].text.strip())

                table_row = table_obj.rows[4]
                text_frame = table_row.cells[1].text_frame
                font = text_frame.paragraphs[0].font
                font.name = 'Arial'
                font.size = Pt(11)
                value = round(MDB_fr_node.x) - target_z_value
                text_frame.paragraphs[0].text = str("+"+str(value) if value>0 else value)

            elif shape.name == "Table 2":
                model = models.Model(0)
                res = model.get_current_resultset()
                struck_subframe_node_ids = self.general_input.struck_subframe_node_ids
                struck_subframe_node1 = nodes.Node(id=int(struck_subframe_node_ids.split("/")[0]), model_id=0)
                struck_subframe_node2 = nodes.Node(id=int(struck_subframe_node_ids.split("/")[1]), model_id=0)
                MDB_fr_node_id = int(self.general_input.MDB_fr_node_id)
                MDB_fr_node = nodes.Node(id=MDB_fr_node_id, model_id=0)
                MDB_rr_node_id = int(self.general_input.MDB_rr_node_id)
                MDB_rr_node = nodes.Node(id=MDB_rr_node_id, model_id=0)
                distance_nfr_s1 = MDB_fr_node.get_distance_from_node(res, struck_subframe_node1, res)
                distance_nrr_s1 = MDB_fr_node.get_distance_from_node(res, struck_subframe_node1, res)
                distance_nfr_s1 = sum([dist**2 for dist in distance_nfr_s1])
                distance_nrr_s1 = sum([dist**2 for dist in distance_nrr_s1])
                if distance_nfr_s1>distance_nrr_s1:
                    suspension_rr_node = struck_subframe_node1
                    suspension_fr_node = struck_subframe_node2
                else:
                    suspension_rr_node = struck_subframe_node2
                    suspension_fr_node = struck_subframe_node1

                table_obj = shape.table
                table_row = table_obj.rows[2]
                text_frame = table_row.cells[2].text_frame
                font = text_frame.paragraphs[0].font
                font.name = 'Arial'
                font.size = Pt(11)
                text_frame.paragraphs[0].text = str(abs(round(MDB_fr_node.x) - round(suspension_fr_node.x)))

                table_obj = shape.table
                table_row = table_obj.rows[2]
                text_frame = table_row.cells[1].text_frame
                front_target_overlap = int(text_frame.paragraphs[0].text.strip())

                table_obj = shape.table
                table_row = table_obj.rows[2]
                text_frame = table_row.cells[3].text_frame
                font = text_frame.paragraphs[0].font
                font.name = 'Arial'
                font.size = Pt(11)
                value = abs(round(MDB_fr_node.x) - round(suspension_fr_node.x)) - front_target_overlap
                text_frame.paragraphs[0].text = str("+"+str(value) if value>0 else value)

                table_row = table_obj.rows[3]
                text_frame = table_row.cells[2].text_frame
                font = text_frame.paragraphs[0].font
                font.name = 'Arial'
                font.size = Pt(11)
                text_frame.paragraphs[0].text = str(abs(round(MDB_rr_node.x) - round(suspension_rr_node.x)))

                table_obj = shape.table
                table_row = table_obj.rows[3]
                text_frame = table_row.cells[1].text_frame
                rear_target_overlap = int(text_frame.paragraphs[0].text.strip())

                table_obj = shape.table
                table_row = table_obj.rows[3]
                text_frame = table_row.cells[3].text_frame
                font = text_frame.paragraphs[0].font
                font.name = 'Arial'
                font.size = Pt(11)
                value = abs(round(MDB_rr_node.x) - round(suspension_rr_node.x)) - rear_target_overlap
                text_frame.paragraphs[0].text = str("+"+str(value) if value>0 else value)

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
