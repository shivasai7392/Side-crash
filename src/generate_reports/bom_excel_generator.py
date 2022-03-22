# PYTHON script
"""
##################################################
#      Copyright BETA CAE Systems USA Inc.,      #
#      2020 All Rights Reserved                  #
#      UNPUBLISHED, LICENSED SOFTWARE.           #
##################################################


Gui for running the side crash report generation.

Developer   : Naresh Medipalli and Shiva Sai
Date        : Dec 31, 2021

"""

import os
from openpyxl import Workbook
import logging

from meta import parts,constants, models

from src.meta_utilities import visualize_3d_critical_section

class ExcelBomGeneration():
    """
    This class is used to automate the excel BOM files of threed meta .

        Args:
            metadb_3d_input (Meta3DInfo): Meta3DInfo class object
            excel_bom_report_folder (str): path of the excel BOM files where we should save
        """
    def __init__(self, metadb_3d_input, excel_bom_report_folder):
        self.metadb_3d_input = metadb_3d_input
        self.excel_bom_report_folder = excel_bom_report_folder
        self.logger = logging.getLogger("side_crash_logger")

    def excel_bom_generation(self):
        """
        This method is used to generating the Excel Bill Of Material files

        Returns:
            0 : 0 for Success,1 for Failure
        """
        # Getting the Critical Sections Data From meta3d input
        critical_section_data = self.metadb_3d_input.critical_sections
        m = models.Model(0)
        # Iterating all the Critical Sections
        for key,value in critical_section_data.items():
            # If "hes" is there in value.keys and value with respective hes is not null
            if 'hes' in value.keys() and value['hes'] != 'null':
                # Generating the BOM for logging
                self.logger.info("GENERATING BOM : {}".format(value["name"] if "name" in value.keys() else "null"))
                # Loading the Workbook and making active then giving the headers for loaded Workbook
                workbook = Workbook()
                spreedsheet = workbook.active
                spreedsheet["A1"] = "PID"
                spreedsheet["B1"] = "Name"
                spreedsheet["C1"] = "Material"
                spreedsheet["D1"] = "Thickness"
                # Getting thr Parts Which are visible
                visualize_3d_critical_section(value)
                visible_parts = m.get_parts('visible')
                # applying length for visible parts
                self.logger.info("Number of parts identified : {}".format(len(visible_parts)))
                # Iterating all the visible parts
                for each_prop_entity in visible_parts:
                    # Getting the part type for each and every visible entity
                    part_type = parts.StringPartType(each_prop_entity.type)
                    # If the part type is PSHELL then getting the materials for part and getting name for material.
                    if part_type == "PSHELL":
                        part = parts.Part(id=each_prop_entity.id,type = constants.PSHELL, model_id=0)
                        materials = part.get_materials('all')
                        material_name = materials[0].name
                    # If the part type is PSHELL then getting the materials for part and getting name for material.
                    elif part_type == "PSOLID":
                        part = parts.Part(id=each_prop_entity.id,type = constants.PSOLID, model_id=0)
                        materials = part.get_materials('all')
                        material_name = materials[0].name
                    # appending the entity id,name,thickness and material name to the spreessheet
                    spreedsheet.append([each_prop_entity.id, each_prop_entity.name,material_name,round(each_prop_entity.shell_thick,1)])
                # Joining the excel path and saving
                excel_path = os.path.join(self.excel_bom_report_folder,"BOM_"+key+".xlsx").replace("\\","/")
                workbook.save(excel_path)
                self.logger.info("OUTPUT BOM : {}".format(excel_path))
                self.logger.info("CELLS WITH DATA : A1:D{}".format(str(len(visible_parts)+1)))
                self.logger.info("")

        return 0
