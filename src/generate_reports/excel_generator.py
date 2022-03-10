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

from meta import parts,constants, models

from src.meta_utilities import visualize_3d_critical_section

class ExcelBomGeneration():
    def __init__(self, metadb_3d_input, excel_bom_report_folder,logger):
        """
        __init__ _summary_

        _extended_summary_

        Args:
            metadb_3d_input (_type_): _description_
        """
        self.metadb_3d_input = metadb_3d_input
        self.excel_bom_report_folder = excel_bom_report_folder
        self.logger = logger

    def excel_bom_generation(self, ):
        critical_section_data = self.metadb_3d_input.critical_sections
        for key,value in critical_section_data.items():
            if 'hes' in value.keys() and value['hes'] != 'null':
                self.logger.log.info("GENERATING BOM : {}".format(value["name"] if "name" in value.keys() else "null"))
                workbook = Workbook()
                spreedsheet = workbook.active
                spreedsheet["A1"] = "PID"
                spreedsheet["B1"] = "Name"
                spreedsheet["C1"] = "Material"
                spreedsheet["D1"] = "Thickness"

                visualize_3d_critical_section(value)
                m = models.Model(0)
                visible_parts = m.get_parts('visible')

                self.logger.log.info("Number of parts identified : {}".format(len(visible_parts)))
                for each_prop_entity in visible_parts:
                    part_type = parts.StringPartType(each_prop_entity.type)
                    if part_type == "PSHELL":
                        part = parts.Part(id=each_prop_entity.id,type = constants.PSHELL, model_id=0)
                        materials = part.get_materials('all')
                        material_name = materials[0].name
                    elif part_type == "PSOLID":
                        part = parts.Part(id=each_prop_entity.id,type = constants.PSOLID, model_id=0)
                        materials = part.get_materials('all')
                        material_name = materials[0].name
                    spreedsheet.append([each_prop_entity.id, each_prop_entity.name,material_name,round(each_prop_entity.shell_thick,1)])

                excel_path = os.path.join(self.excel_bom_report_folder,"BOM_"+key.lower()+".xlsx").replace("\\","/")
                workbook.save(excel_path)
                self.logger.log.info("OUTPUT BOM : {}".format(excel_path))
                self.logger.log.info("CELLS WITH DATA : A1:D{}".format(20))
                self.logger.log.info("")
                self.logger.log.info("")
