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

from meta import parts,constants

class ExcelBomGeneration():
    def __init__(self, metadb_3d_input, excel_bom_report_folder):
        """
        __init__ _summary_

        _extended_summary_

        Args:
            metadb_3d_input (_type_): _description_
        """
        self.metadb_3d_input = metadb_3d_input
        self.excel_bom_report_folder = excel_bom_report_folder
    def excel_bom_generation(self, ):
        critical_section_data = self.metadb_3d_input.critical_sections

        for key,value in critical_section_data.items():
            if 'hes' in value.keys() and value['hes'] != 'null':
                workbook = Workbook()
                spreedsheet = workbook.active
                spreedsheet["A1"] = "PID"
                spreedsheet["B1"] = "Name"
                spreedsheet["C1"] = "Material"
                spreedsheet["D1"] = "Thickness"
                prop_names = value['hes']
                re_props = prop_names.split(",")
                for re_prop in re_props:
                    prop_entities = self.metadb_3d_input.get_props(re_prop)

                    for each_prop_entity in prop_entities:
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
