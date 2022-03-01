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

class ExcelBomGeneration():
    def __init__(self, metadb_3d_input):
        """
        __init__ _summary_

        _extended_summary_

        Args:
            metadb_3d_input (_type_): _description_
        """
        self.metadb_3d_input = metadb_3d_input
    def excel_bom_generation(self, ):
        critical_section_data = self.metadb_3d_input.critical_sections
        print(critical_section_data)