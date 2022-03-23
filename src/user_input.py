#PYTHON SCRIPT
"""
   This script file is used for automating user inputs for side crash reporting either from in GUI or Non-GUI Mode.
"""
import os

from meta import utils

from src.json_loader import JsonLoader

class UserInput():
    """
    This Class is used to run the Script in the Windows and linux in gui mode and non-gui mode.
    """

    def __init__(self, *args):
        self.args = args
        self.metadb_2d_input = None
        self.metadb_3d_input = None
        self.d3hsp_file_path = None
        self.total_args_count_linux_nogui = 1
        self.target_metadb_input = None

    def get_user_input_from_gui(self):
        """
        get_user_input_from_gui [summary]

        [extended_summary]
        """
        # Getting the variables of 2d metadb file
        self.metadb_2d_input = utils.MetaGetVariable("2d_metadb_input")
        self.metadb_3d_input = utils.MetaGetVariable("3d_metadb_input")
        self.d3hsp_file_path = utils.MetaGetVariable("d3hsp_file_path")
        self.target_metadb_input = utils.MetaGetVariable("target_metadb_input")

        return 0

    def get_user_input_from_json(self):
        """This method retrieves the input field values from a json file
        """
        json_loader = JsonLoader(os.path.join(os.path.dirname(os.path.dirname(__file__))),"side_crash_reporter_input_json.json")
        input_json_dict = json_loader.load_json()
        self.metadb_2d_input = input_json_dict["2D_METADB_FILE_PATH"]
        self.metadb_3d_input = input_json_dict["3D_METADB_FILE_PATH"]
        self.d3hsp_file_path = input_json_dict["D3HSP_FILE_PATH"]
        self.target_metadb_input = input_json_dict["TARGET_METADB_FILE_PATH"]

        return 0
