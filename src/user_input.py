#PYTHON SCRIPT
"""
   This script file is used for automating user inputs for side crash reporting either from in GUI or Non-GUI Mode.
"""

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
        """This method retrieves the input field values from a json file given bu user as input
        """
        json_path = UserInput.set_param_from_cmd_input("Please enter the input json path")
        json_loader = JsonLoader(json_path)
        input_json_dict = json_loader.load_json()
        self.metadb_2d_input = input_json_dict["2D_METADB_FILE_PATH"]
        self.metadb_3d_input = input_json_dict["3D_METADB_FILE_PATH"]
        self.d3hsp_file_path = input_json_dict["D3HSP_FILE_PATH"]
        self.target_metadb_input = input_json_dict["TARGET_METADB_FILE_PATH"]

        return 0

    @staticmethod
    def set_param_from_cmd_input(input_str, options_list=None):

        option_str = ""
        if options_list:
            for ind, option in options_list.items():
                if ind.startswith("empty"):
                    option_str += "\n"
                else:
                    option_str += "{} {} \n".format(ind, option)

        try:
            ret = input("{}{} : ".format(option_str, input_str)).strip()
        except EOFError:
            ret = ""

        ret_list = ret.split(",")
        new_ret = list()

        for ret in ret_list:
            # Check the string is not empty and having single or double quotes
            if len(ret) > 1 and (ret[0] == ret[-1]) and ret.startswith(("'", '"')):
                ret = ret[1:-1]

            # Append the ret to the new_ret
            new_ret.append(ret)

            # If input don't have any commas:
            if len(ret_list) == 1:
                return new_ret[0]

        return new_ret
