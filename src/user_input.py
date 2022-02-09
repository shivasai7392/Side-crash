"""
    [summary]

[extended_summary]
"""
import sys

from meta import utils
from meta import session



class UserInput():
    """
    UserInput [summary]

    [extended_summary]
    """

    def __init__(self, *args):
        self.args = args
        self.metadb_2d_input = None
        self.total_args_count_linux_nogui = 1

    def get_user_input_from_gui(self):
        """
        get_user_input_from_gui [summary]

        [extended_summary]
        """
        self.metadb_2d_input = utils.MetaGetVariable("2d_metadb_input")

        return 0

    def get_user_input_from_interactive_mode(self):
        """This method checks whether the script initiated in metapost or shell/bash command line
        """

        if "win" in sys.platform:
            ret = UserInput.continue_in_windows_cmd(self.args)
            if ret == 0: # Script is running in GUI mode in metapost, so no attributes setting is needed
                print("Script is running in GUI mode")
                self.get_user_input_from_gui()
            elif ret == -1:  # Multiple line arguments
                self.run_interactive_mode()

        elif "linux" in sys.platform:
            ret = UserInput.continue_in_linux_cmd(self.total_args_count_linux_nogui)
            if ret == 0:  # Script is running in GUI mode in metapost, so no attributes setting is needed
                print("Script is running in GUI mode")
                self.get_user_input_from_gui()
            elif ret == -1: # Multiple line arguments
                self.run_interactive_mode()

        return 0

    @staticmethod
    def continue_in_windows_cmd(actual_args_list):
        """This method checks the given argument in command line is sufficient to run the script

        Args:
            total_args (list): list of command line arguments
        Returns:
            int: 1 if
        """

        prog_args = session.ProgramArguments()

        if actual_args_list:
            if len(prog_args) > 1  and prog_args[1] == "-b":
                if any(item == "" for item in actual_args_list[:3]):
                    return -1
                return 1

        return 0

    @staticmethod
    def continue_in_linux_cmd(total_args_count_linux_nogui):
        """This method gets and sets variables from the command line arguments

        Returns:
            [int]: 0 as always.
        """

        prog_args = session.ProgramArguments()

        # If the script is executed from the command line
        if len(prog_args) > 1  and prog_args[2] == "-b":

            args = list()

            for ind in range(total_args_count_linux_nogui):
                # Argument variable
                arg_var = "${}".format(ind)

                # Set argument variable
                utils.MetaCommand("opt var add var{} {}".format(ind, arg_var))

                # Get argument variable
                arg_value = utils.MetaGetVariable("var{}".format(ind))

                # Check if argument variable is same as argument value
                if arg_value != arg_var:
                    args.append(arg_value)
                else:
                    args.append("")

            if any(item == "" for item in args[:3]):
                #print("Insufficient input arguments. User should at least provide the load case type, run name and results directory")
                return -1

            return args

        return 0

    def run_interactive_mode(self):
        """This method get the user input from the command line interactive mode.

        Returns:
            int: 0 as always
        """

        self.metadb_2d_input = self.set_param_from_cmd_input("Please enter the 2d metadb file    ",)

        return 0

    @staticmethod
    def set_param_from_cmd_input(input_str, options_list=None):
        """
        set_param_from_cmd_input [summary]

        [extended_summary]

        Args:
            input_str ([type]): [description]
            options_list ([type], optional): [description]. Defaults to None.

        Returns:
            [type]: [description]
        """

        option_str = ""
        if options_list:
            for ind, option in options_list.items():
                if ind.startswith("empty"):
                    option_str += "\n"
                else:
                    option_str += "{} {} \n".format(ind, option)

        try:
            ret = input("{}{}:".format(option_str, input_str)).strip()
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
