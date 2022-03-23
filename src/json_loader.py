# PYTHON script
"""
##################################################
#      Copyright BETA CAE Systems USA Inc.,      #
#      2020 All Rights Reserved                  #
#      UNPUBLISHED, LICENSED SOFTWARE.           #
##################################################

Developer   : Shiva sai krishna Thota
Date        : Apr 30, 2021

This method is used for loading json file for meta post procesing.

"""

import re
import json
from meta import constants
from collections import OrderedDict

class JsonLoader():
    """
    This class is used for loading the json from the specified user directory.

    """

    def __init__(self, json_file):
        self.json_file = json_file
        self.json_dict = dict()

    def load_json(self):
        """
        This method is used the json file as dictionary.

        Return :
            dict: json file data dictionary.
        """
        try:
            with open(self.json_file) as some_file:
                self.json_dict = json.load(some_file, object_pairs_hook=OrderedDict)
            return self.format_json_values(self.json_dict)
        except:
            return 0

    def format_json_values(self, json_dict):
        """
        This method is used to format the key value pairs from the json file.

        Args:
            json_dict ([dict]): dictionary with json file key,value pairs

        Return :
            dict : json file data dictionary.
        """
        #patterns to identity the values
        integer_pattern = re.compile("^[-+]?\d*$")
        bool_pattern = re.compile("^[10]$")
        float_pattern = re.compile("^[-+]?\d+[.]\d*$")
        plot_curves_colors_pattern = re.compile("^([\[].*[\]])(,[\[].*[\]])*$")
        list_pattern = re.compile("^[\w\d\-!@#$%^&*()_+|~=`{}:\";'<>?.\/\\\\]+(,[\w\d\-!@#$%^&*()_+|~=`{}:\";'<>?.\/\\\\]+)*$")
        range_pattern = re.compile("^\d+-\d+$")
        #iterating through the items of the json data dictionary
        for key in list(json_dict.keys()):
            value = json_dict[key]
            #ignoring the comments from the json
            if key.startswith("#"):
                del json_dict[key]
            #recurring exceution of the method if value is of type dictionary
            elif isinstance(value, dict):
                self.format_json_values(value)
            #else formatting the values
            else:
                value = str(value).replace("$META_POST_HOME",str(constants.app_root_dir).replace("\\","/"))
                if  value is not None and value != "" and value != "None":
                    if bool_pattern.match(value):
                        continue
                    elif integer_pattern.match(value):
                        json_dict[key] = int(value)
                    elif float_pattern.match(value):
                        json_dict[key] = float(value)
                    elif plot_curves_colors_pattern.match(value):
                        json_dict[key] = eval(value)
                    elif list_pattern.match(value):
                        list_values = value.split(",")
                        range_values = []
                        for index,each_value in enumerate(list_values):
                            if range_pattern.match(each_value):
                                split_values = each_value.split("-")
                                range_values.extend([i for i in range(int(split_values[0]),int(split_values[1])+1)])
                                list_values[index]="remove"
                            elif integer_pattern.match(each_value):
                                list_values[index] = int(each_value)
                            elif float_pattern.match(each_value):
                                list_values[index] = float(each_value)
                        list_values = [i for i in list_values if i != "remove"]
                        list_values.extend(range_values)
                        json_dict[key] = list_values[0] if len(list_values) == 1 else list_values

        return json_dict
