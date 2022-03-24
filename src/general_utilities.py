# PYTHON script
"""
_summary_

_extended_summary_

Returns:
    _type_: _description_
"""

import os
import time
import functools
import sys

def closest(list_of_values, value):
    """
    closest _summary_

    _extended_summary_

    Args:
        lst (_type_): _description_
        K (_type_): _description_

    Returns:
        _type_: _description_
    """
    if type(value) == int:
        nearest_value = list_of_values[min(range(len(list_of_values)), key = lambda i: abs(list_of_values[i]-value))]
        return nearest_value

    else:
        return None

def meta_block_redraws(func1):
    """ This decorator blocks redraws before
    the decorated function is called and
    enables again before returning
    """

    @functools.wraps(func1)
    def func2(*args, **kwargs):
        """
        This is just a wrapper for the func
        """
        from meta import utils

        # Disable redraw
        utils.MetaCommand("options session controldraw disable")

        # Execute the decorated function
        ret = func1(*args, **kwargs)

        # Enable Redraw
        utils.MetaCommand("options session controldraw enable")

    return func2

def meta_create_delete_3d(func):
    """ This decorator blocks redraws before
    the decorated function is called and
    enables again before returning
    """

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        """
        This is just a wrapper for the func
        """
        from meta import utils

        # Create a window
        utils.MetaCommand('window create "MetaPost"')

        try:
            # Execute the decorated function
            ret = func(*args, **kwargs)
        finally:
            # Create a window
            utils.MetaCommand('window delete "MetaPost"')

        return ret

    return wrapper

def remove_dup(list_obj):
    """This method removes duplicate of a list whilst preserving order using short-circuiting and microoptimization.

    Args:
        list_obj (list): A list

    Returns:
        list: new list
    """
    found = set()
    found_add = found.add

    return [x for x in list_obj if not (x in found or found_add(x))]

def check_for_int(input_str):
    """[summary]

    Args:
        input_str ([type]): [description]

    Returns:
        [type]: [description]
    """
    try:
        int(input_str)
        return True
    except ValueError:
        return False

def check_for_float(input_str):
    """[summary]

    Args:
        input_str ([type]): [description]

    Returns:
        [type]: [description]
    """
    try:
        float(input_str)
        return True
    except ValueError:
        return False

def append_libs_path():
    """ Does the following:
    1. Gets the path to site-packages
    2. Appends it to sys.path

    NOTE: The site-packages folder is situated
    under the appropriate OS folder in docs/

    Returns:
        int: 0 always.
    """

        # Get the libs path
    libs_path = os.path.abspath(
        os.path.join(
            os.path.dirname(os.path.dirname(__file__)),
            "libs",
        )
    )

    # Get site-packages path
    if "win" in sys.platform:
        site_pkgs_path = os.path.join(libs_path, "win64", "Lib_py38")
    else:
        site_pkgs_path = os.path.join(libs_path, "linux", "Lib_py38", "python3.8", "site-packages")

    # Append the path to sys.path
    if site_pkgs_path not in sys.path:
        sys.path.append(site_pkgs_path)

    return 0

def add_row(table):
    """
    add_row [summary]

    [extended_summary]

    Args:
        table ([type]): [description]

    Returns:
        [type]: [description]
    """
    from pptx.table import _Cell
    from copy import deepcopy

    new_row = deepcopy(table._tbl.tr_lst[-1])
    # duplicating last row of the table as a new row to be added

    table._tbl.append(new_row)

    return 0

def remove_row(table,row_to_delete):
    """
    remove_row [summary]

    [extended_summary]

    Args:
        table ([type]): [description]
        row_to_delete ([type]): [description]
    """
    table._tbl.remove(row_to_delete._tr)


def clone_shape(shape):
    """
    Add a duplicate of `shape` to the slide on which it appears.

    """
    import copy
    from pptx.shapes.autoshape import Shape
    # ---access required XML elements---
    sp = shape._sp
    spTree = sp.getparent()
    # ---clone shape element---
    new_sp = copy.deepcopy(sp)
    # ---add it to slide---
    spTree.append(new_sp)
    # ---create a proxy object for the new sp element---
    new_shape = Shape(new_sp, None)
    new_shape.left = shape.left
    new_shape.top = shape.top

    return new_shape
