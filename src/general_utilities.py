# PYTHON script

import os
import time
import functools
import sys

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