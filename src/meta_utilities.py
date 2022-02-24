# PYTHON script
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""

import os

from meta import utils
from meta import windows

class Meta2DWindow():
    """
    Meta2DWindow [summary]

    [extended_summary]
    """
    def __init__(self,name,obj) -> None:
        self.name = name
        self.meta_obj = obj
        self.plot = None

    class Plot():
        """
        Plot [summary]

        [extended_summary]
        """
        def __init__(self,id,obj) -> None:
            self.id = id
            self.meta_obj = obj
            self.curve = None

        class Curve():
            """
            Curve [summary]

            [extended_summary]
            """
            def __init__(self,id,name,obj) -> None:
                self.id = id
                self.name = name
                self.meta_obj = obj

def capture_image(window_name,width,height,file_path,plot_id = None,rotate = None, view = None):
    """
    capture_image _summary_

    _extended_summary_

    Args:
        window_name (_type_): _description_
        width (_type_): _description_
        height (_type_): _description_
        file_path (_type_): _description_
        plot_id (_type_, optional): _description_. Defaults to None.

    Returns:
        _type_: _description_
    """
    from PIL import ImageGrab

    win_obj = windows.Window(window_name, page_id = 0)
    win_obj.set_size((round(width/9525),round(height/9525)))

    if view is not None:
        utils.MetaCommand('view default {}'.format(view))

    if plot_id is not None:
        utils.MetaCommand('clipboard copy plot image "{}" {}'.format(window_name, plot_id))
    else:
        utils.MetaCommand('clipboard copy image "{}"'.format(window_name))

    img = ImageGrab.grabclipboard()
    rgba_img = image_transperent(img)
    if rotate:
        rgba_img = rgba_img.transpose(rotate)

    if not os.path.exists(os.path.dirname(file_path)):
        os.makedirs(os.path.dirname(file_path))
    rgba_img.save(file_path, 'PNG')

    utils.MetaCommand('window maximize {}'.format(window_name))

    return 0

def image_transperent(img):
    """
    image_transperent _summary_

    _extended_summary_

    Args:
        img (_type_): _description_

    Returns:
        _type_: _description_
    """
    rgba_img = img.convert("RGBA")
    rgba_data = rgba_img.getdata()
    new_rgba_data = []
    for item in rgba_data:
        if item[0] == 255 and item[1] == 255 and item[2] == 255:
            new_rgba_data.append((255, 255, 255, 0))
        else:
            new_rgba_data.append(item)
    rgba_img.putdata(new_rgba_data)

    return rgba_img


def capture_resized_image(window_name,width,height,file_path,plot_id = None,rotate = None, view = None):
    """
    capture_resized_image _summary_

    _extended_summary_

    Args:
        window_name (_type_): _description_
        width (_type_): _description_
        height (_type_): _description_
        file_path (_type_): _description_
        plot_id (_type_, optional): _description_. Defaults to None.
        rotate (_type_, optional): _description_. Defaults to None.
        view (_type_, optional): _description_. Defaults to None.

    Returns:
        _type_: _description_
    """
    from PIL import ImageGrab

    if view is not None:
        utils.MetaCommand('view default {}'.format(view))

    if plot_id is not None:
        utils.MetaCommand('clipboard copy plot image "{}" {}'.format(window_name, plot_id))
    else:
        utils.MetaCommand('clipboard copy image "{}"'.format(window_name))

    img = ImageGrab.grabclipboard()
    img = img.resize((round(width/9525),round(height/9525)))
    rgba_img = image_transperent(img)
    if rotate:
        rgba_img = rgba_img.transpose(rotate)

    if not os.path.exists(os.path.dirname(file_path)):
        os.makedirs(os.path.dirname(file_path))
    rgba_img.save(file_path, 'PNG')

    utils.MetaCommand('window maximize {}'.format(window_name))

    return 0
