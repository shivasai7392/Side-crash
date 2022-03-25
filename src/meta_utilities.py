# PYTHON script
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""

import logging
from datetime import datetime

from meta import utils
from meta import windows
from meta import models
from meta import constants
from meta import plot2d
from meta import groups
from meta import annotations

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

def capture_image(file_path,window_name,width,height,rotate = None, view = None,transparent = False):
    """
    This method is used to capture an image of the resized meta window based on the width and height.

    Args:
        file_path (str): path to save image
        window_name (str): meta window name
        width (int): width of the shape
        height (int): height of the shape
        rotate (object): Image rotate object
        view (str): view position for the contents of the window
        transparent (bool): status of transparency

    Returns:
        int: 0 Always for Success,1 for Failure.
    """
    from PIL import Image,ImageFile
    ImageFile.LOAD_TRUNCATED_IMAGES = True
    try:
        #maximizing the window
        utils.MetaCommand('window maximize {}'.format(window_name))
        #getting window object to resize
        win_obj = windows.Window(window_name, page_id = 0)
        win_obj.set_size((round(width/9525),round(height/9525)))
        #applying view to the window
        if view is not None:
            utils.MetaCommand('view default {}'.format(view))
        #saving image of the meta window
        utils.MetaCommand('write png "{}"'.format(file_path))
        #rotating the image based on rotate object and saving it
        if rotate:
            img = Image.open(file_path)
            img = img.transpose(rotate)
            img.save(file_path, 'PNG')
            img.close()
        if transparent:
            img = Image.open(file_path)
            img.save(file_path, 'PNG')
            img.close()
            img = Image.open(file_path)
            img = image_transperent(img)
            img.save(file_path.replace(".png","")+"_transparent.png", 'PNG')
            img.clode()
        #maximizing the window
        utils.MetaCommand('window maximize {}'.format(window_name))
    except:
        return 1
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


def capture_image_and_resize(file_path,width,height,rotate = None,transparent = False):
    """
    This method is used to capture an image of the meta window and resize it based on the width and height.

    Args:
        window_name (str): meta window name
        width (int): width of the shape
        height (int): height of the shape
        file_path (str): path to save image

    Returns:
        int: 0 Always for Success,1 for Failure.
    """
    from PIL import Image,ImageFile
    ImageFile.LOAD_TRUNCATED_IMAGES = True

    try:
        #saving image of the meta window
        utils.MetaCommand('write png "{}"'.format(file_path))
        #creating Image object for the saved image
        img = Image.open(file_path)
        #resizing the image
        img = img.resize((round(width/9525),round(height/9525)))
        #saving the resized image
        img.save(file_path, 'PNG')
        img.close()
        #rotating the image based on rotate object and saving it
        if rotate:
            img = Image.open(file_path)
            img = img.transpose(rotate)
            img.save(file_path, 'PNG')
            img.close()
        if transparent:
            img = Image.open(file_path)
            img.save(file_path, 'PNG')
            img.close()
            img = Image.open(file_path)
            img = image_transperent(img)
            img.save(file_path.replace(".png","")+"_transparent.png", 'PNG')
            img.close()
    except:
        return 1

    return 0

def visualize_3d_critical_section(data,and_filter = None):
    """
    visualize_3d_critical_section _summary_

    _extended_summary_

    Returns:
        _type_: _description_
    """
    get_var = lambda key: data[key] if key in data.keys() else None

    prop_names = get_var("hes")
    hes_exceptions = get_var("hes_exceptions")
    exclude = "null"
    erase_pids = get_var("erase_pids")
    comp_view = get_var("view")
    transparency_level = '50'
    transparent_pids = get_var("transparent_pids")
    erase_box = get_var("erase_box")

    if prop_names:
        if and_filter:
            utils.MetaCommand('add advfilter partoutput add:Parts:name:{}:Keep All'.format(prop_names))
        else:
            utils.MetaCommand('or advfilter partoutput add:Parts:name:{}:Keep All'.format(prop_names))
    if hes_exceptions:
        utils.MetaCommand('add pid {}'.format(hes_exceptions))
    utils.MetaCommand('erase advfilter partoutput add:Parts:name:{}:Keep All'.format(exclude))
    if erase_pids:
        utils.MetaCommand('erase pid {}'.format(erase_pids))
    if erase_box:
        utils.MetaCommand('erase shells box {}'.format(erase_box))
        utils.MetaCommand('erase solids box {}'.format(erase_box))
    if comp_view:
        utils.MetaCommand('view default {}'.format(comp_view))
        utils.MetaCommand('view center')
    if transparent_pids:
        utils.MetaCommand('color pid transparency {} {}'.format(transparency_level,transparent_pids))

    return 0

def visualize_annotation(spotweld_id_elements,bins_path):
    """
        annotation

        _extended_summary_
    """
    logger = logging.getLogger("side_crash_logger")
    start_time = datetime.now()
    utils.MetaCommand('add element connected')
    utils.MetaCommand('add element connected')
    meta_post_window_object = windows.Window(name = 'MetaPost', page_id=0)

    text_rgb = "Black"
    text_rgb_values = windows.RgbFromNamedColor(text_rgb)
    text_color = windows.Color(text_rgb_values[0], text_rgb_values[1], text_rgb_values[2],text_rgb_values[3])
    marginal = "Orange"
    marginal_rgb_values = windows.RgbFromNamedColor(marginal)
    marginal_color = windows.Color(marginal_rgb_values[0], marginal_rgb_values[1], marginal_rgb_values[2],marginal_rgb_values[3])
    bad = "red"
    bad_rgb_values = windows.RgbFromNamedColor(bad)
    bad_color = windows.Color(bad_rgb_values[0], bad_rgb_values[1], bad_rgb_values[2],bad_rgb_values[3])

    model_get = models.Model(0)
    visible_elements = model_get.get_elements('visible', window =meta_post_window_object, element_type = constants.SOLID )
    clusters = []
    group_start_time = datetime.now()
    identified_elements = []
    for key,values in spotweld_id_elements.items():
        if any(each_element.id in values for each_element in visible_elements):
            clusters.append(key)
            identified_elements.extend(values)
        if all(visible_id in identified_elements for visible_id in visible_elements):
            break

    group_end_time = datetime.now()
    logger.info("BINOUT'S DIRECTORY PATH : {}".format(bins_path))
    logger.info("SPOTWELD ID IDENTIFICATION AND CLUSTER GROUP GENERATION AVERAGE TIME : {}".format(group_end_time-group_start_time))

    curve_start_time = datetime.now()
    utils.MetaCommand('xyplot create "Temporary Window"')
    utils.MetaCommand('xyplot read lsdyna "Temporary Window" "{}" swforc-SpotweldAssmy {}  failure_(f)'.format(bins_path,",".join(str(key) for key in clusters)))
    curve_end_time = datetime.now()
    logger.info("CURVES GENERATION AVERAGE TIME : {}".format(curve_end_time-curve_start_time))

    plot = plot2d.Plot(0,"Temporary Window",0)
    curves = plot.get_curves('all')
    meta_post_window_object.maximize()
    annot_start_time = datetime.now()
    failed_welds = 0
    annots = meta_post_window_object.get_annotations('all')
    annotation_id = len(annots)+1
    for curve in curves:
        failure_point = plot2d.MaxPointYOfCurve("Temporary Window", curve.id, 'real')
        failure_value = str(round(failure_point.y,2))
        if float(failure_value) > 0.7:
            failed_welds += 1
            failure_time = failure_point.x
            failure_time = str(round(failure_time,3))
            #Create an annotation in the 3D data for the cluster which is above the threshold value
            annotation_label = failure_value+' @ '+failure_time
            annotation_group = 'spotweld_cluster_'+str(curve.entity_id)
            g = groups.Group(annotation_group,0)
            annotation_id += 1
            a = annotations.CreateEmptyAnnotation("MetaPost",annotation_label,annotation_id)
            utils.MetaCommand('annotation line {} width 1'.format(annotation_id))
            utils.MetaCommand('annotation text {} font "MS Shell Dlg 2,8,-1,5,75,0,0,0,0,0"'.format(annotation_id))
            utils.MetaCommand('annotation border {} padding 3'.format(annotation_id))
            a.set_group(g)

            if float(failure_value) > 0.9:
                a.set_background_color(bad_color)
            else:
                a.set_background_color(marginal_color)
            a.set_border_color(text_color)

    utils.MetaCommand('window active MetaPost')
    utils.MetaCommand('annotation explode all center 10')
    utils.MetaCommand('annotation extparam all shape off')
    utils.MetaCommand('annotation text all format auto')
    utils.MetaCommand('window delete "Temporary Window"')

    annot_end_time = datetime.now()
    logger.info("CURVES MAX DETERMINATION AND ANNOTATIONS GENERATION AVERAGE TIME : {}".format(annot_end_time-annot_start_time))
    logger.info("PROCESSED WELDS : {} | WELDS ABOVE THRESHOLD : {} | TOTAL WELD IDENTIFICATION AND ANNOTATIONS ADD TIME : {}".format(len(curves),failed_welds,annot_end_time - start_time))
    logger.info("")

    return 0

def deformation_plot_formmatter(window_name,plot1_id,plot2_id,plot3_id):
    """
    deformation_plot_formmatter _summary_

    _extended_summary_

    Returns:
        _type_: _description_
    """
    utils.MetaCommand('xyplot axisoptions yaxis active "{}" {} 0'.format(window_name,plot1_id))
    utils.MetaCommand('xyplot axisoptions ylabel font "{}" {} "Arial,12,-1,5,75,0,0,0,0,0"'.format(window_name,plot1_id))
    utils.MetaCommand('xyplot axisoptions yaxis deactive "{}" {} 0'.format(window_name,plot1_id))

    utils.MetaCommand('xyplot axisoptions yaxis active "{}" {} 0'.format(window_name,plot2_id))
    utils.MetaCommand('xyplot axisoptions ylabel font "{}" {} "Arial,12,-1,5,75,0,0,0,0,0"'.format(window_name,plot2_id))
    utils.MetaCommand('xyplot axisoptions yaxis deactive "{}" {} 0'.format(window_name,plot2_id))

    utils.MetaCommand('xyplot axisoptions yaxis active "{}" {} 0'.format(window_name,plot3_id))
    utils.MetaCommand('xyplot axisoptions ylabel font "{}" {} "Arial,12,-1,5,75,0,0,0,0,0"'.format(window_name,plot3_id))
    utils.MetaCommand('xyplot axisoptions yaxis deactive "{}" {} 0'.format(window_name,plot3_id))

    utils.MetaCommand('xyplot axisoptions xaxis active "{}" {} 0'.format(window_name,plot1_id))
    utils.MetaCommand('xyplot axisoptions xlabel font "{}" {} "Arial,12,-1,5,75,0,0,0,0,0"'.format(window_name,plot1_id))
    utils.MetaCommand('xyplot axisoptions xaxis deactive "{}" {} 0'.format(window_name,plot1_id))

    utils.MetaCommand('xyplot axisoptions xaxis active "{}" {} 0'.format(window_name,plot2_id))
    utils.MetaCommand('xyplot axisoptions xlabel font "{}" {} "Arial,12,-1,5,75,0,0,0,0,0"'.format(window_name,plot2_id))
    utils.MetaCommand('xyplot axisoptions xaxis deactive "{}" {} 0'.format(window_name,plot2_id))

    utils.MetaCommand('xyplot axisoptions xaxis active "{}" {} 0'.format(window_name,plot3_id))
    utils.MetaCommand('xyplot axisoptions xlabel font "{}" {} "Arial,12,-1,5,75,0,0,0,0,0"'.format(window_name,plot3_id))
    utils.MetaCommand('xyplot axisoptions xaxis deactive "{}" {} 0'.format(window_name,plot3_id))

    utils.MetaCommand('xyplot plotoptions title font "{}" {} "Arial,14,-1,5,75,0,0,0,0,0"'.format(window_name,plot1_id))
    utils.MetaCommand('xyplot plotoptions title font "{}" {} "Arial,14,-1,5,75,0,0,0,0,0"'.format(window_name,plot2_id))
    utils.MetaCommand('xyplot plotoptions title font "{}" {} "Arial,14,-1,5,75,0,0,0,0,0"'.format(window_name,plot3_id))
    return 0
