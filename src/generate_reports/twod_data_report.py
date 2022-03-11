# PYTHON script
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""

import os

from meta import utils

from src.meta_utilities import visualize_3d_critical_section,annotation
from src.generate_reports.excel_generator import ExcelBomGeneration


class TwoDDataReport():
    """
    __init__ _summary_

    _extended_summary_

    Args:
        metadb_3d_input (_type_): _description_
        threed_images_report_folder (_type_): _description_
        thrred_videos_report_folder (_type_): _description_
    """

    def __init__(self,
                metadb_2d_input,
                twod_images_report_folder,
                logger) -> None:

        self.metadb_2d_input = metadb_2d_input
        self.twod_images_report_folder = twod_images_report_folder
        self.logger = logger

    def run_process(self):
        """
        run_process _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """

        from PIL import ImageFile
        ImageFile.LOAD_TRUNCATED_IMAGES = True

        window_2d_objects = self.metadb_2d_input.window_objects
        for window in window_2d_objects:
            window_name = window.name
            window_layout = window.meta_obj.get_plot_layout()
            plot = window.plot
            curve = plot.curve
            utils.MetaCommand('window active "{}"'.format(window_name))
            utils.MetaCommand('window maximize "{}"'.format(window_name))
            utils.MetaCommand('xyplot plotdeactive "{}" all'.format(window_name))
            curve.meta_obj.show()
            utils.MetaCommand('xyplot plotactive "{}" {}'.format(window_name, plot.id))
            utils.MetaCommand('xyplot curve visible and "{}" selected'.format(window_name))
            utils.MetaCommand('xyplot rlayout "{}" 1'.format(window_name))
            image_path = os.path.join(self.twod_images_report_folder,window_name+"_"+curve.name.lower()+".png")

            utils.MetaCommand('write png "{}"'.format(image_path))

            utils.MetaCommand('xyplot rlayout "{}" {}'.format(window_name,window_layout))

        return 0
