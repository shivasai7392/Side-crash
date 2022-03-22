#PYTHON SCRIPT

import os

from meta import utils
from meta import windows
from meta import constants

from src.user_input import UserInput
from src.metadb_info import Meta2DInfo
from src.metadb_info import GeneralVarInfo
from src.metadb_info import Meta3DInfo
from src.generate_reports.reporter import Reporter
from src.general_utilities import append_libs_path
from src.logger import SideCrashLogger

def main(*args):
    """
    apply [summary]

    [extended_summary]

    Returns:
        [type]: [description]
    """
    append_libs_path()
    app_config_dir = os.path.join(constants.app_root_dir,"config")
    gui_mode = utils.MetaGetVariable("side_crash_toolbar_gui_mode")
    user_input = UserInput(*args)

    if gui_mode == "True":
        user_input.get_user_input_from_gui()
    else:
        user_input.get_user_input_from_interactive_mode()

    windows.CollectNewWindowsStart()
    utils.MetaCommand("read project {}".format(user_input.metadb_2d_input))

    general_input_info = GeneralVarInfo()
    log_file = os.path.join(app_config_dir,"res",general_input_info.get_log_directory().replace("/","",1)).replace("\\",os.sep)
    logger = SideCrashLogger(log_file)

    general_input_info = general_input_info.get_info()
    metadb_2d_input_info = Meta2DInfo().get_info()
    metadb_3d_input_info = Meta3DInfo().get_info()


    threed_metadb_file = os.path.join(app_config_dir,"res",general_input_info.threed_metadb_file.replace("/","",1)).replace("\\",os.sep)
    utils.MetaCommand('read project overlay "{}" ""'.format(threed_metadb_file))

    d3hsp_file_path = os.path.join(app_config_dir,"res",general_input_info.d3hsp_file.replace("/","",1)).replace("\\",os.sep)
    metadb_3d_input_info.get_spotweld_clusters(d3hsp_file_path)

    utils.MetaCommand('read options boundary materials')
    utils.MetaCommand('read dis MetaDB {} {},{} lossy_compressed:0:Displacements'.format(threed_metadb_file,general_input_info.peak_state_value,general_input_info.final_state_value))
    utils.MetaCommand('read onlyfun MetaDB {} {},{} lossy_compressed:0:MetaResult::Stresses(ECS),,PlasticStrain'.format(threed_metadb_file,general_input_info.peak_state_value,general_input_info.final_state_value))

    utils.MetaCommand('options title off')
    utils.MetaCommand('options axis off')
    utils.MetaCommand('0:options state variable "serial=1"')
    utils.MetaCommand('grstyle scalarfringe enable')
    utils.MetaCommand('options fringebar on')
    utils.MetaCommand('srange window "MetaPost" set .0,.15')
    utils.MetaCommand('opt fringe visibility novaluecolor enabled off')
    utils.MetaCommand('color fringebar update "StressTensor" "Red,255_92_0_255,255_185_0_255,231_255_0_255,139_255_0_255,46_255_0_255,0_255_46_255,0_255_139_255,0_255_231_255,0_185_255_255,White,LightGray"')
    utils.MetaCommand('color fringebar scalarset StressTensor')
    utils.MetaCommand('0:options state original')

    new_windows = windows.CollectNewWindowsEnd()

    metadb_2d_input_info.prepare_info(new_windows)

    reporter = Reporter(new_windows,general_input_info,metadb_2d_input_info,metadb_3d_input_info,app_config_dir)
    reporter.run_process()

    logger.log.info("""Log Session Report End
--------------------------
--------------------------

""")
    logger.shutdown()

    return 0
