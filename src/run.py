#PYTHON SCRIPT
"""
     _summary_

    _extended_summary_

    Returns:
        _type_: _description_
"""
import os
from datetime import datetime
import sys

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
        [Int]: 0 always
    """
    utils.MetaCommand('options session controldraw disable')
    messenger = utils.Messenger()
    append_libs_path()
    # Getting the config directory filepath
    app_config_dir = os.path.join(constants.app_root_dir,"config")
    gui_mode = utils.MetaGetVariable("side_crash_toolbar_gui_mode")
    user_input = UserInput(*args)

    # getting the 2d metadb input and running the script in windows and linux system
    if gui_mode == "True":
        user_input.get_user_input_from_gui()
    else:
        user_input.get_user_input_from_json()

    # Creating the new windows
    windows.CollectNewWindowsStart()

    # Reading the 2d metadb input
    messenger.start_buffering()
    twod_start_time = datetime.now()
    utils.MetaCommand("read project {}".format(user_input.metadb_2d_input))
    messenger.stop_buffering()
    buffer = messenger.get_buffer()

    # Joining the config directory path and log directory path
    general_input_info = GeneralVarInfo()
    if "win" in sys.platform:
        log_file = os.path.join(app_config_dir,"res",general_input_info.get_log_directory().replace("/","",1)).replace("\\",os.sep)
    else:
        log_file = os.path.join(general_input_info.get_log_directory())
    logger = SideCrashLogger(log_file)
    logger.log.info("--- STARTED LOADING 2D METADB FILE")
    logger.log.info("2D METADB FILE PATH : {}".format(user_input.metadb_2d_input))
    #_ret = [logger.log.info(item) for item in buffer if item.strip() != ""]
    logger.log.info("TIME TAKEN FOR LOADING 2D METADB FILE : {}".format(datetime.now() - twod_start_time))
    logger.log.info("")

    logger.log.info("--- STARTED OVERLAYING TARGET METADB FILE")
    logger.log.info("TARGET METADB FILE PATH : {}".format(user_input.target_metadb_input))
    #overlaying target metdb file
    target_start_time = datetime.now()
    general_input_info.target_2d_metadb = user_input.target_metadb_input
    messenger.start_buffering()
    utils.MetaCommand('read project overlay "{}" ""'.format(general_input_info.target_2d_metadb))
    messenger.stop_buffering()
    buffer = messenger.get_buffer()
    #_ret = [logger.log.info(item) for item in buffer if item.strip() != ""]
    logger.log.info("TIME TAKEN FOR OVERLAYING TARGET METADB FILE : {}".format(datetime.now() - target_start_time))
    logger.log.info("")

    # Getting the meta2dinfo and meta3d info
    general_input_info = general_input_info.get_info()
    metadb_2d_input_info = Meta2DInfo().get_info()
    metadb_3d_input_info = Meta3DInfo().get_info()

    # Joining the config file and 3dmetadb file
    # threed_metadb_file = os.path.join(app_config_dir,"res",general_input_info.threed_metadb_file.replace("/","",1)).replace("\\",os.sep)
    # Reading the threed_metadb_file
    logger.log.info("--- STARTED OVERLAYING AND READING PEAK,FINAL STATE RESULTS FROM 3D METADB FILE")
    logger.log.info("3D METADB FILE PATH : {}".format(user_input.metadb_3d_input))
    threed_start_time = datetime.now()
    general_input_info.target_3d_metadb = user_input.metadb_3d_input
    messenger.start_buffering()
    utils.MetaCommand('read project overlay "{}" ""'.format(general_input_info.target_3d_metadb))
    messenger.stop_buffering()
    buffer = messenger.get_buffer()
    #_ret = [logger.log.info(item) for item in buffer if item.strip() != ""]
    logger.log.info("TIME TAKEN FOR OVERLAYING 3D METADB FILE : {}".format(datetime.now() - threed_start_time))
    #reading the results
    messenger.start_buffering()
    utils.MetaCommand('read options boundary materials')
    utils.MetaCommand('read dis MetaDB {} {},{} lossy_compressed:0:Displacements'.format(user_input.metadb_3d_input,general_input_info.peak_state_value,general_input_info.final_state_value))
    utils.MetaCommand('read onlyfun MetaDB {} {},{} lossy_compressed:0:MetaResult::Stresses(ECS),,PlasticStrain'.format(user_input.metadb_3d_input,general_input_info.peak_state_value,general_input_info.final_state_value))
    messenger.stop_buffering()
    buffer = messenger.get_buffer()
    #_ret = [logger.log.info(item) for item in buffer if item.strip() != ""]
    logger.log.info("TIME TAKEN FOR READING PEAK,FINAL STATE RESULTS FROM 3D METADB FILE : {}".format(datetime.now() - threed_start_time))
    logger.log.info("")

    # # Joining the config directory path and d3hsp file path for d3hsp file
    # d3hsp_file_path = os.path.join(app_config_dir,"res",general_input_info.d3hsp_file.replace("/","",1)).replace("\\",os.sep)
    # Getting the spotweld clusters from d3hsp file
    logger.log.info("--- STARTED SPOTWELD CLUSTERS IDENTIFICATION")
    if "win" in sys.platform:
        general_input_info.binout_directory = os.path.join(app_config_dir,"res",general_input_info.binout_directory.replace("/","",1)).replace("\\",os.sep)
    else:
        general_input_info.binout_directory = os.path.join(general_input_info.binout_directory)
    general_input_info.d3hsp_file_path = user_input.d3hsp_file_path
    logger.log.info("SPOTWELD ID'S SOURCE FILE PATH : {}".format(general_input_info.d3hsp_file_path))
    spotweld_start_time = datetime.now()
    metadb_3d_input_info.get_spotweld_clusters(general_input_info.d3hsp_file_path)
    logger.log.info("--- SPOTWELD CLUSTERS IDENTIFICATION IS COMPLETED")
    logger.log.info("TIME TAKEN : {}".format(datetime.now() - spotweld_start_time))
    logger.log.info("")

    #setting global 3d settings
    utils.MetaCommand('options title off')
    utils.MetaCommand('options axis off')
    utils.MetaCommand('0:options state variable "serial=1"')
    utils.MetaCommand('grstyle scalarfringe enable')
    utils.MetaCommand('options fringebar on')
    utils.MetaCommand('srange window "{}" set .0,.15'.format(general_input_info.threed_window_name))
    utils.MetaCommand('opt fringe visibility novaluecolor enabled off')
    utils.MetaCommand('color fringebar update "StressTensor" "Red,255_92_0_255,255_185_0_255,231_255_0_255,139_255_0_255,46_255_0_255,0_255_46_255,0_255_139_255,0_255_231_255,0_185_255_255,White,White"')
    utils.MetaCommand('color fringebar scalarset StressTensor')
    utils.MetaCommand('0:options state original')

    # Ends the Collecting the new windows
    new_windows = windows.CollectNewWindowsEnd()

    # Preparing the 2d metadb information
    metadb_2d_input_info.prepare_info(new_windows)

    # Calling the Reporter Class and executing the run_process function
    logger.log.info("--- SIDE CRASH REPORTING FUNCTIONALITY STARTED")
    logger.log.info("")
    report_start_time = datetime.now()
    reporter = Reporter(new_windows,general_input_info,metadb_2d_input_info,metadb_3d_input_info,app_config_dir)
    reporter.run_process()
    logger.log.info("SIDE CRASH REPORTING FUNCTIONALITY COMPLETED")
    logger.log.info("TOTAL TIME TAKEN : {}".format(datetime.now() - report_start_time))

    logger.log.info("""Log Session Report End
--------------------------
--------------------------

""")
    logger.shutdown()
    utils.MetaCommand('options session controldraw enable')

    return 0
