# PYTHON script
"""
##################################################
#      Copyright BETA CAE Systems USA Inc.,      #
#      2020 All Rights Reserved                  #
#      UNPUBLISHED, LICENSED SOFTWARE.           #
##################################################


Loads the user input file and Fetches all meta db info and storing in instances

Developer   : Naresh Medipalli and Shiva Sai
Date        : Dec 31, 2021

"""

import os

from meta import utils
from meta import models
from meta import plot2d

from src.meta_utilities import Meta2DWindow

class Meta2DInfo:
    """ This module contains complete variable info of user provided 2d data .
    """

    re_curve_name = "curve"

    def __init__(self):
        self.curves = {}
        self.window_objects = []

    def get_info(self):
        """[summary]

        [extended_summary]

        Returns:
            [type]: [description]
        """
        curve_variables = utils.MetaGetVariablesByName("*_{}".format(Meta2DInfo.re_curve_name))
        for one_var in curve_variables :
            var_name = one_var[0]
            if var_name[0].isdigit():
                continue
            else:
                var_value = one_var[1]
                if not var_value.isdigit():
                    self.curves[var_name] = var_value
                else:
                    self.curves[var_name] = int(var_value)

        return self

    def prepare_info(self,windows):
        """
        prepare_info [summary]

        [extended_summary]

        Returns:
            [type]: [description]
        """
        for win in windows:
            curves = win.get_curves('all')
            for a_curve in self.curves.values():
                curve_names = [curve.name for curve in curves]
                if a_curve in curve_names:
                    index = curve_names.index(a_curve)
                    curve_meta_obj = curves[index]
                    win_obj = Meta2DWindow(win.name,win)
                    plot_meta_obj = plot2d.PlotById(win.name, curve_meta_obj.plot_id)
                    plot_obj = win_obj.Plot(plot_meta_obj.id,plot_meta_obj)
                    win_obj.plot = plot_obj
                    curve_obj = plot_obj.Curve(curve_meta_obj.id,curve_meta_obj.name,curve_meta_obj)
                    plot_obj.curve = curve_obj
                    self.window_objects.append(win_obj)

        return 0

class Meta3DInfo:
    """ This module contains complete variable info of user provided 3d data .
    """

    re_erase_box = "erase_box"
    re_erase_pids = "erase_pids"
    re_hes = "hes"
    re_hes_exceptions = "hes_exceptions"
    re_name = "name"
    re_transparent_pids = "transparent_pids"
    re_view = "view"

    def __init__(self):

        self.critical_sections = {}
        self.properties = models.CollectModelEntities(0, "parts")


    def get_info(self):
        """[summary]

        [extended_summary]

        Returns:
            [type]: [description]
        """

        for attr in [Meta3DInfo.re_name,Meta3DInfo.re_erase_box,Meta3DInfo.re_erase_pids,Meta3DInfo.re_hes,
                     Meta3DInfo.re_hes_exceptions,Meta3DInfo.re_transparent_pids,Meta3DInfo.re_view]:
            variables = utils.MetaGetVariablesByName("*_{}".format(attr))
            for one_var in variables :
                var_name = one_var[0]
                var_value = one_var[1]
                critical_section_name = var_name.replace("_{}".format(attr),"")
                if critical_section_name in self.critical_sections.keys():
                    self.critical_sections[critical_section_name][attr] = var_value
                else:
                    self.critical_sections[critical_section_name]= {attr:var_value}

        return self

    def hide_all(self):
        """
        hide_all [summary]

        [extended_summary]

        Returns:
            [type]: [description]
        """
        utils.MetaCommand("add all")
        utils.MetaCommand("add invert")

        return 0

    def hide_floor(self):
        """
        hide_floor _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """
        utils.MetaCommand('0:era pid 14001/180000100')

        return 0

    def show_all(self):
        """
        show_all [summary]

        [extended_summary]

        Returns:
            [type]: [description]
        """
        utils.MetaCommand("add all")

        return 0

    def get_props(self,re_props,attr = "Name"):
        """
        get_props [summary]

        [extended_summary]

        Args:
            re_props ([type]): [description]

        Returns:
            [type]: [description]
        """

        entities = []
        if attr == "Name":
            if "*" in re_props:
                for prop in models.CollectModelEntities(0, "parts"):
                    if prop.name.startswith(re_props.replace("*","")):
                        entities.append(prop)
            elif "-" in re_props:
                start = re_props.split("-")[0]
                end = re_props.split("-")[1]
                entities = [prop for prop in self.properties if prop.name in range(start,end+1)]

        return entities

    def show_only_props(self,props):
        """
        show_only_props [summary]

        [extended_summary]

        Args:
            props ([type]): [description]

        Returns:
            [type]: [description]
        """
        pids_string = ",".join([str(entity.id) for entity in props])
        utils.MetaCommand("add pid {}".format(pids_string))

        return 0

    # def peak_state_read():
    #     """
    #     peak_state_read [summary]

    #     [extended_summary]

    #     Returns:
    #         [type]: [description]
    #     """
    #     return 0

class GeneralVarInfo:
    """ This module contains input/output variable Information.
    """

    source_template_file_directory_key = "pptx_template_path"
    source_template_file_name_key = "pptx_filename"
    font_size_info_key = "font_info"
    font1x1_size_info_key = "font_info_1x1"
    font2x2_size_info_key = "font_info_2x2"
    font3x3_size_info_key = "font_info_3x3"
    image1x1_size_key = "image1x1_size"
    image1x2_size_key = "image1x2_size"
    image2x1_size_key = "image2x1_size"
    image2x2_size_key = "image2x2_size"
    image3x1_size_key = "image3x1_size"
    image3x3_size_key = "image3x3_size"
    frames_per_second_key = "frames_per_second"
    log_file_key = "pL"
    log_file_directory_key = "pM"
    target_2d_metadb_key = "pT"
    geometry_key = "pG"
    report_directory_key = "pR"
    displacement_file_key = "pD"
    binout_directory_key = "pA"
    threed_metadb_key = "3D"
    cae_window_key = "quality_check"
    survival_space_window_key = "survival_space"
    verification_mode_key = "verification_mode"
    termination_type_key = "trm_type"
    computation_time_key = "compute_time"
    core_count_key = "compute_proc"
    compute_cluster_key = "cluster_node_1"
    peak_time_key = ""
    run_directory_key = "run_directory"
    survival_space_final_time_key="survival-space_final_time"
    survival_space_final_key="survival-space_final"
    front_door_accel_window_key = "frdoor_accel"
    rear_door_accel_window_key = "rrdoor_accel"
    front_abdomen_intrusion_curve_key = "front_abdomen_intrusion_curve"
    front_femur_intrusion_curve_key = "front_femur_intrusion_curve"
    front_pelvis_intrusion_curve_key = "front_pelvis_intrusion_curve"
    front_shoulder_intrusion_curve_key = "front_shoulder_intrusion_curve"
    rear_abdomen_intrusion_curve_key = "rear_abdomen_intrusion_curve"
    rear_femur_intrusion_curve_key = "rear_femur_intrusion_curve"
    rear_pelvis_intrusion_curve_key = "rear_pelvis_intrusion_curve"
    rear_shoulder_intrusion_curve_key = "rear_shoulder_intrusion_curve"
    survival_space_peak_key = "survival-space_peak"
    biw_stiff_ring_deformation_key = "sill_roof_intrusion"
    biw_accel_window_key = "bplr_disp_window"
    peak_time_display_key = "peak_time_display"
    peak_state_key = "peak_state"
    final_state_key = "survival-space_final_state"
    total_mass_key = "total_mass"
    added_mass_key = "added_mass"
    barrier_mass_key = "barrier_mass"
    floor_mass_key = "floor_mass"
    test_mass_key = "test_mass"
    physical_mass_key = "physical_mass"
    MDB_fr_node_key = "MDB_fr_node"
    struck_subframe_nodes_key = "struck_subframe_nodes"
    MDB_rr_node_key = "MDB_rr_node"
    door_skin_intrusion_window_key = "door_skin_intrusion"
    door_panel_intrusion_window_key = "door_panel_intrusion"

    def __init__(self):

        self.source_template_file_directory = None
        self.source_template_file_name = None
        self.font_size_info = None
        self.font1x1_size_info = None
        self.font2x2_size_info = None
        self.font3x3_size_info = None
        self.image1x1_size = None
        self.image1x2_size = None
        self.image2x1_size = None
        self.image2x2_size = None
        self.image3x1_size = None
        self.image3x3_size = None
        self.frames_per_second = None
        self.log_file = None
        self.log_file_directory = None
        self.target_2d_metadb = None
        self.target_3d_metadb = None
        self.report_directory = None
        self.displacement_file = None
        self.binout_directory = None
        self.threed_metadb_file = None
        self.cae_quality_window_name = None
        self.verification_mode = None
        self.termination_type = None
        self.computation_time = None
        self.core_count = None
        self.compute_cluster = None
        self.peak_time = None
        self.run_directory = None
        self.survival_space_window_name = None
        self.survival_space_final_time = None
        self.front_door_accel_window_name = None
        self.rear_door_accel_window_name = None
        self.front_abdomen_intrusion_curve_name = None
        self.front_femur_intrusion_curve_name = None
        self.front_pelvis_intrusion_curve_name = None
        self.front_shoulder_intrusion_curve_name = None
        self.rear_abdomen_intrusion_curve_name = None
        self.rear_femur_intrusion_curve_name = None
        self.rear_pelvis_intrusion_curve_name = None
        self.rear_shoulder_intrusion_curve_name = None
        self.survival_space_final_value = None
        self.biw_stiff_ring_deformation_name = None
        self.biw_accel_window_name = None
        self.survival_space_peak_value = None
        self.peak_time_display_value = None
        self.threed_window_name = "MetaPost"
        self.peak_state_value = None
        self.final_state_value = None
        self.total_mass_value = None
        self.added_mass_value = None
        self.barrier_mass_value = None
        self.floor_mass_value = None
        self.test_mass_value = None
        self.physical_mass_value = None
        self.MDB_fr_node_id = None
        self.struck_subframe_node_ids = None
        self.MDB_rr_node_id = None
        self.door_skin_intrusion_window_name = None
        self.door_panel_intrusion_window_name = None

    def get_log_directory(self):
        """
        get_log_directory _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """
        get_var = lambda a: utils.MetaGetVariable(a)
        #self.log_file = get_var(GeneralVarInfo.log_file_key)
        self.log_file = "2TN_MP_log"
        self.log_file_directory = get_var(GeneralVarInfo.log_file_directory_key)

        return os.path.join(self.log_file_directory,self.log_file)


    def get_info(self):
        """[summary]

        [extended_summary]

        Returns:
            [type]: [description]
        """

        get_var = lambda a: utils.MetaGetVariable(a)
        self.source_template_file_directory = get_var(GeneralVarInfo.source_template_file_directory_key)
        self.source_template_file_name = get_var(GeneralVarInfo.source_template_file_name_key)
        self.font_size_info = get_var(GeneralVarInfo.font_size_info_key)
        self.font1x1_size_info = get_var(GeneralVarInfo.font1x1_size_info_key)
        self.font2x2_size_info = get_var(GeneralVarInfo.font2x2_size_info_key)
        self.font3x3_size_info = get_var(GeneralVarInfo.font3x3_size_info_key)
        self.image1x1_size = get_var(GeneralVarInfo.image1x1_size_key)
        self.image1x2_size = get_var(GeneralVarInfo.image1x2_size_key)
        self.image2x1_size = get_var(GeneralVarInfo.image2x1_size_key)
        self.image2x2_size = get_var(GeneralVarInfo.image2x2_size_key)
        self.image3x1_size = get_var(GeneralVarInfo.image3x1_size_key)
        self.image3x3_size = get_var(GeneralVarInfo.image3x3_size_key)
        self.frames_per_second = get_var(GeneralVarInfo.frames_per_second_key)
        self.target_2d_metadb = get_var(GeneralVarInfo.target_2d_metadb_key)
        self.geometry_file = get_var(GeneralVarInfo.geometry_key)
        self.report_directory = get_var(GeneralVarInfo.report_directory_key)
        self.displacement_file = get_var(GeneralVarInfo.displacement_file_key)
        self.binout_directory = get_var(GeneralVarInfo.binout_directory_key)
        #self.threed_metadb_file = get_var(GeneralVarInfo.threed_metadb_key)
        self.threed_metadb_file = "/cae/data/tmp/fr2/ra067381/3NT/02_SIDE/05_SICE-2p0/CORRELATION-RERUN/2TN_V2_NP0_DWB_4WD_WB_Master_CntrPllrThinningTop_111821_d_eps_vm.metadb"
        self.cae_quality_window_name = get_var(GeneralVarInfo.cae_window_key)
        self.verification_mode = get_var(GeneralVarInfo.verification_mode_key)
        self.termination_type = get_var(GeneralVarInfo.termination_type_key)
        self.computation_time = get_var(GeneralVarInfo.computation_time_key)
        self.core_count = get_var(GeneralVarInfo.core_count_key)
        self.compute_cluster = get_var(GeneralVarInfo.compute_cluster_key)
        self.peak_time = get_var(GeneralVarInfo.peak_time_key)
        self.run_directory = get_var(GeneralVarInfo.run_directory_key)
        self.survival_space_window_name = get_var(GeneralVarInfo.survival_space_window_key)
        self.survival_space_final_time = get_var(GeneralVarInfo.survival_space_final_time_key)
        self.front_door_accel_window_name = get_var(GeneralVarInfo.front_door_accel_window_key)
        self.rear_door_accel_window_name = get_var(GeneralVarInfo.rear_door_accel_window_key)
        self.front_abdomen_intrusion_curve_name = get_var(GeneralVarInfo.front_abdomen_intrusion_curve_key)
        self.front_femur_intrusion_curve_name = get_var(GeneralVarInfo.front_femur_intrusion_curve_key)
        self.front_pelvis_intrusion_curve_name = get_var(GeneralVarInfo.front_pelvis_intrusion_curve_key)
        self.front_shoulder_intrusion_curve_name = get_var(GeneralVarInfo.front_shoulder_intrusion_curve_key)
        self.rear_abdomen_intrusion_curve_name = get_var(GeneralVarInfo.rear_abdomen_intrusion_curve_key)
        self.rear_femur_intrusion_curve_name = get_var(GeneralVarInfo.rear_femur_intrusion_curve_key)
        self.rear_pelvis_intrusion_curve_name = get_var(GeneralVarInfo.rear_pelvis_intrusion_curve_key)
        self.rear_shoulder_intrusion_curve_name = get_var(GeneralVarInfo.rear_shoulder_intrusion_curve_key)
        self.survival_space_final_value = get_var(GeneralVarInfo.survival_space_final_key)
        self.biw_stiff_ring_deformation_name = get_var(GeneralVarInfo.biw_stiff_ring_deformation_key)
        self.biw_accel_window_name = get_var(GeneralVarInfo.biw_accel_window_key)
        self.survival_space_peak_value = get_var(GeneralVarInfo.survival_space_peak_key)
        self.peak_time_display_value = get_var(GeneralVarInfo.peak_time_display_key)
        self.peak_state_value = get_var(GeneralVarInfo.peak_state_key)
        self.final_state_value = get_var(GeneralVarInfo.final_state_key)
        self.total_mass_value = get_var(GeneralVarInfo.total_mass_key)
        self.added_mass_value = get_var(GeneralVarInfo.added_mass_key)
        self.barrier_mass_value = get_var(GeneralVarInfo.barrier_mass_key)
        self.floor_mass_value = get_var(GeneralVarInfo.floor_mass_key)
        self.test_mass_value = get_var(GeneralVarInfo.test_mass_key)
        self.physical_mass_value = get_var(GeneralVarInfo.physical_mass_key)
        self.MDB_fr_node_id = get_var(GeneralVarInfo.MDB_fr_node_key)
        self.struck_subframe_node_ids = get_var(GeneralVarInfo.struck_subframe_nodes_key)
        self.MDB_rr_node_id = get_var(GeneralVarInfo.MDB_rr_node_key)
        self.door_skin_intrusion_window_name = get_var(GeneralVarInfo.door_skin_intrusion_window_key)
        self.door_panel_intrusion_window_name = get_var(GeneralVarInfo.door_panel_intrusion_window_key)

        return self

class VerifyInfo:
    def __init__(self):
        self.verification_status = True

    def verify(self):
        return 0