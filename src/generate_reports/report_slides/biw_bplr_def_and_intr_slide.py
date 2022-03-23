# PYTHON script
"""
This script is used for all the automation process of Body in white B-pillar deformation ans intrusion slide of thesis report.
"""
import os
import logging
from datetime import datetime

from meta import utils
from meta import plot2d
from meta import windows

from src.meta_utilities import capture_image
from src.meta_utilities import capture_image_and_resize
from src.meta_utilities import visualize_3d_critical_section
from src.general_utilities import closest
from src.general_utilities import clone_shape

class BIWBplrDeformationAndIntrusion():
    """
       This class is used to automate the BIW Bpillar deformation and intrusion slide of thesis report.

        Args:
            slide (object): biw bpillar deformation and intrusion pptx slide object.
            general_input (GeneralInfo): GeneralInfo class object.
            metadb_3d_input (Meta3DInfo): Meta3DInfo class object.
            twod_images_report_folder (str): folder path to save twod data images.
            threed_images_report_folder (str): folder path to save threed data images.
        """
    def __init__(self,
                slide,
                general_input,
                metadb_3d_input,
                twod_images_report_folder,
                threed_images_report_folder) -> None:
        self.shapes = slide.shapes
        self.general_input = general_input
        self.metadb_3d_input = metadb_3d_input
        self.twod_images_report_folder = twod_images_report_folder
        self.threed_images_report_folder = threed_images_report_folder
        self.logger = logging.getLogger("side_crash_logger")
        self.biw_accel_window_name = None
        self.biw_accel_window_obj = None
        self.biw_accel_window_layout = 0

    def setup(self):
        """
        This method is used to setup data for meta windows based on 2d formats

        Returns:
            int: 0 Always for Sucess.1 for Failure.
        """
        try:
            self.biw_accel_window_name = self.general_input.biw_accel_window_name
            self.biw_accel_window_name = self.biw_accel_window_name.replace("\"","")
            utils.MetaCommand('window maximize "{}"'.format(self.biw_accel_window_name))
            self.biw_accel_window_obj = windows.Window(self.biw_accel_window_name, page_id = 0)
            self.biw_accel_window_layout = self.biw_accel_window_obj.get_plot_layout()

        except Exception  as e:
            return 1

        return 0

    def kinematics_curve_format(self, biw_accel_window_name, plot_id,velocity_min_max_values, displacement_min_max_values,title = None):
        """
        This method is used to format kinematics curves.

        Args:
            biw_accel_window_name (str): biw accel window name.
            plot_id (int): id of the plot.
            velocity_min_max_values (list): list of velocity curve y min and max values.
            displacement_min_max_values (list): list of displacement curve y min nad max values.
            title (str, optional): title of the plot. Defaults to None.

        Returns:
            int: 0 Always for Sucess.1 for Failure.
        """
        try:
            # self.logger.info("Intrusion curve fomat : {}".format(curve_name))
            # self.logger.info("")
            # starttime = datetime.now()
            # self.logger.info("Moving the {} curve to a Temporary window for custom formatting of title,yaxis,xaxis options and attibutes".format(curve_name))
            # self.logger.info(".")
            # self.logger.info(".")
            # self.logger.info(".")
            #rounding the velocity,displacement curves y min and max values
            velocity_values = [round(each_velocity_min_max_value) for each_velocity_min_max_value in velocity_min_max_values]
            displacement_values = [round(each_displacement_min_max_value) for each_displacement_min_max_value in displacement_min_max_values]
            #applying custom style and size for plot title,xaxis,yaxis options and attributes
            utils.MetaCommand('xyplot axisoptions yaxis active "BIW - Accel" {} 0'.format(plot_id))
            utils.MetaCommand('xyplot axisoptions yaxis hideaxis "BIW - Accel" {} 0'.format(plot_id))
            utils.MetaCommand('xyplot curve select "{}" vis'.format(biw_accel_window_name))
            utils.MetaCommand('xyplot curve set linewidth "{}" selected 9'.format(biw_accel_window_name))
            if title:
                utils.MetaCommand('xyplot plotoptions title set "{}" {} "{}"'.format(biw_accel_window_name, plot_id, title))
            utils.MetaCommand('xyplot axisoptions xaxis active "{}" {} 0'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions xlabel font "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions labels xfont "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions xaxis deactive "{}" {} 0'.format(biw_accel_window_name, plot_id))

            utils.MetaCommand('xyplot axisoptions yaxis active "{}" {} 1'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions ylabel font "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions axyrange "{}" {} 1 {} {}'.format(biw_accel_window_name, plot_id,str(round((min(velocity_values)-1000)//1000)*1000), str(round(((max(velocity_values)+1000)//1000)*1000))))
            utils.MetaCommand('xyplot gridoptions yspace "{}" {} 500'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions labels yfont "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions yaxis deactive "{}" {} 1'.format(biw_accel_window_name, plot_id))

            utils.MetaCommand('xyplot axisoptions yaxis active "{}" {} 2'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions ylabel font "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions axyrange "{}" {} 2 {} {}'.format(biw_accel_window_name, plot_id,str(min(displacement_values)-100), str(max(displacement_values)+100)))
            utils.MetaCommand('xyplot gridoptions yspace "{}" {} 50'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions labels yfont "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
            utils.MetaCommand('xyplot axisoptions yaxis deactive "{}" {} 2'.format(biw_accel_window_name, plot_id))

            utils.MetaCommand('xyplot plotoptions title font "{}" {} "Arial,20,-1,5,75,0,0,0,0,0"'.format(biw_accel_window_name, plot_id))
            #endtime = datetime.now()
        except Exception as e:
            self.logger.exception("Error while seeding data into executive report slide:\n{}".format(e))
            self.logger.info("")
            return 1
        # self.logger.info("Intrusion curve format complete")
        # self.logger.info("Time Taken : {}".format(endtime - starttime))
        # self.logger.info("")

        return 0

    def edit(self):
        """
        This method is used to iterate all the shapes of the biw bpillar deformation and intrusion slide and insert respective data.

        Returns:
            int: 0 Always for Sucess,1 for Failure.
        """
        try:
            self.logger.info("Started seeding data into biw bpillar deformation and intrusion slide")
            self.logger.info("")
            starttime = datetime.now()
            oval_shapes = [shape for shape in self.shapes if "Oval" in shape.name]
            survival_space_window_name = self.general_input.survival_space_window_name
            utils.MetaCommand('window maximize "{}"'.format(survival_space_window_name))
            #iterating through the shapes of the biw bpillar deformation and intrusion slide
            for shape in self.shapes:
                #image insertion for the shape named "Image 1"
                if shape.name == "Image 1":
                    #visualizing "f21_upb_inner" critical part set with deformation on
                    data = self.metadb_3d_input.critical_sections["f21_upb_inner"]
                    visualize_3d_critical_section(data)
                    #utils.MetaCommand('color pid transparency reset act')
                    utils.MetaCommand('grstyle scalarfringe enable')
                    #adding a yz plane and slicing the critiical part set with width 500
                    utils.MetaCommand('plane new DEFAULT_PLANE_YZ xyz 1657.996826,-16.504395,576.072754 1,0,0')
                    utils.MetaCommand('plane edit perpendicular 0/1/0 DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane toggleopts enable sectionclip DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane toggleopts disable stateauto DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane options cut autovisible DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane options onlysection enable DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane toggleopts enable clip DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane toggleopts disable sectionclip DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane toggleopts enable slice DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane options slicewidth 500.000000 DEFAULT_PLANE_YZ')
                    utils.MetaCommand('grstyle scalarfringe enable')
                    utils.MetaCommand('view default front')
                    utils.MetaCommand('0:options state variable "serial=1"')
                    utils.MetaCommand('options fringebar off')
                    #capturing "f21_upb_inner" image at peak state
                    image_path = os.path.join(self.threed_images_report_folder,"{}_{}.png".format(self.general_input.threed_window_name,"F21_UPBPILLAR_AT_PEAK_STATE_WITH_DEFORMATION")).replace(" ","_")
                    capture_image(image_path,self.general_input.threed_window_name,shape.width,shape.height)
                    self.logger.info("--- 3D MODEL IMAGE GENERATOR")
                    self.logger.info("")
                    self.logger.info("SOURCE WINDOW : {} ".format(self.general_input.threed_window_name))
                    self.logger.info("SOURCE MODEL : 0")
                    self.logger.info("STATE : PEAK STATE")
                    self.logger.info("PID NAME SHOW FILTER : {} ".format(data["hes"] if "hes" in data.keys() else "null"))
                    self.logger.info("ADDITIONAL PID'S SHOWN : {} ".format(data["hes_exceptions"] if "hes_exceptions" in data.keys() else "null"))
                    self.logger.info("PID NAME ERASE FILTER : {} ".format(data["hes_exceptions"] if "hes_exceptions" in data.keys() else "null"))
                    self.logger.info("PID'S TO ERASE : {} ".format(data["erase_pids"] if "erase_pids" in data.keys() else "null"))
                    self.logger.info("ERASE BOX : {} ".format(data["erase_box"] if "erase_box" in data.keys() else "null"))
                    self.logger.info("IMAGE VIEW : {} ".format(data["view"] if "view" in data.keys() else "null"))
                    self.logger.info("TRANSPARENCY LEVEL : 50" )
                    self.logger.info("TRANSPARENT PID'S : {} ".format(data["transparent_pids"] if "transparent_pids" in data.keys() else "null"))
                    self.logger.info("PLANE CUT & SLICE WIDTH: DEFAULT_PLANE_YZ & 500" )
                    self.logger.info("COMP NAME : {} ".format(data["name"] if "name" in data.keys() else "null"))
                    self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                    self.logger.info("OUTPUT MODEL IMAGES :")
                    self.logger.info(image_path)
                    self.logger.info("")
                    #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                    picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                    picture.crop_left = 0
                    picture.crop_right = 0
                    #deleting created plane and reverting transarency on critical part set visualization
                    utils.MetaCommand('plane delete DEFAULT_PLANE_YZ')
                    utils.MetaCommand('color pid transparency reset act')
                #image insertion for the shape named "Image 2"
                elif shape.name == "Image 2":
                    #visualizing "f21_upb_inner" critical part set with deformation off
                    data = self.metadb_3d_input.critical_sections["f21_upb_inner"]
                    visualize_3d_critical_section(data)
                    #utils.MetaCommand('color pid transparency reset act')
                    utils.MetaCommand('grstyle scalarfringe enable')
                    #adding a yz plane and slicing the critiical part set with width 500
                    utils.MetaCommand('plane new DEFAULT_PLANE_YZ xyz 1657.996826,-16.504395,576.072754 1,0,0')
                    utils.MetaCommand('plane edit perpendicular 0/1/0 DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane toggleopts enable sectionclip DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane toggleopts disable stateauto DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane options cut autovisible DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane options onlysection enable DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane toggleopts enable clip DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane toggleopts disable sectionclip DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane toggleopts enable slice DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane options slicewidth 500.000000 DEFAULT_PLANE_YZ')
                    utils.MetaCommand('grstyle scalarfringe enable')
                    utils.MetaCommand('view default front')
                    utils.MetaCommand('0:options state variable "serial=1"')
                    utils.MetaCommand('options fringebar off')
                    utils.MetaCommand('grstyle deform off')
                    #capturing "f21_upb_inner" image at peak state
                    image_path = os.path.join(self.threed_images_report_folder,"{}_{}.png".format(self.general_input.threed_window_name,"F21_UPBPILLAR_AT_PEAK_STATE_WITHOUT_DEFORMATION")).replace(" ","_")
                    capture_image(image_path,self.general_input.threed_window_name,shape.width,shape.height)
                    self.logger.info("--- 3D MODEL IMAGE GENERATOR")
                    self.logger.info("")
                    self.logger.info("SOURCE WINDOW : {} ".format(self.general_input.threed_window_name))
                    self.logger.info("SOURCE MODEL : 0")
                    self.logger.info("STATE : PEAK STATE")
                    self.logger.info("PID NAME SHOW FILTER : {} ".format(data["hes"] if "hes" in data.keys() else "null"))
                    self.logger.info("ADDITIONAL PID'S SHOWN : {} ".format(data["hes_exceptions"] if "hes_exceptions" in data.keys() else "null"))
                    self.logger.info("PID NAME ERASE FILTER : {} ".format(data["hes_exceptions"] if "hes_exceptions" in data.keys() else "null"))
                    self.logger.info("PID'S TO ERASE : {} ".format(data["erase_pids"] if "erase_pids" in data.keys() else "null"))
                    self.logger.info("ERASE BOX : {} ".format(data["erase_box"] if "erase_box" in data.keys() else "null"))
                    self.logger.info("IMAGE VIEW : {} ".format(data["view"] if "view" in data.keys() else "null"))
                    self.logger.info("TRANSPARENCY LEVEL : 50" )
                    self.logger.info("TRANSPARENT PID'S : {} ".format(data["transparent_pids"] if "transparent_pids" in data.keys() else "null"))
                    self.logger.info("PLANE CUT & SLICE WIDTH: DEFAULT_PLANE_YZ & 500" )
                    self.logger.info("COMP NAME : {} ".format(data["name"] if "name" in data.keys() else "null"))
                    self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                    self.logger.info("OUTPUT MODEL IMAGES :")
                    self.logger.info(image_path)
                    self.logger.info("")
                    #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                    picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                    picture.crop_left = 0
                    picture.crop_right = 0
                    #deleting created plane and reverting transarency on critical part set visualization
                    utils.MetaCommand('plane delete DEFAULT_PLANE_YZ')
                    utils.MetaCommand('color pid transparency reset act')
                    utils.MetaCommand('grstyle deform on')
                #image insertion for the shape named "Image 3"
                elif shape.name == "Image 3":
                    #getting "Survival Space" plot object to activate and showing initial,final and peak time curves
                    final_time_curve_name = "SS_{}MS".format(round(float(self.general_input.survival_space_final_time)))
                    initial_time_curve_name = "SS_0MS"
                    plot_id = 0
                    page_id=0
                    plot = plot2d.Plot(plot_id, survival_space_window_name, page_id)
                    title = plot.get_title()
                    org_name = title.get_text()
                    plot.activate()
                    final_curve = plot2d.CurvesByName(survival_space_window_name, final_time_curve_name, 0)[0]
                    final_curve.show()
                    initial_curve = plot2d.CurvesByName(survival_space_window_name, initial_time_curve_name, 0)[0]
                    initial_curve.show()
                    peak_time_value = self.general_input.peak_time_display_value
                    peak_time_value = peak_time_value.split(".")[0]
                    survival_space_curves = plot.get_curves('all')
                    survival_space_cuves_list = list()
                    for each_survival_space_curves in survival_space_curves:
                        ms = each_survival_space_curves.name.split("_")[1]
                        if 'MS' in ms:
                            ms_replacing = ms.replace('MS',"")
                            survival_space_cuves_list.append(int(ms_replacing))
                    peak_curve_value = closest(survival_space_cuves_list, int(peak_time_value))
                    peak_curve_value = plot.get_curves('byname', name ="SS_"+str(peak_curve_value)+"MS")
                    peak_curve = plot2d.CurvesByName(survival_space_window_name, peak_curve_value[0].name, 0)[0]
                    peak_curve.show()
                    #custom formatting for the "Survival Space" plot title,yaxis,xaxis options
                    utils.MetaCommand('xyplot axisoptions axyrange "{}" 0 0 0 1200'.format(survival_space_window_name))
                    utils.MetaCommand('xyplot axisoptions xaxis active "{}" {} 0'.format(survival_space_window_name, plot_id))
                    utils.MetaCommand('xyplot axisoptions xlabel font "{}" {} "Arial,10,-1,5,75,0,0,0,0,0"'.format(survival_space_window_name, plot_id))
                    utils.MetaCommand('xyplot axisoptions labels xfont "{}" {} "Arial,10,-1,5,75,0,0,0,0,0"'.format(survival_space_window_name, plot_id))
                    utils.MetaCommand('xyplot axisoptions xaxis deactive "{}" {} 0'.format(survival_space_window_name, plot_id))
                    utils.MetaCommand('xyplot axisoptions yaxis active "{}" {} 0'.format(survival_space_window_name, plot_id))
                    utils.MetaCommand('xyplot axisoptions yauto on "{}" {}'.format(survival_space_window_name, plot_id))
                    utils.MetaCommand('xyplot axisoptions ylabel font "{}" {} "Arial,10,-1,5,75,0,0,0,0,0"'.format(survival_space_window_name, plot_id))
                    utils.MetaCommand('xyplot axisoptions labels yfont "{}" {} "Arial,10,-1,5,75,0,0,0,0,0"'.format(survival_space_window_name, plot_id))
                    utils.MetaCommand('xyplot axisoptions yaxis deactive "{}" {} 0'.format(survival_space_window_name, plot_id))
                    utils.MetaCommand('xyplot plotoptions title font "{}" {} "Arial,10,-1,5,75,0,0,0,0,0"'.format(survival_space_window_name, plot_id))
                    utils.MetaCommand('xyplot axisoptions axxrange "{}" 0 0 100 400'.format(survival_space_window_name))
                    utils.MetaCommand('xyplot gridoptions line major style "{}" 0 2'.format(survival_space_window_name))
                    utils.MetaCommand('xyplot plotoptions title set "{}" 0 "{}"'.format(survival_space_window_name,survival_space_window_name))
                    utils.MetaCommand('xyplot curve set color "{}" {} "Cyan"'.format(survival_space_window_name, initial_curve.id))
                    utils.MetaCommand('xyplot curve set style "{}" {} 9'.format(survival_space_window_name, initial_curve.id))
                    utils.MetaCommand('xyplot curve set style "{}" {} 5'.format(survival_space_window_name, peak_curve.id))
                    #capturing "Survival Space" plot image
                    image_path = os.path.join(self.twod_images_report_folder,survival_space_window_name+title.get_text()+".png").replace(" ","_")
                    capture_image(image_path,survival_space_window_name,shape.width,shape.height)
                    #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                    picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                    title.set_text(org_name)
                    picture.crop_left = 0
                    picture.crop_right = 0
                #image insertion for the shape named "Image 5"
                elif shape.name == "Image 4":
                    #calling default 2d data setup and getting the X velocity and X displacement curve from biw accel window of MDB plot id 3
                    self.setup()
                    plot_id = 2
                    page_id = 0
                    velocity_min_max_values = []
                    displacement_min_max_values = []
                    plot = plot2d.Plot(plot_id, self.biw_accel_window_name, page_id)
                    biw_accel_bpillar_upr_l_y_velocity_curve = plot.get_curves('byname', name = '*B-PLR_L_UPR*Y velocity')[0]
                    #collecting min max y values for setting the y axis range of the plot
                    biw_accel_max_velocity = biw_accel_bpillar_upr_l_y_velocity_curve.get_limit_value_y(specifier = 'max')
                    velocity_min_max_values.append(biw_accel_max_velocity)
                    biw_accel_min_velocity = biw_accel_bpillar_upr_l_y_velocity_curve.get_limit_value_y(specifier = 'min')
                    velocity_min_max_values.append(biw_accel_min_velocity)
                    biw_accel_bpillar_upr_l_y_velocity_curve.show()
                    biw_accel_bpillar_upr_l_y_displacement_curve = plot.get_curves('byname', name = '*B-PLR_L_UPR*Follow node Y displacement*')[0]
                    biw_accel_max_curve_displacement = biw_accel_bpillar_upr_l_y_displacement_curve.get_limit_value_y(specifier = 'max')
                    displacement_min_max_values.append(biw_accel_max_curve_displacement)
                    biw_accel_min_displacement = biw_accel_bpillar_upr_l_y_displacement_curve.get_limit_value_y(specifier = 'min')
                    displacement_min_max_values.append(biw_accel_min_displacement)
                    biw_accel_bpillar_upr_l_y_displacement_curve.show()
                    #actiavting plot,getting plot title object and formatting the plot title,yaxis,xaxis options and attributes
                    title = plot2d.Title(plot_id, self.biw_accel_window_name, page_id)
                    plot.activate()
                    self.activated_plot = plot
                    utils.MetaCommand('xyplot rlayout "{}" {}'.format(self.biw_accel_window_name, 1))
                    self.kinematics_curve_format(self.biw_accel_window_name, plot_id, velocity_min_max_values, displacement_min_max_values)
                    #capturing "MDB" plot image
                    image_path = os.path.join(self.twod_images_report_folder,self.biw_accel_window_name+"_"+title.get_text()+".png").replace(" ","_")
                    capture_image_and_resize(image_path,shape.width,shape.height)
                    self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                    self.logger.info("")
                    self.logger.info("CURVES : {},{} | SOURCE PLOT : {} | SOURCE WINDOW : {}".format("*B-PLR_L_UPR*Y velocity","*B-PLR_L_UPR*Follow node Y displacement*",title.get_text(),self.biw_accel_window_name))
                    self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                    self.logger.info("OUTPUT CURVE IMAGES : ")
                    self.logger.info(image_path)
                    self.logger.info("")
                    #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                    picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                    picture.crop_left = 0
                    picture.crop_right = 0
                    #reverting 2d data setup
                    biw_accel_bpillar_upr_l_y_displacement_curve.hide()
                    biw_accel_bpillar_upr_l_y_velocity_curve.hide()
                    self.revert()
                #image insertion for the shape named "Image 6"
                elif shape.name == "Image 5":
                    #calling default 2d data setup and getting the Y velocity and Y displacement curve from biw accel window of UNIT plot id 0
                    self.setup()
                    plot_id = 2
                    page_id = 0
                    plot = plot2d.Plot(plot_id, self.biw_accel_window_name, page_id)
                    velocity_min_max_values = []
                    displacement_min_max_values = []
                    biw_accel_bpillar_mid_upr_l_y_velocity_curve = plot.get_curves('byname', name = '*B-PLR_L_MID_UPR*Y velocity')[0]
                    #collecting min max y values for setting the y axis range of the plot
                    biw_accel_max_velocity = biw_accel_bpillar_mid_upr_l_y_velocity_curve.get_limit_value_y(specifier = 'max')
                    velocity_min_max_values.append(biw_accel_max_velocity)
                    biw_accel_min_velocity = biw_accel_bpillar_mid_upr_l_y_velocity_curve.get_limit_value_y(specifier = 'min')
                    velocity_min_max_values.append(biw_accel_min_velocity)
                    biw_accel_bpillar_mid_upr_l_y_velocity_curve.show()
                    biw_accel_bpillar_mid_upr_l_y_displacement_curve = plot.get_curves('byname', name = '*B-PLR_L_MID_UPR*Follow node Y displacement*')[0]
                    biw_accel_max_curve_displacement = biw_accel_bpillar_mid_upr_l_y_displacement_curve.get_limit_value_y(specifier = 'max')
                    displacement_min_max_values.append(biw_accel_max_curve_displacement)
                    biw_accel_min_displacement = biw_accel_bpillar_mid_upr_l_y_displacement_curve.get_limit_value_y(specifier = 'min')
                    displacement_min_max_values.append(biw_accel_min_displacement)
                    biw_accel_bpillar_mid_upr_l_y_displacement_curve.show()
                    #actiavting plot,getting plot title object and formatting the plot title,yaxis,xaxis options and attributes
                    title = plot2d.Title(plot_id, self.biw_accel_window_name, page_id)
                    plot.activate()
                    self.activated_plot = plot
                    utils.MetaCommand('xyplot rlayout "{}" {}'.format(self.biw_accel_window_name, 1))
                    self.kinematics_curve_format(self.biw_accel_window_name, plot_id,velocity_min_max_values, displacement_min_max_values, title = "UNIT")
                    #capturing "UNIT" plot image
                    image_path = os.path.join(self.twod_images_report_folder,self.biw_accel_window_name+"_"+title.get_text()+".png").replace(" ","_")
                    capture_image_and_resize(image_path,shape.width,shape.height)
                    self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                    self.logger.info("")
                    self.logger.info("CURVES : {},{} | SOURCE PLOT : {} | SOURCE WINDOW : {}".format("*B-PLR_L_MID_UPR*Y velocity","*B-PLR_L_MID_UPR*Follow node Y displacement*",title.get_text(),self.biw_accel_window_name))
                    self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                    self.logger.info("OUTPUT CURVE IMAGES : ")
                    self.logger.info(image_path)
                    self.logger.info("")
                    #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                    picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                    picture.crop_left = 0
                    picture.crop_right = 0
                    #reverting 2d data setup
                    biw_accel_bpillar_mid_upr_l_y_displacement_curve.hide()
                    biw_accel_bpillar_mid_upr_l_y_velocity_curve.hide()
                    self.revert()
                #image insertion for the shape named "Image 7"
                elif shape.name == "Image 6":
                    #calling default 2d data setup and getting the Y velocity and Y displacement curve from biw accel window of APLR R plot id 1
                    self.setup()
                    plot_id = 2
                    page_id = 0
                    plot = plot2d.Plot(plot_id, self.biw_accel_window_name, page_id)
                    velocity_min_max_values = []
                    displacement_min_max_values = []
                    biw_accel_bpillar_mid_l_y_velocity_curve = plot.get_curves('byname', name = '*B-PLR_L_MID*Y velocity')[0]
                    #collecting min max y values for setting the y axis range of the plot
                    biw_accel_max_velocity = biw_accel_bpillar_mid_l_y_velocity_curve.get_limit_value_y(specifier = 'max')
                    velocity_min_max_values.append(biw_accel_max_velocity)
                    biw_accel_min_velocity = biw_accel_bpillar_mid_l_y_velocity_curve.get_limit_value_y(specifier = 'min')
                    velocity_min_max_values.append(biw_accel_min_velocity)
                    biw_accel_bpillar_mid_l_y_velocity_curve.show()
                    biw_accel_bpillar_mid_l_y_displacement_curve = plot.get_curves('byname', name = '*B-PLR_L_MID*Follow node Y displacement*')[0]
                    biw_accel_max_curve_displacement = biw_accel_bpillar_mid_l_y_displacement_curve.get_limit_value_y(specifier = 'max')
                    displacement_min_max_values.append(biw_accel_max_curve_displacement)
                    biw_accel_min_displacement = biw_accel_bpillar_mid_l_y_displacement_curve.get_limit_value_y(specifier = 'min')
                    displacement_min_max_values.append(biw_accel_min_displacement)
                    biw_accel_bpillar_mid_l_y_displacement_curve.show()
                    #actiavting plot,getting plot title object and formatting the plot title,yaxis,xaxis options and attributes
                    title = plot2d.Title(plot_id, self.biw_accel_window_name, page_id)
                    plot.activate()
                    self.activated_plot = plot
                    utils.MetaCommand('xyplot rlayout "{}" {}'.format(self.biw_accel_window_name, 1))
                    self.kinematics_curve_format(self.biw_accel_window_name, plot_id, velocity_min_max_values, displacement_min_max_values,title="APLR_R")
                    #capturing "APLR R" plot image
                    image_path = os.path.join(self.twod_images_report_folder,self.biw_accel_window_name+"_"+title.get_text()+".png").replace(" ","_")
                    capture_image_and_resize(image_path,shape.width,shape.height)
                    self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                    self.logger.info("")
                    self.logger.info("CURVES : {},{} | SOURCE PLOT : {} | SOURCE WINDOW : {}".format("*B-PLR_L_MID*Y velocity","*B-PLR_L_MID*Follow node Y displacement*",title.get_text(),self.biw_accel_window_name))
                    self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                    self.logger.info("OUTPUT CURVE IMAGES : ")
                    self.logger.info(image_path)
                    self.logger.info("")
                    #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                    picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                    picture.crop_left = 0
                    picture.crop_right = 0
                    #reverting 2d data setup
                    biw_accel_bpillar_mid_l_y_displacement_curve.hide()
                    biw_accel_bpillar_mid_l_y_velocity_curve.hide()
                    self.revert()
                #image insertion for the shape named "Image 8"
                elif shape.name == "Image 7":
                    #calling default 2d data setup and getting the Y velocity and Y displacement curve from biw accel window of SIS ROW2 R plot id 1
                    self.setup()
                    plot_id = 2
                    page_id = 0
                    plot = plot2d.Plot(plot_id, self.biw_accel_window_name, page_id)
                    velocity_min_max_values = []
                    displacement_min_max_values = []
                    biw_accel_bpillar_mid_lwr_l_y_velocity_curve = plot.get_curves('byname', name = '*B-PLR_L_MID_LWR*Y velocity')[0]
                    #collecting min max y values for setting the y axis range of the plot
                    biw_accel_max_velocity = biw_accel_bpillar_mid_lwr_l_y_velocity_curve.get_limit_value_y(specifier = 'max')
                    velocity_min_max_values.append(biw_accel_max_velocity)
                    biw_accel_min_velocity = biw_accel_bpillar_mid_lwr_l_y_velocity_curve.get_limit_value_y(specifier = 'min')
                    velocity_min_max_values.append(biw_accel_min_velocity)
                    biw_accel_bpillar_mid_lwr_l_y_velocity_curve.show()
                    biw_accel_bpillar_mid_lwr_l_y_displacement_curve = plot.get_curves('byname', name = '*B-PLR_L_MID_LWR*Follow node Y displacement*')[0]
                    biw_accel_max_curve_displacement = biw_accel_bpillar_mid_lwr_l_y_displacement_curve.get_limit_value_y(specifier = 'max')
                    displacement_min_max_values.append(biw_accel_max_curve_displacement)
                    biw_accel_min_displacement = biw_accel_bpillar_mid_lwr_l_y_displacement_curve.get_limit_value_y(specifier = 'min')
                    displacement_min_max_values.append(biw_accel_min_displacement)
                    biw_accel_bpillar_mid_lwr_l_y_displacement_curve.show()
                    title = plot2d.Title(plot_id, self.biw_accel_window_name, page_id)
                    #actiavting plot,getting plot title object and formatting the plot title,yaxis,xaxis options and attributes
                    plot.activate()
                    self.activated_plot = plot
                    utils.MetaCommand('xyplot rlayout "{}" {}'.format(self.biw_accel_window_name, 1))
                    self.kinematics_curve_format(self.biw_accel_window_name, plot_id,velocity_min_max_values, displacement_min_max_values,title="SS_BP_R")
                    #capturing "SIS ROW2 R" plot image
                    image_path = os.path.join(self.twod_images_report_folder,self.biw_accel_window_name+"_"+title.get_text()+".png").replace(" ","_")
                    capture_image_and_resize(image_path,shape.width,shape.height)
                    self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                    self.logger.info("")
                    self.logger.info("CURVES : {},{} | SOURCE PLOT : {} | SOURCE WINDOW : {}".format("*B-PLR_L_MID_LWR*Y velocity","*B-PLR_L_MID_LWR*Follow node Y displacement*",title.get_text(),self.biw_accel_window_name))
                    self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                    self.logger.info("OUTPUT CURVE IMAGES : ")
                    self.logger.info(image_path)
                    self.logger.info("")
                    #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                    picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                    picture.crop_left = 0
                    picture.crop_right = 0
                    #reverting 2d data setup
                    biw_accel_bpillar_mid_lwr_l_y_displacement_curve.hide()
                    biw_accel_bpillar_mid_lwr_l_y_velocity_curve.hide()
                    self.revert()
                #image insertion for the shape named "Image 9"
                elif shape.name == "Image 8":
                    #calling default 2d data setup and getting the Y velocity and Y displacement curve from biw accel window of SS_R_RR_TOP_G_Y1 plot id 1
                    self.setup()
                    plot_id = 2
                    page_id = 0
                    plot = plot2d.Plot(plot_id, self.biw_accel_window_name, page_id)
                    velocity_min_max_values = []
                    displacement_min_max_values = []
                    biw_accel_bpillar_lwr_l_y_velocity_curve = plot.get_curves('byname', name = '*B-PLR_L_LWR*Y velocity')[0]
                    #collecting min max y values for setting the y axis range of the plot
                    biw_accel_max_velocity = biw_accel_bpillar_lwr_l_y_velocity_curve.get_limit_value_y(specifier = 'max')
                    velocity_min_max_values.append(biw_accel_max_velocity)
                    biw_accel_min_velocity = biw_accel_bpillar_lwr_l_y_velocity_curve.get_limit_value_y(specifier = 'min')
                    velocity_min_max_values.append(biw_accel_min_velocity)
                    biw_accel_bpillar_lwr_l_y_velocity_curve.show()
                    biw_accel_bpillar_lwr_l_y_displacement_curve = plot.get_curves('byname', name = '*B-PLR_L_LWR*Follow node Y displacement*')[0]
                    biw_accel_max_curve_displacement = biw_accel_bpillar_lwr_l_y_displacement_curve.get_limit_value_y(specifier = 'max')
                    displacement_min_max_values.append(biw_accel_max_curve_displacement)
                    biw_accel_min_displacement = biw_accel_bpillar_lwr_l_y_displacement_curve.get_limit_value_y(specifier = 'min')
                    displacement_min_max_values.append(biw_accel_min_displacement)
                    biw_accel_bpillar_lwr_l_y_displacement_curve.show()
                    title = plot2d.Title(plot_id, self.biw_accel_window_name, page_id)
                    #actiavting plot,getting plot title object and formatting the plot title,yaxis,xaxis options and attributes
                    plot.activate()
                    self.activated_plot = plot
                    utils.MetaCommand('xyplot rlayout "{}" {}'.format(self.biw_accel_window_name, 1))
                    self.kinematics_curve_format(self.biw_accel_window_name, plot_id,velocity_min_max_values, displacement_min_max_values, title="SS_R_RR_TOP")
                    #capturing "SS_R_RR_TOP_G_Y1" plot image
                    image_path = os.path.join(self.twod_images_report_folder,self.biw_accel_window_name+"_"+title.get_text()+".png").replace(" ","_")
                    capture_image_and_resize(image_path,shape.width,shape.height)
                    self.logger.info("--- 2D CURVE IMAGE GENERATOR")
                    self.logger.info("")
                    self.logger.info("CURVES : {},{} | SOURCE PLOT : {} | SOURCE WINDOW : {}".format("*B-PLR_L_LWR*Y velocity","*B-PLR_L_LWR*Follow node Y displacement*",title.get_text(),self.biw_accel_window_name))
                    self.logger.info("OUTPUT IMAGE SIZE (PIXELS) : {}x{}".format(round(shape.width/9525),round(shape.height/9525)))
                    self.logger.info("OUTPUT CURVE IMAGES : ")
                    self.logger.info(image_path)
                    self.logger.info("")
                    #adding picture based on the shape width and height, which will hide the original shape and add a picture shape on top of that
                    picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                    picture.crop_left = 0
                    picture.crop_right = 0
                    #reverting 2d data setup
                    biw_accel_bpillar_lwr_l_y_displacement_curve.hide()
                    biw_accel_bpillar_lwr_l_y_velocity_curve.hide()
                    self.revert()
            # duplicating oval shapes
            for shape in oval_shapes:
                clone_shape(shape)
            endtime = datetime.now()
        except Exception as e:
            self.logger.exception("Error while seeding data into biw bpillar deformation and intrusion slide:\n{}".format(e))
            self.logger.info("")
            return 1
        self.logger.info("Completed seeding data into biw bpillar deformation and intrusion slide")
        self.logger.info("Time Taken : {}".format(endtime - starttime))
        self.logger.info("")

        return 0


    def revert(self):
        """
        This method is used to revert data for meta windows based on 2d formats

        Returns:
            int: 0 Always for Sucess.1 for Failure.
        """
        try:
            self.activated_plot.deactivate()
            utils.MetaCommand('xyplot rlayout "{}" {}'.format(self.biw_accel_window_name, self.biw_accel_window_layout))

        except Exception as e:
            return 1

        return 0