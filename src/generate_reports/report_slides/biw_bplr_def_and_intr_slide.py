import os

from meta import utils
from meta import plot2d

from src.meta_utilities import capture_image
from src.general_utilities import closest

class BIWBplrDeformationAndIntrusion():

    def __init__(self,
                slide,
                meta_windows,
                general_input,
                metadb_2d_input,
                metadb_3d_input,
                template_file,
                twod_images_report_folder,
                threed_images_report_folder,
                ppt_report_folder) -> None:
        self.shapes = slide.shapes
        self.meta_windows = meta_windows
        self.general_input = general_input
        self.meta2d_input = metadb_2d_input
        self.meta3d_input = metadb_3d_input
        self.template_file = template_file
        self.twod_images_report_folder = twod_images_report_folder
        self.threed_images_report_folder = threed_images_report_folder
        self.ppt_report_folder = ppt_report_folder


    def edit(self):
        """
        edit _summary_

        _extended_summary_
        """
        window_name = self.general_input.survival_space_window_name
        utils.MetaCommand('window maximize "{}"'.format(window_name))
        for shape in self.shapes:
            if shape.name == "Image 1":
                data = self.meta3d_input.critical_sections["f21_upb_inner"]
                prop_names = data["hes"]
                erase_box = data["erase_box"]
                re_props = prop_names.split(",")
                entities = []
                for re_prop in re_props:
                    utils.MetaCommand('window maximize "MetaPost"')
                    utils.MetaCommand('add all')
                    utils.MetaCommand('view default isometric')
                    entities.extend(self.meta3d_input.get_props(re_prop))
                self.meta3d_input.hide_all()
                self.meta3d_input.show_only_props(entities)
                utils.MetaCommand('erase pid box {}'.format(erase_box))
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
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"f21_upb_inner_with_deform".lower()+".png")
                capture_image("MetaPost",shape.width,shape.height,image_path)
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('plane delete DEFAULT_PLANE_YZ')
                utils.MetaCommand('color pid transparency reset act')
            elif shape.name == "Image 2":
                data = self.meta3d_input.critical_sections["f21_upb_inner"]
                prop_names = data["hes"]
                erase_box = data["erase_box"]
                re_props = prop_names.split(",")
                entities = []
                for re_prop in re_props:
                    utils.MetaCommand('window maximize "MetaPost"')
                    utils.MetaCommand('add all')
                    utils.MetaCommand('view default isometric')
                    entities.extend(self.meta3d_input.get_props(re_prop))
                self.meta3d_input.hide_all()
                self.meta3d_input.show_only_props(entities)
                utils.MetaCommand('erase pid box {}'.format(erase_box))
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
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"f21_upb_inner_Without_deform".lower()+".png")
                capture_image("MetaPost",shape.width,shape.height,image_path)
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('plane delete DEFAULT_PLANE_YZ')
                utils.MetaCommand('color pid transparency reset act')
                utils.MetaCommand('grstyle deform on')
            elif shape.name == "Image 3":
                final_time_curve_name = "SS_{}MS".format(round(float(self.general_input.survival_space_final_time)))
                initial_time_curve_name = "SS_0MS"
                plot_id = 0
                page_id=0
                plot = plot2d.Plot(plot_id, window_name, page_id)
                title = plot.get_title()
                org_name = title.get_text()
                org_name = title.get_text()
                title.set_text("{} 0MS AND 150MS".format(org_name))
                plot.activate()
                final_curve = plot2d.CurvesByName(window_name, final_time_curve_name, 0)[0]
                final_curve.show()
                initial_curve = plot2d.CurvesByName(window_name, initial_time_curve_name, 0)[0]
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
                peak_curve = plot2d.CurvesByName(window_name, peak_curve_value[0].name, 0)[0]
                peak_curve.show()
                utils.MetaCommand('xyplot plotoptions title set "{}" 0 "Survival Space[mm]"'.format(window_name))
                utils.MetaCommand('xyplot curve set color "{}" {} "Cyan"'.format(window_name, initial_curve.id))
                utils.MetaCommand('xyplot curve set style "{}" {} 9'.format(window_name, initial_curve.id))
                utils.MetaCommand('xyplot curve set style "{}" {} 5'.format(window_name, peak_curve.id))

                image_path = os.path.join(self.twod_images_report_folder,window_name+"_with_peak_time"+title.get_text().lower()+".png")
                capture_image(window_name,shape.width,shape.height,image_path,plot_id=plot.id)
                picture = self.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                title.set_text(org_name)
                picture.crop_left = 0
                picture.crop_right = 0
        return 0
