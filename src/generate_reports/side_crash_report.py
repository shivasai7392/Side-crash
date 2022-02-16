import os
import time


from meta import utils
from meta import windows
from meta import plot2d

class SideCrashReport():

    def __init__(self,windows,general_input,metadb_2d_input,metadb_3d_input,config_folder) -> None:
        self.windows = windows
        self.general_input = general_input
        self.metadb_2d_input = metadb_2d_input
        self.metadb_3d_input = metadb_3d_input
        self.config_folder = config_folder
        self.template_file = os.path.join(self.config_folder,"res",self.general_input.source_template_file_directory.replace("/","",1),self.general_input.source_template_file_name).replace("\\",os.sep)
        self.get_reporting_folders()

    def get_reporting_folders(self):
        self.twod_images_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.report_directory).replace("/","",1),"2d-data-images").replace("\\",os.sep)
        self.threed_images_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.report_directory).replace("/","",1),"3d-data-images").replace("\\",os.sep)
        self.threed_videos_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.report_directory).replace("/","",1),"3d-data-videos").replace("\\",os.sep)
        self.excel_bom_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.report_directory).replace("/","",1),"excel-bom").replace("\\",os.sep)
        self.ppt_report_folder = os.path.join(self.config_folder,"res",os.path.dirname(self.general_input.report_directory).replace("/","",1),"reports").replace("\\",os.sep)

        return 0

    def run_process(self):

        self.twod_data_reporting()
        self.threed_data_reporting()
        self.thesis_report_generation()
        return 0


    def thesis_report_generation(self):
        """
        thesis_report_generation [summary]

        [extended_summary]

        Returns:
            [type]: [description]
        """

        self.report_composer = PPTXReportComposer(report_name="Run1",template_pptx=self.template_file)
        self.report_composer.create_prs_obj()

        utils.MetaCommand('options title off')
        utils.MetaCommand('options axis off')
        self.edit_title_slide(self.report_composer.prs_obj.slides[0])
        self.edit_cae_quality_slide(self.report_composer.prs_obj.slides[1])
        self.edit_executive_slide(self.report_composer.prs_obj.slides[2])
        self.edit_cbu_and_barrier_position_slide(self.report_composer.prs_obj.slides[3])
        self.edit_body_in_white_kinematics_slide(self.report_composer.prs_obj.slides[6])
        self.edit_bill_of_materials_f21_upb(self.report_composer.prs_obj.slides[8])

        if not os.path.exists(self.ppt_report_folder):
            os.makedirs(self.ppt_report_folder)
        file_name = os.path.join(self.ppt_report_folder,"output.pptx")
        self.report_composer.save_pptx(file_name)

        return 0

    def edit_bill_of_materials_f21_upb(self, slide):

        from PIL import ImageGrab
        from pptx.util import Pt

        for shape in slide.shapes:
            if shape.name == "Image 2":
                data = self.metadb_3d_input.critical_sections["f21_upb_inner"]
                prop_names = data["hes"]
                re_props = prop_names.split(",")
                entities = []
                for re_prop in re_props:
                    utils.MetaCommand('window maximize "MetaPost"')
                    entities.extend(self.metadb_3d_input.get_props(re_prop))
                print("entities:", len(entities))
                self.metadb_3d_input.hide_all()
                self.metadb_3d_input.show_only_props(entities)
                utils.MetaCommand('view default right')
                utils.MetaCommand('options fringebar off')
                utils.MetaCommand('clipboard copy image "MetaPost"')
                img = ImageGrab.grabclipboard()
                img = img.resize((round(shape.width/9525),round(shape.height/9525)))
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"f21_upb_inner".lower()+".jpeg")
                if not os.path.exists(os.path.dirname(image_path)):
                    print(os.path.dirname(image_path))
                    os.makedirs(os.path.dirname(image_path))
                img.save(image_path, 'PNG')
                picture = slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
            elif shape.name == "Image 1":
                data = self.metadb_3d_input.critical_sections["f21_upb_outer"]
                prop_names = data["hes"]
                re_props = prop_names.split(",")
                entities = []
                for re_prop in re_props:
                    utils.MetaCommand('window maximize "MetaPost"')
                    entities.extend(self.metadb_3d_input.get_props(re_prop))
                print("image1", len(entities))
                self.metadb_3d_input.hide_all()
                self.metadb_3d_input.show_only_props(entities)
                utils.MetaCommand('view default left')
                utils.MetaCommand('options fringebar off')
                utils.MetaCommand('clipboard copy image "MetaPost"')
                img = ImageGrab.grabclipboard()
                img = img.resize((round(shape.width/9525),round(shape.height/9525)))
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"f21_upb_inner".lower()+".jpeg")
                if not os.path.exists(os.path.dirname(image_path)):
                    print(os.path.dirname(image_path))
                    os.makedirs(os.path.dirname(image_path))
                img.save(image_path, 'PNG')
                picture = slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
            elif shape.name == "Table 1":
                data = self.metadb_3d_input.critical_sections["f21_upb_inner"]
                prop_names = data["hes"]
                re_props = prop_names.split(",")
                entities = []
                for re_prop in re_props:
                    utils.MetaCommand('window maximize "MetaPost"')
                    entities.extend(self.metadb_3d_input.get_props(re_prop))
                table_obj = shape.table
                for id,prop in enumerate(entities[:15]):
                    self.report_composer.add_row(table_obj)
                    prop_row = table_obj.rows[id+1]
                    text_frame = prop_row.cells[0].text_frame
                    font = text_frame.paragraphs[0].font
                    font.size = Pt(8)
                    text_frame.paragraphs[0].text = str(prop.id)

        return 0

    def edit_body_in_white_kinematics_slide(self, slide):
        from PIL import ImageGrab,Image
        utils.MetaCommand('window maximize "MetaPost"')

        for shape in slide.shapes:
            if shape.name == "Image 1":
                utils.MetaCommand('add all')
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('view default front')
                utils.MetaCommand('color fringebar scalarset Critical')
                utils.MetaCommand('clipboard copy image "MetaPost"')
                utils.MetaCommand('color fringebar scalarset default')
                img = ImageGrab.grabclipboard()
                img = img.resize((round(shape.width/9525),round(shape.height/9525)))
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"model_front".lower()+".jpeg")
                if not os.path.exists(os.path.dirname(image_path)):
                    print(os.path.dirname(image_path))
                    os.makedirs(os.path.dirname(image_path))
                img.save(image_path, 'PNG')
                picture = slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0

            elif shape.name == "Image 2":
                utils.MetaCommand('add all')
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('view default top')
                utils.MetaCommand('color fringebar scalarset Critical')
                utils.MetaCommand('clipboard copy image "MetaPost"')
                utils.MetaCommand('color fringebar scalarset default')
                img = ImageGrab.grabclipboard()
                img = img.resize((round(shape.width/9525),round(shape.height/9525)))
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"model_top".lower()+".jpeg")
                if not os.path.exists(os.path.dirname(image_path)):
                    print(os.path.dirname(image_path))
                    os.makedirs(os.path.dirname(image_path))
                img.save(image_path, 'PNG')
                picture = slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0

            elif shape.name == "Image 3":
                utils.MetaCommand('add all')
                utils.MetaCommand('view default isometric')
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('view default front')
                utils.MetaCommand('color fringebar scalarset Critical')
                utils.MetaCommand('clipboard copy image "MetaPost"')
                utils.MetaCommand('color fringebar scalarset default')
                img = ImageGrab.grabclipboard()
                img = img.resize((round(shape.width/9525),round(shape.height/9525)))
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"model_front".lower()+".jpeg")
                if not os.path.exists(os.path.dirname(image_path)):
                    print(os.path.dirname(image_path))
                    os.makedirs(os.path.dirname(image_path))
                img.save(image_path, 'PNG')
                picture = slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0

            elif shape.name == "Image 4":
                utils.MetaCommand('add all')
                utils.MetaCommand('view default isometric')
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('view default top')
                utils.MetaCommand('color fringebar scalarset Critical')
                utils.MetaCommand('clipboard copy image "MetaPost"')
                utils.MetaCommand('color fringebar scalarset default')
                img = ImageGrab.grabclipboard()
                img = img.resize((round(shape.width/9525),round(shape.height/9525)))
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"model_top".lower()+".jpeg")
                if not os.path.exists(os.path.dirname(image_path)):
                    print(os.path.dirname(image_path))
                    os.makedirs(os.path.dirname(image_path))
                img.save(image_path, 'PNG')
                picture = slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0

    def edit_cbu_and_barrier_position_slide(self, slide):

        from PIL import ImageGrab,Image
        utils.MetaCommand('window maximize "MetaPost"')

        for shape in slide.shapes:
            if shape.name == "Image 4":
                #data = self.metadb_3d_input.critical_sections["barrier_and_cbu"]
                data = self.metadb_3d_input.critical_sections["cbu"]
                prop_names = data["hes"]
                re_props = prop_names.split(",")
                entities = []
                for re_prop in re_props:
                    utils.MetaCommand('0:options state original')
                    utils.MetaCommand('add all')
                    utils.MetaCommand('view default isometric')
                    utils.MetaCommand('view default top')
                    entities.extend(self.metadb_3d_input.get_props(re_prop))
                self.metadb_3d_input.hide_all()
                self.metadb_3d_input.show_only_props(entities)
                utils.MetaCommand('clipboard copy image "MetaPost"')
                img = ImageGrab.grabclipboard()
                img = img.resize((round(shape.width/9525),round(shape.height/9525)))
                img = img.transpose(Image.ROTATE_90)
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"cbu".lower()+".jpeg")
                if not os.path.exists(os.path.dirname(image_path)):
                    print(os.path.dirname(image_path))
                    os.makedirs(os.path.dirname(image_path))
                img.save(image_path, 'PNG')
                picture = slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0

            if shape.name == "Image 3":
                #data = self.metadb_3d_input.critical_sections["barrier_and_cbu"]
                data = self.metadb_3d_input.critical_sections["cbu"]
                prop_names = data["hes"]
                re_props = prop_names.split(",")
                entities = []
                for re_prop in re_props:
                    utils.MetaCommand('0:options state variable "serial=1"')
                    utils.MetaCommand('add all')
                    utils.MetaCommand('view default isometric')
                    utils.MetaCommand('grstyle deform off')
                    utils.MetaCommand('color fringebar scalarset Critical')
                    utils.MetaCommand('view default left')
                    entities.extend(self.metadb_3d_input.get_props(re_prop))
                self.metadb_3d_input.hide_all()
                self.metadb_3d_input.show_only_props(entities)
                utils.MetaCommand('clipboard copy image "MetaPost"')
                utils.MetaCommand('grstyle deform on')
                utils.MetaCommand('color fringebar scalarset default')
                img = ImageGrab.grabclipboard()
                img = img.resize((round(shape.width/9525),round(shape.height/9525)))
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"cbu_critical".lower()+".jpeg")
                if not os.path.exists(os.path.dirname(image_path)):
                    print(os.path.dirname(image_path))
                    os.makedirs(os.path.dirname(image_path))
                img.save(image_path, 'PNG')
                picture = slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0

        return 0

    def edit_executive_slide(self,slide):

        from PIL import ImageGrab

        for shape in slide.shapes:
            if shape.name == "Image 3":
                data = self.metadb_3d_input.critical_sections["f28_front_door"]
                prop_names = data["hes"]
                re_props = prop_names.split(",")
                entities = []
                for re_prop in re_props:
                    utils.MetaCommand('window maximize "MetaPost"')
                    utils.MetaCommand('add all')
                    utils.MetaCommand('view default isometric')
                    entities.extend(self.metadb_3d_input.get_props(re_prop))
                self.metadb_3d_input.hide_all()
                self.metadb_3d_input.show_only_props(entities)
                utils.MetaCommand('view default isometric')
                utils.MetaCommand('grstyle scalarfringe enable')
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('options fringebar off')
                utils.MetaCommand('clipboard copy image "MetaPost"')
                img = ImageGrab.grabclipboard()
                img = img.resize((round(shape.width/9525),round(shape.height/9525)))
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"f28_front_door".lower()+".jpeg")
                if not os.path.exists(os.path.dirname(image_path)):
                    print(os.path.dirname(image_path))
                    os.makedirs(os.path.dirname(image_path))
                img.save(image_path, 'PNG')
                picture = slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
            elif shape.name == "Image 4":
                data = self.metadb_3d_input.critical_sections["f28_rear_door"]
                prop_names = data["hes"]
                re_props = prop_names.split(",")
                entities = []
                for re_prop in re_props:
                    utils.MetaCommand('window maximize "MetaPost"')
                    utils.MetaCommand('add all')
                    utils.MetaCommand('view default isometric')
                    entities.extend(self.metadb_3d_input.get_props(re_prop))
                self.metadb_3d_input.hide_all()
                self.metadb_3d_input.show_only_props(entities)
                utils.MetaCommand('view default isometric')
                utils.MetaCommand('grstyle scalarfringe enable')
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('options fringebar off')
                utils.MetaCommand('clipboard copy image "MetaPost"')
                img = ImageGrab.grabclipboard()
                img = img.resize((round(shape.width/9525),round(shape.height/9525)))
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"f28_front_door".lower()+".jpeg")
                if not os.path.exists(os.path.dirname(image_path)):
                    print(os.path.dirname(image_path))
                    os.makedirs(os.path.dirname(image_path))
                img.save(image_path, 'PNG')
                picture = slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
            elif shape.name == "Image 2":
                data = self.metadb_3d_input.critical_sections["f21_upb_inner"]
                prop_names = data["hes"]
                re_props = prop_names.split(",")
                entities = []
                for re_prop in re_props:
                    utils.MetaCommand('window maximize "MetaPost"')
                    utils.MetaCommand('add all')
                    utils.MetaCommand('view default isometric')
                    entities.extend(self.metadb_3d_input.get_props(re_prop))
                self.metadb_3d_input.hide_all()
                self.metadb_3d_input.show_only_props(entities)
                utils.MetaCommand('plane new default yz')
                utils.MetaCommand('plane toggleopts enable sectionclip DEFAULT_PLANE_YZ')
                utils.MetaCommand('plane toggleopts disable stateauto DEFAULT_PLANE_YZ')
                utils.MetaCommand('plane options cut autovisible DEFAULT_PLANE_YZ')
                utils.MetaCommand('plane options onlysection enable DEFAULT_PLANE_YZ')
                utils.MetaCommand('plane toggleopts enable clip DEFAULT_PLANE_YZ')
                utils.MetaCommand('plane toggleopts disable sectionclip DEFAULT_PLANE_YZ')
                utils.MetaCommand('plane toggleopts enable slice DEFAULT_PLANE_YZ')
                utils.MetaCommand('plane options slicewidth 500.000000 DEFAULT_PLANE_YZ')
                utils.MetaCommand('view default front')
                utils.MetaCommand('grstyle scalarfringe enable')
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('options fringebar off')
                utils.MetaCommand('clipboard copy image "MetaPost"')
                img = ImageGrab.grabclipboard()
                img = img.resize((round(shape.width/9525),round(shape.height/9525)))
                image_path = os.path.join(self.threed_images_report_folder,"MetaPost"+"_"+"f21_upb_inner".lower()+".jpeg")
                if not os.path.exists(os.path.dirname(image_path)):
                    print(os.path.dirname(image_path))
                    os.makedirs(os.path.dirname(image_path))
                img.save(image_path, 'PNG')
                picture = slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0

        return 0

    def edit_title_slide(self,slide):
        from datetime import datetime
        from pptx.util import Pt

        for shape in slide.shapes:
            if shape.name == "Table 2":
                table_obj = shape.table
                rows = table_obj.rows
                ordinal = lambda n: "%d%s" % (n,"tsnrhtdd"[(n//10%10!=1)*(n%10<4)*n%10::4])
                format_data = "%B %-d, %Y"
                month = datetime.today().strftime('%B')
                day = int(datetime.today().strftime('%d'))
                year = datetime.today().strftime('%Y')
                text_frame_1 = rows[3].cells[0].text_frame
                font = text_frame_1.paragraphs[0].font
                font.size = Pt(16)
                text_frame_2 = rows[2].cells[0].text_frame
                font = text_frame_2.paragraphs[0].font
                font.size = Pt(16)
                text_frame_3 = rows[4].cells[2].text_frame
                font = text_frame_3.paragraphs[0].font
                font.size = Pt(16)
                text_frame_1.paragraphs[0].text = " {} {}, {}".format(month,ordinal(day),year)
                text_frame_2.paragraphs[0].text = " " + self.general_input.verification_mode
                text_frame_3.paragraphs[0].text = " " + self.general_input.run_directory

        return 0

    def edit_cae_quality_slide(self,slide):

        from PIL import ImageGrab
        from pptx.util import Pt

        window_name = self.general_input.cae_quality_window_name
        utils.MetaCommand('window maximize "{}"'.format(window_name))

        for shape in slide.shapes:

            if shape.name == "Image 2":
                plot_id = 0
                page_id=0
                title = plot2d.Title(plot_id, window_name, page_id)
                plot = title.get_plot()
                plot.activate()
                utils.MetaCommand('xyplot rlayout "{}" 1'.format(window_name))
                utils.MetaCommand('xyplot curve select "{}" all'.format(window_name))
                utils.MetaCommand('xyplot curve visible and "{}" sel'.format(window_name))
                utils.MetaCommand('clipboard copy plot image "{}" {}'.format(window_name, plot.id))
                img = ImageGrab.grabclipboard()
                img = img.resize((round(shape.width/9525),round(shape.height/9525)))
                image_path = os.path.join(self.twod_images_report_folder,window_name+"_"+title.get_text().lower()+".jpeg")
                img.save(image_path, 'PNG')
                picture = slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('xyplot plotoptions legend on "CAE quality" 0')
                utils.MetaCommand('xyplot legend hook left "CAE quality" 0')
                utils.MetaCommand('xyplot legend hook hout "CAE quality" 0')
                utils.MetaCommand('xyplot legend ymove "CAE quality" 0 1.060000')
                utils.MetaCommand('clipboard copy plot image "{}" {}'.format(window_name, plot.id))
                img_2 = ImageGrab.grabclipboard()
                legend = plot2d.Legend(plot_id, window_name, page_id)
                left,top = legend.get_position()
                width = legend.get_width()
                height = legend.get_height()
                img_2 = img_2.crop((left,top,width+8,height+8))
                image2_path = os.path.join(self.twod_images_report_folder,window_name+"_"+title.get_text().lower()+"_Legend"+".jpeg")
                img_2.save(image2_path,"PNG")
                shape2 = [shape for shape in slide.shapes if shape.name == "Image 1"][0]
                picture = slide.shapes.add_picture(image2_path,shape2.left,shape2.top,width = shape2.width,height = shape2.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('xyplot rlayout "{}" 2'.format(window_name))
                plot.deactivate()
            elif shape.name == "Image 3":
                plot_id = 1
                page_id=0
                title = plot2d.Title(plot_id, window_name, page_id)
                plot = title.get_plot()
                plot.activate()
                utils.MetaCommand('xyplot rlayout "{}" 1'.format(window_name))
                utils.MetaCommand('xyplot curve select "{}" all'.format(window_name))
                utils.MetaCommand('xyplot curve visible and "{}" sel'.format(window_name))
                utils.MetaCommand('clipboard copy plot image "{}" {}'.format(window_name, plot.id))
                img = ImageGrab.grabclipboard()
                img = img.resize((round(shape.width/9525),round(shape.height/9525)))
                image_path = os.path.join(self.twod_images_report_folder,window_name+"_"+title.get_text().lower()+".jpeg")
                img.save(image_path, 'PNG')
                picture = slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                plot.deactivate()
                utils.MetaCommand('xyplot rlayout "{}" 2'.format(window_name))
            elif shape.name == "Table 1":
                plot_id = 0
                page_id=0
                title = plot2d.Title(plot_id, window_name, page_id)
                plot = title.get_plot()
                curvelist = plot.get_curves('all')
                index = 0
                for curve in curvelist:
                    if curve.id == 5:
                        continue
                    index = index+1
                    min_y = curve.get_limit_value_y(specifier = 'min')
                    max_y = curve.get_limit_value_y(specifier = 'max')
                    table_obj = shape.table
                    self.report_composer.add_row(table_obj)
                    row = table_obj.rows[index]
                    text_frame_1 = row.cells[0].text_frame
                    font_1 = text_frame_1.paragraphs[0].font
                    font_1.size = Pt(12)
                    text_frame_1.paragraphs[0].text = str(curve.name).replace(" energy","")

                    text_frame_2 = row.cells[1].text_frame
                    font_2 = text_frame_2.paragraphs[0].font
                    font_2.size = Pt(12)
                    text_frame_2.paragraphs[0].text = "{:.2e}".format(max_y)

                    text_frame_3 = row.cells[2].text_frame
                    font_3 = text_frame_3.paragraphs[0].font
                    font_3.size = Pt(12)
                    text_frame_3.paragraphs[0].text = "{:.2e}".format(min_y)
            elif shape.name == "Table 2":
                table_obj = shape.table
                table_value_dict ={"Termination type":self.general_input.termination_type,
                                    "Computation time":self.general_input.computation_time,
                                    "Core count":self.general_input.core_count,
                                    "Verification mode":self.general_input.verification_mode,
                                    "Compute cluster":self.general_input.compute_cluster}
                index = 0
                for item,value in table_value_dict.items():
                    index = index+1
                    self.report_composer.add_row(table_obj)
                    row = table_obj.rows[index]
                    text_frame_1 = row.cells[0].text_frame
                    font_1 = text_frame_1.paragraphs[0].font
                    font_1.size = Pt(12)
                    text_frame_1.paragraphs[0].text = item

                    if item == "Core count":
                        value = value.split("with")[1].rstrip()
                    text_frame_2 = row.cells[1].text_frame
                    font_2 = text_frame_2.paragraphs[0].font
                    font_2.size = Pt(12)
                    text_frame_2.paragraphs[0].text = value
            #elif shape.name == "Image 1":







        return 0

    def threed_data_reporting(self):
        """
        threed_data_reporting [summary]

        [extended_summary]

        Returns:
            [type]: [description]
        """
        critical_sections_data = self.metadb_3d_input.critical_sections
        # for section,value in critical_sections_data.items():
        #     for key,vvalue in value.items():
        #         if key == "hes":

        return 0

    def twod_data_reporting(self):
        """
        twod_data_reporting [summary]

        [extended_summary]

        Returns:
            [type]: [description]
        """
        from PIL import ImageGrab

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
            #utils.MetaCommand('xyplot rlayout "{}" 1'.format(window_name))
            image_path = os.path.join(self.twod_images_report_folder,window_name+"_"+curve.name.lower()+".jpeg")
            if not os.path.exists(os.path.dirname(image_path)):
                print(os.path.dirname(image_path))
                os.makedirs(os.path.dirname(image_path))
            utils.MetaCommand('clipboard copy plot image "{}" {}'.format(window_name, plot.id))
            img = ImageGrab.grabclipboard()
            img.save(image_path, 'PNG')
            #utils.MetaCommand('write jpeg "{}" 100'.format(image_path))
            #utils.MetaCommand('xyplot rlayout "{}" {}'.format(window_name,window_layout))

        return 0


class PPTXReportComposer():
    def __init__(self, report_name, template_pptx):

        self.report_name = report_name
        self.template_pptx = template_pptx
        self.prs_obj = None

    def create_prs_obj(self):
        """ Creates the PPTx report using
        the python-pptx module
        """
        from pptx import Presentation

        # Instantiate
        if not self.prs_obj:
            self.prs_obj = Presentation(self.template_pptx)

        return 0

    # def add_slide(self,layout_name):

    #     slide = self.prs_obj.slides.add_slide(layout_name)

    #     return slide

    @staticmethod
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

    @staticmethod
    def remove_row(table,row_to_delete):
        """
        remove_row [summary]

        [extended_summary]

        Args:
            table ([type]): [description]
            row_to_delete ([type]): [description]
        """
        table._tbl.remove(row_to_delete._tr)

    def save_pptx(self, pptx_filepath, datestamp=""):
        """ Saves the PPTx at the given filepath
        with the given datestamp

        Args:
            pptx_filepath (str): Absolute path to the pptx file for saving.
            datestamp (str, optional): Date stamp at the bottom right of slide. Defaults to "".

        Returns:
            int: 0 always.
        """
        from pptx.enum.text import PP_ALIGN

        # Get current date if not provided
        if not datestamp:
            datestamp = time.strftime("%B %d, %Y")

        # Set the date in the ppt master slide
        master_slide = self.prs_obj.slide_master
        for shape in master_slide.shapes:
            try:
                if shape.text == "Date_Stamp":
                    shape.text = datestamp
                    shape.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

            except AttributeError:
                continue

        # Save the pptx
        self.prs_obj.save(pptx_filepath)

        return 0
