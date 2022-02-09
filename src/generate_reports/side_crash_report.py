from html import entities
import os

from meta import utils
from meta import windows
from meta import plot2d
import time
import re

class SideCrashReport():

    def __init__(self,windows,general_input,metadb_2d_input,metadb_3d_input) -> None:
        self.windows = windows
        self.general_input = general_input
        self.metadb_2d_input = metadb_2d_input
        self.metadb_3d_input = metadb_3d_input
        self.get_reporting_folders()

    def get_reporting_folders(self):
        self.twod_images_report_folder = os.path.join(os.path.dirname(self.general_input.report_directory),"2d-data-images")
        self.threed_images_report_folder = os.path.join(os.path.dirname(self.general_input.report_directory),"3d-data-images")
        self.threed_videos_report_folder = os.path.join(os.path.dirname(self.general_input.report_directory),"3d-data-videos")
        self.excel_bom_report_folder = os.path.join(os.path.dirname(self.general_input.report_directory),"excel-bom")
        self.ppt_report_folder = os.path.join(os.path.dirname(self.general_input.report_directory),"reports")

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

        template_file = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))),"res",self.general_input.source_template_file_directory.replace("/","",1),self.general_input.source_template_file_name).replace("/",os.sep)
        self.report_composer = PPTXReportComposer(report_name="Run1",template_pptx=template_file)
        self.report_composer.create_prs_obj()

        self.edit_title_slide(self.report_composer.prs_obj.slides[0])
        self.edit_cae_quality_slide(self.report_composer.prs_obj.slides[1])
        #self.edit_executive_slide(self.report_composer.prs_obj.slides[2])

        output_directory = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))),"res",self.ppt_report_folder.replace("/","",1)).replace("/",os.sep)
        if not os.path.exists(output_directory):
            os.makedirs(output_directory)
        file_name = os.path.join(output_directory,"output.pptx")
        self.report_composer.save_pptx(file_name)

        return 0

    def edit_executive_slide(self,slide):

        from PIL import ImageGrab

        for shape in slide.shapes:
            if shape.name == "Image 2":
                data = self.metadb_3d_input.critical_sections["f21_upb_inner"]
                prop_names = data["hes"]
                re_props = prop_names.split(",")
                entities = []
                for re_prop in re_props:
                    utils.MetaCommand('window maximize "MetaPost"')
                    utils.MetaCommand('plane options onlysection enable DEFAULT_PLANE_YZ')
                    utils.MetaCommand('plane options slicewidth 500.000000 DEFAULT_PLANE_YZ ')
                    entities.extend(self.metadb_3d_input.get_props(re_prop))
                self.metadb_3d_input.hide_all()
                self.metadb_3d_input.show_only_props(entities)
                utils.MetaCommand('0:options state variable "serial=1"')
                utils.MetaCommand('grstyle scalarfringe enable')
                utils.MetaCommand('view default front')

                utils.MetaCommand('options fringebar off')
                utils.MetaCommand('clipboard copy image "MetaPost"')
                img = ImageGrab.grabclipboard()
                img = img.resize((round(shape.width/9525),round(shape.height/9525)))
                src_parent_folder = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
                image_path = os.path.join(src_parent_folder,"res",self.threed_images_report_folder.replace("/","",1).replace("/",os.sep),"MetaPost"+"_"+"f21_upb_inner".lower()+".jpeg")
                if not os.path.exists(os.path.dirname(image_path)):
                    print(os.path.dirname(image_path))
                    os.makedirs(os.path.dirname(image_path))
                img.save(image_path, 'PNG')
                picture = slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0

        return 0

    def edit_title_slide(self,slide):
        for shape in slide.shapes:
            if shape.name == "TextBox 1":
                shape.text = "PT"
            elif shape.name == "TextBox 2":
                shape.text = "fr"


        return 0

    def edit_cae_quality_slide(self,slide):

        from PIL import ImageGrab

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
                src_parent_folder = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
                image_path = os.path.join(src_parent_folder,"res",self.twod_images_report_folder.replace("/","",1).replace("/",os.sep),window_name+"_"+title.get_text().lower()+".jpeg")
                img.save(image_path, 'PNG')
                picture = slide.shapes.add_picture(image_path,shape.left,shape.top,width = shape.width,height = shape.height)
                picture.crop_left = 0
                picture.crop_right = 0
                utils.MetaCommand('xyplot rlayout "{}" 2'.format(window_name))


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
            src_parent_folder = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
            image_path = os.path.join(src_parent_folder,"res",self.twod_images_report_folder.replace("/","",1).replace("/",os.sep),window_name+"_"+curve.name.lower()+".jpeg")
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