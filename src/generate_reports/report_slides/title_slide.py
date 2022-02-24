# PYTHON script
"""
    _summary_

_extended_summary_

Returns:
    _type_: _description_
"""

class TitleSlide():

    def __init__(self,
                slide,
                windows,
                general_input,
                metadb_2d_input,
                metadb_3d_input,
                template_file,
                twod_images_report_folder,
                threed_images_report_folder,
                ppt_report_folder) -> None:
        self.slide = slide
        self.shapes = slide.shapes
        self.windows = windows
        self.general_input = general_input
        self.metadb_2d_input = metadb_2d_input
        self.metadb_3d_input = metadb_3d_input
        self.template_file = template_file
        self.twod_images_report_folder = twod_images_report_folder
        self.threed_images_report_folder = threed_images_report_folder
        self.ppt_report_folder = ppt_report_folder

    def setup(self):
        """
        setup _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """

        return 0

    def edit(self):
        """
        edit _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """
        from datetime import datetime
        from pptx.util import Pt

        self.setup()

        for shape in self.shapes:
            if shape.name == "Table 2":
                table_obj = shape.table
                rows = table_obj.rows
                ordinal = lambda n: "%d%s" % (n,"tsnrhtdd"[(n//10%10!=1)*(n%10<4)*n%10::4])
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

        self.revert()

        return 0

    def revert(self):
        """
        revert _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """

        return 0