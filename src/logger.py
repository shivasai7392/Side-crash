#PYTHON SCRIPT

import os
import logging
import datetime

class SideCrashLogger():
    def __init__(self,file_path,name = "side_crash_logger"):
        #name of the logger
        self.name = name

        #beta logger format for logging
        self.side_crash_logger_format = CustomFormatter('%(message)s')

        #getting logger object as name 'beta_logger'
        self.log = logging.getLogger(name)

        if not self.log.handlers:
            #setting log level as INFO
            self.log.setLevel(logging.INFO)

            #adding console stream handler for logging
            self.console_handler = self.get_console_handler()
            self.set_formatter(self.console_handler,self.side_crash_logger_format if name == "side_crash_logger" else None)
            self.log.addHandler(self.console_handler)

            #adding file stream handler for logging
            if file_path is not None:
                #log output file preperation
                current_datetime = datetime.datetime.now()
                output_folder = os.path.dirname(file_path)
                if not os.path.exists(output_folder):
                    os.mkdir(output_folder)
                self.log_file = "{}_{}.log".format(file_path,current_datetime.strftime('%Y-%d-%m'))
                if not os.path.exists(self.log_file):
                    file_object = open(self.log_file, 'w+')
                    if os.stat(self.log_file).st_size == 0:
                        initial_str = """##################################################
        #      Copyright BETA CAE Systems USA Inc.,      #
        #      2022 All Rights Reserved                  #
        ##################################################

        Side Crash Automation Log file

        Log Session Report {}
        --------------------------
        --------------------------

        """.format(current_datetime.strftime("%H:%M:%S"))
                        file_object.write(initial_str)
                    file_object.close()
                else:
                    file_object = open(self.log_file, 'a')
                    if os.stat(self.log_file).st_size == 0:
                        initial_str = """##################################################
        #      Copyright BETA CAE Systems USA Inc.,      #
        #      2022 All Rights Reserved                  #
        ##################################################

        Side Crash Automation Log file

        Log Session Report {}
        --------------------------
        --------------------------

        """.format(current_datetime.strftime("%H:%M:%S"))
                    else:
                        initial_str = """Log Session Report {}
        --------------------------
        --------------------------

        """.format(current_datetime.strftime("%H:%M:%S"))
                    file_object.write(initial_str)
                    file_object.close()

                self.file_handler = self.get_file_handler(self.log_file)
                self.set_formatter(self.file_handler,self.side_crash_logger_format if name == "side_crash_logger" else  None)
                self.log.addHandler(self.file_handler)
            else:
                self.log.info("ERROR : META 2D variable 'pM' is not available or valid. Please update.")
                self.log.info("")

    def set_formatter(self,handler,format):
        """
        This method is used to set the formatter to the handler

        Returns:
            int: 0
        """
        handler.setFormatter(format)
        return 0

    def shutdown(self):
        """
        shutdown _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """
        self.log.handlers = []

        return 0

    def get_console_handler(self):
        """
        This method is used to get the console handler

        Returns:
            Handler: Console handler
        """
        console_handler = logging.StreamHandler()
        return console_handler

    def get_file_handler(self,file):
        """
        This method is used to get the file handler

        Returns:
            Handler: File handler
        """
        file_handler = logging.FileHandler(file)
        return file_handler

class CustomFormatter(logging.Formatter):
    """ Custom Formatter does these 2 things:
    1. Overrides 'funcName' with the value of 'func_name_override', if it exists.
    2. Overrides 'filename' with the value of 'file_name_override', if it exists.
    """

    def format(self, record):
        if hasattr(record, 'func_name_override'):
            record.funcName = record.func_name_override
        if hasattr(record, 'file_name_override'):
            record.filename = record.file_name_override
        return super(CustomFormatter, self).format(record)
