#PYTHON SCRIPT

import sys
import functools
import datetime
from openpyxl import Workbook
workbook = Workbook()

class SideCrashLogger():

    @staticmethod
    def excel_resource_log_decorator(_func=None,**kwargss):
        def excel_resource_log_decorator_info(func):
            @functools.wraps(func)
            def wrapper(*args, **kwargs):
                global workbook
                start_datetime = datetime.datetime.now()
                #logging the event
                worksheet = workbook.active
                worksheet.append(["-------{}-------".format(kwargss["Description"]),0])
                worksheet.append([f"START TIME : {start_datetime.strftime('%Y-%m-%d %H:%M:%S')}",1])
                try:
                    value = func(*args, **kwargs)
                    #loggging return value from function
                    end_datetime = datetime.datetime.now()
                    print(worksheet.append([f"END TIME : {end_datetime.strftime('%Y-%m-%d %H:%M:%S')}",2]))
                    time_taken_string = str(end_datetime - start_datetime)
                    worksheet.append([f"TOTAL TIME TAKEN:{time_taken_string}",3])

                except:
                    worksheet.append([f"Exception: {str(sys.exc_info()[1])}"])
                    raise
                return value
            return wrapper
        if _func is None:
            return excel_resource_log_decorator_info
        else:
            return excel_resource_log_decorator_info(_func)

    @staticmethod
    def save_workbook(file_path):
        """
        save_workbook _summary_

        _extended_summary_

        Returns:
            _type_: _description_
        """
        workbook.save(file_path)
        return 0
