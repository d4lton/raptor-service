#
# Copyright ©2024 Dana Basken
#

import multiprocessing
import exceltypes
from excel_pool.ExcelPoolTask import ExcelPoolTask

class BaseHandler:

    def run(self, task: any, excel_pool_task: ExcelPoolTask, excel: exceltypes.Application, responses: multiprocessing.Queue):
        pass
