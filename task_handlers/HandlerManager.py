#
# Copyright ©2024 Dana Basken
#

from excel_pool.ExcelPoolTask import ExcelPoolTask
from task_handlers.DemoTaskHandler import DemoTaskHandler

class HandlerManager:

    @staticmethod
    def get_handler_for_task(excel_pool_task: ExcelPoolTask):
        match excel_pool_task.type:
            case "demo":
                return DemoTaskHandler() # TODO - refactor constructor args for BaseHandler and children
            case _:
                raise Exception(f"unhandled task type: {excel_pool_task.type}")
