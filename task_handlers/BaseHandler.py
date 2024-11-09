#
# Copyright ©2024 Dana Basken
#

import multiprocessing
import os
import tempfile
import time
from requests import request
import exceltypes
from models.Worksheet import Worksheet
from tenacity import retry, stop_after_attempt, wait_random_exponential
from excel_pool.ExcelPoolTask import ExcelPoolTask
from models.DurableIds import DurableIds
from services import MicrosoftGraphQLClient

class BaseHandler:

    def __init__(self):
        self.drive_item = None
        self.workbook: exceltypes.Workbook | None = None
        self.temp_file_path: str | None = None
        self.task_id: str | None = None
        self.start_time: float | None = None
        self.task: any = None
        self.excel_pool_task: ExcelPoolTask | None = None
        self.excel: exceltypes.Application | None = None
        self.responses: multiprocessing.Queue | None = None

    def run(self, task: any, excel_pool_task: ExcelPoolTask, excel: exceltypes.Application, responses: multiprocessing.Queue):
        self.task = task
        self.excel_pool_task = excel_pool_task
        self.excel = excel
        self.responses = responses
        self.set_up()
        self.process()
        self.shutdown()

    def set_up(self):
        self.start_time = time.time()
        self.task_id = self.task["id"]
        self.add_response("running", "get_drive_item")
        self.drive_item = self.get_drive_item(self.excel_pool_task.site_id, self.excel_pool_task.item_id)
        self.add_response("running", "download_drive_item")
        response = request("GET", self.drive_item["@microsoft.graph.downloadUrl"], stream=True)
        self.add_response("running", "stream_drive_item")
        with tempfile.NamedTemporaryFile(delete=False) as temp_file:
            for chunk in response.iter_content(chunk_size=8192): temp_file.write(chunk)
            self.temp_file_path = temp_file.name
        self.add_response("running", "open_workbook")
        self.workbook = self.open_workbook(self.excel, self.temp_file_path)

    def process(self):
        pass

    def shutdown(self):
        self.add_response("running", "close_workbook")
        self.workbook.Close(SaveChanges=False) # TODO: RETRY
        self.add_response("running", "delete_workbook")
        os.remove(self.temp_file_path)
        self.add_response("success")

    def add_response(self, state: str, phase: str | None = None):
        self.responses.put({"id": self.task_id, "state": state, "phase": phase, "duration": time.time() - self.start_time})

    @staticmethod
    @retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
    def get_drive_item(group_id: str, item_id: str) -> any:
        return MicrosoftGraphQLClient.request("GET", f"/groups/{group_id}/drive/items/{item_id}")

    @staticmethod
    @retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
    def open_workbook(excel: exceltypes.Application, file_path: str) -> exceltypes.Workbook:
        return excel.Workbooks.Open(file_path)

    @staticmethod
    @retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
    def get_durable_ids(worksheet: Worksheet) -> DurableIds:
        return worksheet.get_durable_ids()

    @staticmethod
    @retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
    def set_durable_id_values(durable_ids: DurableIds, durable_id: str, values: any):
        durable_ids.set_durable_id_values(durable_id, values)

    @staticmethod
    @retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
    def calculate(workbook: exceltypes.Workbook):
        workbook.Application.CalculateFull()
