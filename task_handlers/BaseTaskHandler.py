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

class BaseTaskHandler:

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
        try:
            self.set_up()
            self.process()
        finally:
            self.shutdown()
            self.add_response("success")

    def set_up(self):
        self.start_time = time.time()
        self.task_id = self.task["id"]
        # read the Sharepoint drive item from MS Graph API:
        self.add_response("running", "get_drive_item")
        self.drive_item = self.get_drive_item(self.excel_pool_task.group_id, self.excel_pool_task.item_id)
        self.temp_file_path = self.download_drive_item_into_temp_file()
        # open the local temp file in Excel and set our workbook variable:
        self.add_response("running", "open_workbook")
        self.workbook = self.open_workbook(self.excel, self.temp_file_path)

    def process(self):
        pass

    def shutdown(self):
        # close Workbook in Excel, if it was opened:
        self.close_workbook()
        # delete the temp file, if it was created:
        if self.temp_file_path:
            self.add_response("running", "delete_workbook")
            os.remove(self.temp_file_path)

    @retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
    def download_drive_item_into_temp_file(self) -> str:
        temp_file_path = None
        try:
            # read the Sharepoint drive item from MS Graph API:
            self.add_response("running", "get_drive_item")
            self.drive_item = self.get_drive_item(self.excel_pool_task.group_id, self.excel_pool_task.item_id)
            if not "@microsoft.graph.downloadUrl" in self.drive_item: raise Exception(f"Sharepoint Drive Item did not have '@microsoft.graph.downloadUrl'")
            # use the "@microsoft.graph.downloadUrl" URL to stream the contents of the Sharepoint drive item down:
            response = request("GET", self.drive_item["@microsoft.graph.downloadUrl"], stream=True)
            self.add_response("running", "stream_drive_item")
            # create a temp file and store the file stream into it:
            with tempfile.NamedTemporaryFile(delete=False) as temp_file:
                for chunk in response.iter_content(chunk_size=8192): temp_file.write(chunk)
                temp_file_path = temp_file.name
            return temp_file_path
        except:
            if temp_file_path: os.remove(self.temp_file_path)
            raise

    @retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
    def close_workbook(self):
        if self.workbook:
            self.add_response("running", "close_workbook")
            self.workbook.Close(SaveChanges=False)

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
