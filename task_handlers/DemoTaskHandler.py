#
# Copyright ©2024 Dana Basken
#

import os
import time
import multiprocessing
import exceltypes
from requests import request
import tempfile
from excel_pool.ExcelPoolTask import ExcelPoolTask
from task_handlers.BaseHandler import BaseHandler
from services.SharepointGroupService import get_drive_item
from models.Worksheet import Worksheet
from utilities.HandlerCommon import calculate, get_durable_ids, open_workbook, set_durable_id_values

class DemoTaskHandler(BaseHandler):

    def run(self, task: any, excel_pool_task: ExcelPoolTask, excel: exceltypes.Application, responses: multiprocessing.Queue):
        start_time = time.time()
        task_id = task["id"]

        responses.put({"id": task_id, "state": "running", "phase": "get_drive_item", "duration": time.time() - start_time})
        drive_item = get_drive_item(excel_pool_task.site_id, excel_pool_task.item_id)

        responses.put({"id": task_id, "state": "running", "phase": "download_drive_item", "duration": time.time() - start_time})
        response = request("GET", drive_item["@microsoft.graph.downloadUrl"], stream=True)
        responses.put({"id": task_id, "state": "running", "phase": "stream_drive_item", "duration": time.time() - start_time})
        with tempfile.NamedTemporaryFile(delete=False) as temp_file:
            for chunk in response.iter_content(chunk_size=8192): temp_file.write(chunk)
            temp_file_path = temp_file.name

        responses.put({"id": task_id, "state": "running", "phase": "open_workbook", "duration": time.time() - start_time})
        workbook = open_workbook(excel, temp_file_path)

        responses.put({"id": task_id, "state": "running", "phase": "construct_worksheet", "duration": time.time() - start_time})
        worksheet = Worksheet("M - Monthly", workbook)

        responses.put({"id": task_id, "state": "running", "phase": "get_durable_ids", "duration": time.time() - start_time})
        durable_ids = get_durable_ids(worksheet)

        responses.put({"id": task_id, "state": "running", "phase": "get_durable_id", "duration": time.time() - start_time})
        income_statement_returns, durable_id_type = durable_ids.get_durable_id("incomeStatement.returns")

        # TODO: do something to income_statement_returns

        responses.put({"id": task_id, "state": "running", "phase": "set_durable_id_values", "duration": time.time() - start_time})
        set_durable_id_values(durable_ids, "incomeStatement.returns", income_statement_returns)

        responses.put({"id": task_id, "state": "running", "phase": "calculate_workbook", "duration": time.time() - start_time})
        calculate(workbook)

        responses.put({"id": task_id, "state": "running", "phase": "get_durable_ids_2", "duration": time.time() - start_time})
        durable_ids = get_durable_ids(worksheet)

        # TODO: send durable_ids to DWH

        responses.put({"id": task_id, "state": "running", "phase": "close_workbook", "duration": time.time() - start_time})
        workbook.Close(SaveChanges=False) # TODO: RETRY

        responses.put({"id": task_id, "state": "running", "phase": "delete_workbook", "duration": time.time() - start_time})
        os.remove(temp_file_path)

        responses.put({"id": task_id, "result": {}, "state": "complete", "duration": time.time() - start_time})