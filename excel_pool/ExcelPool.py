#
# Copyright ©2024 Dana Basken
#

import asyncio
import logging
import multiprocessing
import os
import tempfile
import time
import exceltypes
import win32com.client as win32
import pythoncom
import uuid
from excel_pool.ExcelPoolTask import ExcelPoolTask
from config import settings
from models.DurableIds import DurableIds
from models.Worksheet import Worksheet
from services.SharepointGroupService import get_drive_item
from requests import request
from tenacity import retry, stop_after_attempt, wait_random_exponential

logger = logging.getLogger(__name__)

class ExcelPool(object):

    _instance = None
    _task_status: any = {}

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(ExcelPool, cls).__new__(cls, *args, **kwargs)
            cls._instance._initialize()
        return cls._instance

    def _initialize(self):
        self._workers = []
        self._requests = multiprocessing.Queue()
        self._responses = multiprocessing.Queue()
        self._start_workers(settings.get("excel_pool.workers", 5))
        asyncio.create_task(self._response_handler(self._responses))

    def add_task(self, excel_pool_task: ExcelPoolTask) -> str:
        id = str(uuid.uuid4())
        self._responses.put({"id": id, "state": "pending"})
        self._requests.put({"id": id, "excel_pool_task": excel_pool_task})
        return id

    def get_task_status(self, id: str) -> any:
        return self._task_status[id]

    def _start_workers(self, worker_count: int):
        logger.debug(f"starting {worker_count} workers...")
        for index in range(worker_count):
            worker_process = multiprocessing.Process(target=self._worker, args=(self._requests, self._responses))
            worker_process.start()
            self._workers.append(worker_process)
        logger.debug("workers started.")

    async def _response_handler(self, responses):
        loop = asyncio.get_event_loop()
        while True:
            status: any = await loop.run_in_executor(None, responses.get)
            if "error" in status:
                logger.error(f"_response_handler {status}")
            else:
                logger.info(f"_response_handler {status}")
            self._task_status[id] = status

    @staticmethod
    def _worker(requests: multiprocessing.Queue, responses: multiprocessing.Queue):

        @retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
        def _open_workbook(excel: exceltypes.Application, file_path: str) -> exceltypes.Workbook:
            return excel.Workbooks.Open(file_path)

        @retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
        def _get_durable_ids(worksheet: Worksheet) -> DurableIds:
            return worksheet.get_durable_ids()

        @retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
        def _set_durable_id_values(durable_ids: DurableIds, durable_id: str, values: any):
            durable_ids.set_durable_id_values(durable_id, values)

        @retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
        def _calculate(workbook: exceltypes.Workbook):
            workbook.Application.CalculateFull()

        def _handle_task(task: any, responses: multiprocessing.Queue):
            start_time = time.time()
            task_id = task["id"]
            responses.put({"id": task_id, "state": "running", "phase": "starting", "duration": time.time() - start_time})
            excel_pool_task: ExcelPoolTask = task["excel_pool_task"]

        try:
            pythoncom.CoInitialize()
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            while True:
                try:
                    task = requests.get()
                    if task == "STOP": break

                    # TODO: move this into a _handle_task function, which can then figure out what the task should do
                    # _handle_task(task, responses)

                    start_time = time.time()
                    task_id = task["id"]
                    responses.put({"id": task_id, "state": "running", "phase": "starting", "duration": time.time() - start_time})
                    excel_pool_task: ExcelPoolTask = task["excel_pool_task"]

                    responses.put({"id": task_id, "state": "running", "phase": "get_drive_item", "duration": time.time() - start_time})
                    drive_item = get_drive_item(excel_pool_task.site_id, excel_pool_task.item_id)

                    responses.put({"id": task_id, "state": "running", "phase": "download_drive_item", "duration": time.time() - start_time})
                    response = request("GET", drive_item["@microsoft.graph.downloadUrl"], stream=True)
                    responses.put({"id": task_id, "state": "running", "phase": "stream_drive_item", "duration": time.time() - start_time})
                    with tempfile.NamedTemporaryFile(delete=False) as temp_file:
                        for chunk in response.iter_content(chunk_size=8192): temp_file.write(chunk)
                        temp_file_path = temp_file.name

                    responses.put({"id": task_id, "state": "running", "phase": "open_workbook", "duration": time.time() - start_time})
                    workbook = _open_workbook(excel, temp_file_path)

                    responses.put({"id": task_id, "state": "running", "phase": "construct_worksheet", "duration": time.time() - start_time})
                    worksheet = Worksheet("M - Monthly", workbook)

                    responses.put({"id": task_id, "state": "running", "phase": "get_durable_ids", "duration": time.time() - start_time})
                    durable_ids = _get_durable_ids(worksheet)

                    responses.put({"id": task_id, "state": "running", "phase": "get_durable_id", "duration": time.time() - start_time})
                    income_statement_returns, durable_id_type = durable_ids.get_durable_id("incomeStatement.returns")

                    # TODO: do something to income_statement_returns

                    responses.put({"id": task_id, "state": "running", "phase": "set_durable_id_values", "duration": time.time() - start_time})
                    _set_durable_id_values(durable_ids, "incomeStatement.returns", income_statement_returns)

                    responses.put({"id": task_id, "state": "running", "phase": "calculate_workbook", "duration": time.time() - start_time})
                    _calculate(workbook)

                    responses.put({"id": task_id, "state": "running", "phase": "get_durable_ids_2", "duration": time.time() - start_time})
                    durable_ids = _get_durable_ids(worksheet)

                    # TODO: send durable_ids to DWH

                    responses.put({"id": task_id, "state": "running", "phase": "close_workbook", "duration": time.time() - start_time})
                    workbook.Close(SaveChanges=False) # TODO: RETRY

                    responses.put({"id": task_id, "state": "running", "phase": "delete_workbook", "duration": time.time() - start_time})
                    os.remove(temp_file_path)

                    responses.put({"id": task_id, "result": {}, "state": "complete", "duration": time.time() - start_time})
                except KeyboardInterrupt:
                    excel.Quit()
                    break
                except Exception as exception:
                    print("error", exception)
                    responses.put({"id": task_id, "error": str(exception), "state": "complete"})
        except Exception as exception:
            print(exception)
        finally:
            pythoncom.CoUninitialize()
            responses.put({"id": "EXCEL", "state": "exited"})
