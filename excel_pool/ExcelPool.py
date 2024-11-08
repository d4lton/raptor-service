#
# Copyright ©2024 Dana Basken
#

import asyncio
import logging
import multiprocessing
import os
import time
import win32com.client as win32
import pythoncom
import uuid
from excel_pool.ExcelPoolTask import ExcelPoolTask
from config import settings
from task_handlers.HandlerManager import HandlerManager

logger = logging.getLogger(__name__)

class ExcelPool(object):

    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(ExcelPool, cls).__new__(cls, *args, **kwargs)
            cls._instance._initialize()
        return cls._instance

    def _initialize(self):
        self._workers = []
        self._task_status = {}
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
        return self._task_status.get(id)

    async def wait_for_task_completion(self, id: str, max_ttl: int = 600) -> any:
        start_time = time.time()
        while True:
            if time.time() - start_time >= max_ttl: raise TimeoutError(f"Task took longer than {max_ttl} seconds")
            task_status = self.get_task_status(id)
            if task_status:
                if task_status["state"] == "success" or task_status["state"] == "failure": return task_status
            await asyncio.sleep(1)

    def _start_workers(self, worker_count: int):
        for index in range(worker_count):
            worker_process = multiprocessing.Process(target=self._worker, args=(self._requests, self._responses))
            worker_process.start()
            self._workers.append(worker_process)
        logger.debug(f"started {worker_count} workers")

    async def _response_handler(self, responses):
        loop = asyncio.get_event_loop()
        while True:
            status: any = await loop.run_in_executor(None, responses.get)
            if "failure" in status:
                logger.error(f"_response_handler {status}")
            else:
                logger.info(f"_response_handler {status}")
            if "id" in status:
                self._task_status[status["id"]] = status

    @staticmethod
    def _worker(requests: multiprocessing.Queue, responses: multiprocessing.Queue):
        try:
            pythoncom.CoInitialize()
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            responses.put({"id": "EXCEL", "process_id": os.getpid(), "state": "started"})
            task = {}
            while True:
                try:
                    task = requests.get()
                    if task == "STOP": break
                    excel_pool_task: ExcelPoolTask = task["excel_pool_task"]
                    handler = HandlerManager.get_handler_for_task(excel_pool_task)
                    handler.run(task, excel_pool_task, excel, responses)
                except KeyboardInterrupt:
                    excel.Quit()
                    break
                except Exception as exception:
                    responses.put({"id": task["id"], "error": str(exception), "state": "failure"})
        except Exception as exception:
            print(exception)
        finally:
            pythoncom.CoUninitialize()
            responses.put({"id": "EXCEL", "process_id": os.getpid(), "state": "exited"})
