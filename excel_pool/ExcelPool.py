#
# Copyright ©2024 Dana Basken
#

import asyncio
import logging
import multiprocessing
import os
import queue
import time
from typing import List
import win32com.client as win32
import pythoncom
import uuid
from win32com.universal import com_error
from excel_pool.ExcelPoolTask import ExcelPoolTask
from config import settings
from task_handlers.TaskHandlerManager import TaskHandlerManager
from utilities.LRUCache import LRUCache

logger = logging.getLogger(__name__)

class ExcelPool(object):

    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(ExcelPool, cls).__new__(cls, *args, **kwargs)
            cls._instance._initialize()
        return cls._instance

    def _initialize(self):
        self._workers: List[multiprocessing.Process] = []
        self._task_status = LRUCache(size=2000, ttl=120, on_event=self._on_cache_event)
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

    def shutdown(self, signum, frame):
        for worker in self._workers:
            logger.debug(f"joining {worker.pid}...")
            worker.join(10)

    def _on_cache_event(self, event):
        logger.debug(f"cache event: {event}")

    def _start_workers(self, worker_count: int):
        for index in range(worker_count): self._start_worker()
        logger.debug(f"started {worker_count} workers")

    def _start_worker(self):
        logger.debug("starting worker...")
        worker_process = multiprocessing.Process(target=self._worker, args=(self._requests, self._responses))
        worker_process.start()
        self._workers.append(worker_process)
        logger.debug(f"worker {worker_process.pid} started.")

    def _clear_worker(self, pid: int):
        self._workers = [worker for worker in self._workers if worker.pid != pid]

    async def _response_handler(self, responses):
        loop = asyncio.get_event_loop()
        while True:
            response: any = await loop.run_in_executor(None, responses.get)
            if not response: continue
            if response.get("error"):
                logger.error(f"_response_handler {response}")
            else:
                logger.info(f"_response_handler {response}")
            if response.get("state") == "died":
                self._clear_worker(response.get("process_id"))
                self._start_worker()
        if "id" in response:
                self._task_status.put(response.get("id"), response)

    # this method runs in a child process, so doesn't have normal access to logger or anything in the ExcelPool instance
    @staticmethod
    def _worker(requests: multiprocessing.Queue, responses: multiprocessing.Queue):
        task = {"id": f"EXCEL_{os.getpid()}"}
        try:
            pythoncom.CoInitialize()
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            responses.put({"id": f"EXCEL_{os.getpid()}", "process_id": os.getpid(), "state": "started"})
            while True:
                try:
                    _ = excel.Visible # this will raise com_error if the underlying Excel has died
                    task = requests.get(timeout=60)
                    if not task: continue
                    if not "excel_pool_task" in task: continue
                    excel_pool_task: ExcelPoolTask = task["excel_pool_task"]
                    handler = TaskHandlerManager.get_handler_for_task(excel_pool_task)
                    handler.run(task, excel_pool_task, excel, responses)
                except KeyboardInterrupt:
                    excel.Quit()
                    del excel
                    break
                except com_error as exception: # most likely Excel died
                    responses.put({"id": task["id"], "error": str(exception), "process_id": os.getpid(), "state": "died"})
                    del excel
                    break
                except queue.Empty:
                    responses.put({"id": f"EXCEL_{os.getpid()}", "process_id": os.getpid(), "state": "empty_queue"})
                except Exception as exception:
                    responses.put({"id": task["id"], "error": str(exception), "state": "failure"})
        except Exception as exception:
            responses.put({"id": f"EXCEL_{os.getpid()}", "error": str(exception), "state": "failure"})
        finally:
            pythoncom.CoUninitialize()
            responses.put({"id": f"EXCEL_{os.getpid()}", "process_id": os.getpid(), "state": "exited"})
