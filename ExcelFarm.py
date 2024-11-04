#
# Copyright ©2024 Dana Basken
#

import asyncio
import multiprocessing
import win32com.client as win32
import pythoncom
import uuid

class ExcelFarm(object):

    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(ExcelFarm, cls).__new__(cls, *args, **kwargs)
            cls._instance._initialize()
        return cls._instance

    def _initialize(self):
        self._workers = []
        self._requests = multiprocessing.Queue()
        self._responses = multiprocessing.Queue()
        self._start_workers(5)
        asyncio.create_task(self._response_handler(self._responses))

    def add_task(self, workbook_path: str) -> str:
        id = str(uuid.uuid4())
        self._requests.put({"id": id, "workbook_path": workbook_path})
        return id

    def _start_workers(self, worker_count: int):
        for index in range(worker_count):
            worker_process = multiprocessing.Process(target=self._worker, args=(self._requests, self._responses))
            worker_process.start()
            self._workers.append(worker_process)

    @staticmethod
    async def _response_handler(responses):
        loop = asyncio.get_event_loop()
        while True:
            result = await loop.run_in_executor(None, responses.get)
            print(result)

    @staticmethod
    def _worker(requests: multiprocessing.Queue, responses: multiprocessing.Queue):
        try:
            pythoncom.CoInitialize()
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            while True:
                try:
                    task = requests.get()
                    print("task", task)
                    if task == "STOP": break
                    task_id = task["id"]
                    workbook_path = task["workbook_path"]
                    print("opening workbook", workbook_path)
                    workbook = excel.Workbooks.Open(workbook_path)
                    # TODO: something
                    print("closing workbook", workbook_path)
                    workbook.Close(SaveChanges=False)
                    responses.put({"id": task_id, "result": {}})
                    print("task complete", task)
                except Exception as exception:
                    print("error", exception)
                    responses.put({"id": task_id, "error": str(exception)})
            print("quitting excel")
            excel.Quit()
        except Exception as exception:
            print(exception)
        finally:
            pythoncom.CoUninitialize()
