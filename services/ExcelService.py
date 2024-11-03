import asyncio
import queue
import threading

import pythoncom
from win32com.client import Dispatch

excel_service_app = None
excel_service_thread = None
excel_service_queue = queue.Queue()

def start_excel_service():
    global excel_service_thread
    print("lifespan startup")
    excel_service_thread = threading.Thread(target=excel_worker, daemon=True)
    excel_service_thread.start()

def shutdown_excel_service():
    global excel_service_app
    print("lifespan shutdown")
    if excel_service_app:
        excel_service_app.Quit()
        excel_service_app = None

def excel_worker():
    global excel_service_app
    print("starting excel")
    pythoncom.CoInitialize()
    excel_service_app = Dispatch("Excel.Application")
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    print("creating task")
    loop.create_task(process_tasks())
    print("looping")
    loop.run_forever()

async def process_tasks():
    while True:
        print("awaiting task")
        future = excel_service_queue.get()
        try:
            print("running task")
            future.set_result(process_excel_task())
        except Exception as e:
            print("task exception", e)
            future.set_exception(e)
        finally:
            print("task finally")
            excel_service_queue.task_done()

def process_excel_task():
    print("process_excel_task start")
    pythoncom.CoInitialize()
    workbook = excel_service_app.Workbooks.Open(r"C:\Users\dana\raptor-service\test.xlsx")
    workbook.Close(SaveChanges=False)
    print("process_excel_task end")
    return {}
