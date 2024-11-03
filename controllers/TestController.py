#
# Copyright ©2024 Dana Basken
#
import asyncio
from contextlib import asynccontextmanager
import time
import pythoncom
import win32com.client as win32
from fastapi import APIRouter, FastAPI, Path
from win32com.client import Dispatch

from models.Worksheet import Worksheet

router = APIRouter(prefix="/api/v1", tags=["Test"])

excel_service_app = None
# semaphore = asyncio.Semaphore(1)
workbook_semaphore = asyncio.Semaphore(5)

@asynccontextmanager
async def lifespan(app: FastAPI):
    print("lifespan startup")
    global excel_service_app
    pythoncom.CoInitialize()
    excel_app = Dispatch("Excel.Application")
    yield
    print("lifespan shutdown")
    if excel_app:
        excel_app.Quit()
        excel_app = None

# router.lifespan_context = lifespan

# def do_thing():
#     excel = None
#     try:
#         print("starting")
#         pythoncom.CoInitialize()
#         excel = win32.Dispatch("Excel.Application")
#         workbook = excel.Workbooks.Open(r"C:\Users\dana\raptor-service\test.xlsx")
#         workbook.Close(SaveChanges=False)
#         print("ending")
#     except Exception as exception:
#         print(exception)
#     finally:
#         print("finally")
#         if excel: excel.Quit()
#         pythoncom.CoUninitialize()

def do_stuff():
    pythoncom.CoInitialize()
    try:
        workbook = excel_service_app.Workbooks.Open(r"C:\Users\dana\raptor-service\test.xlsx")
        workbook.Close(SaveChanges=False)
    finally:
        pythoncom.CoUninitialize()
    return {}

@router.get("/stuff")
async def get_stuff():
    async with workbook_semaphore:
        result = await asyncio.to_thread(do_stuff)
    return result

# @router.get("/thing")
# async def get_thing():
#     async with semaphore:
#         result = await asyncio.to_thread(do_thing)
#     return result

@router.get("/site/{site_id}/item/{item_id}/test", summary="Perform some example operations with Excel")
async def get_test(site_id: str = Path(..., description="Sharepoint Site ID"), item_id: str = Path(..., description="Sharepoint Item ID")):
    start_time = time.time()
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    try:
        workbook = excel.Workbooks.Open(r"C:\Users\dana\raptor-service\test.xlsx")
        worksheets = [sheet.Name for sheet in workbook.Worksheets]

        worksheet = Worksheet("M - Monthly", workbook)
        durable_ids = worksheet.get_durable_ids()
        income_statement_returns, durable_id_type = durable_ids.get_durable_id("incomeStatement.returns")
        # TODO: do something to income_statement_returns
        durable_ids.set_durable_id_values("incomeStatement.returns", income_statement_returns)

        # 'incomeStatement.returns'
        # TODO:
        #  - update some durableId values
        #  - put those values back into worksheet

        workbook.Application.CalculateFull()

        # TODO:
        #  - pull out durable_ids from output sheets
        #  - send data to DWH

        # workbook.Save()
        workbook.Close(SaveChanges=False)

        print(f"{time.time() - start_time}")
        return {"worksheets": worksheets}
    except Exception as exception:
        print(exception)
        return {"error": f"{type(exception)}"}
