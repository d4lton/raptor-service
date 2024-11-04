#
# Copyright ©2024 Dana Basken
#

import time
import win32com.client as win32
from fastapi import APIRouter, Path
from excel_pool.ExcelPoolTask import ExcelPoolTask
from excel_pool.ExcelPool import ExcelPool
from models.Worksheet import Worksheet

router = APIRouter(prefix="/api/v1", tags=["Test"])

@router.post("/stuff/")
async def post_stuff(excel_pool_task: ExcelPoolTask):
    farm = ExcelPool()
    id = farm.add_task(excel_pool_task)
    return {"id": id, "excel_pool_task": excel_pool_task}

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
