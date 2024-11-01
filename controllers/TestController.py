#
# Copyright ©2024 Dana Basken
#

import time
import win32com.client as win32
from fastapi import APIRouter, Path
from models.Worksheet import Worksheet
from services import SharepointGroupService

router = APIRouter(prefix="/api/v1", tags=["Test"])

@router.get("/group", summary="Get groups")
async def get_groups():
    return await SharepointGroupService.get_groups()

@router.get("/group/{group_id}/item/{item_id}", summary="Get a drive item")
async def get_drive_item(group_id, item_id):
    return await SharepointGroupService.get_drive_item(group_id, item_id)

@router.get("/group/{group_id}/item/{item_id}/list", summary="Get drive items under a folder")
async def get_drive_items(group_id, item_id):
    return await SharepointGroupService.get_folder_items(group_id, item_id)

@router.get("/site/{site_id}/item/{item_id}/test", summary="Perform some example operations with Excel")
async def get_test(
        site_id: str = Path(..., description="Sharepoint Site ID"),
        item_id: str = Path(..., description="Sharepoint Item ID")):
    start_time = time.time()
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    try:
        print(site_id, item_id)
        workbook = excel.Workbooks.Open(r"C:\Users\dana\raptor-service\test.xlsx")
        worksheets = [sheet.Name for sheet in workbook.Worksheets]

        worksheet = Worksheet("M - Monthly", workbook)
        durable_ids = worksheet.get_durable_ids()

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

        excel.Quit()
        print(f"{time.time() - start_time}")
        return {"worksheets": worksheets}
    except Exception as exception:
        print(exception)
        excel.Quit()
        return {"error": f"{type(exception)}"}
