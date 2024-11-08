#
# Copyright ©2024 Dana Basken
#

from fastapi import APIRouter
from excel_pool.ExcelPoolTask import ExcelPoolTask
from excel_pool.ExcelPool import ExcelPool

router = APIRouter(prefix="/api/v1", tags=["Test"])

@router.post("/stuff/")
async def post_stuff(excel_pool_task: ExcelPoolTask):
    farm = ExcelPool()
    id = farm.add_task(excel_pool_task)
    return {"id": id, "excel_pool_task": excel_pool_task}
