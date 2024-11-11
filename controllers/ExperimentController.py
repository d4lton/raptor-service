#
# Copyright ©2024 Dana Basken
#

from typing import List
from fastapi import APIRouter
from excel_pool.ExcelPoolTask import ExcelPoolTask
from excel_pool.ExcelPool import ExcelPool

from pydantic import BaseModel

class Treatment(BaseModel):
    durable_id: str

class TreatmentRequest(BaseModel):
    group_id: str
    item_id: str
    treatments: List[Treatment]

router = APIRouter(prefix="/api/v1/group/{group_id}/item/{item_id}/experiment", tags=["Experiments"])

@router.post("/treatments")
async def post_stuff(request: TreatmentRequest):
    excel_pool_task = ExcelPoolTask(**request.__dict__, type="demo")
    farm = ExcelPool()

    # TODO:
    #  - generate combinations
    #  - add task for each combination
    #  - poll farm for completion of each task
    #  - respond with results

    id = farm.add_task(excel_pool_task)
    task_status = await farm.wait_for_task_completion(id)
    return {"id": id, "task_status": task_status}
