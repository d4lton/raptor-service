#
# Copyright ©2024 Dana Basken
#

from fastapi import FastAPI
from config import settings
from ExcelFarm import ExcelFarm
from controllers import TestController, DriveItemController, GroupController

farm = ExcelFarm(10)

app = FastAPI(title="Drivepoint Raptor Service", version=settings.get("version"))

app.include_router(DriveItemController.router)
app.include_router(GroupController.router)
app.include_router(TestController.router)

farm.add_task(r"C:\Users\dana\raptor-service\test.xlsx")
farm.add_task(r"C:\Users\dana\raptor-service\test2.xlsx")