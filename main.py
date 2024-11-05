#
# Copyright ©2024 Dana Basken
#

from fastapi import FastAPI
from config import settings
from excel_pool.ExcelPool import ExcelPool
from controllers import TestController, DriveItemController, GroupController
from utilities import Logging

Logging.colorize()

ExcelPool() # initialize ExcelPool singleton

app = FastAPI(title="Drivepoint Raptor Service", version=settings.get("version"))

app.include_router(DriveItemController.router)
app.include_router(GroupController.router)
app.include_router(TestController.router)
