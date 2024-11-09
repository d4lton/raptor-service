#
# Copyright ©2024 Dana Basken
#

import signal
from fastapi import FastAPI
from config import settings
from excel_pool.ExcelPool import ExcelPool
from controllers import TestController, DriveItemController, GroupController
from utilities import Logging

Logging.colorize()

pool = ExcelPool() # initialize ExcelPool singleton to "warm up" Excel processes

app = FastAPI(title="Drivepoint Raptor Service", version=settings.get("version"))

app.include_router(DriveItemController.router)
app.include_router(GroupController.router)
app.include_router(TestController.router)

signal.signal(signal.SIGINT, pool.shutdown)
signal.signal(signal.SIGTERM, pool.shutdown)
