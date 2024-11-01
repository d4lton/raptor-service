#
# Copyright ©2024 Dana Basken
#

from fastapi import FastAPI
from config import settings
from controllers import TestController

app = FastAPI(title="Drivepoint Raptor Service", version=settings.get("version"))

app.include_router(TestController.router)
