#
# Copyright ©2024 Dana Basken
#

from contextlib import asynccontextmanager
from fastapi import FastAPI
from config import settings
from services.ExcelService import shutdown_excel_service, start_excel_service
from controllers import ChatGPTController, TestController, DriveItemController, GroupController

# @asynccontextmanager
# async def lifespan(app: FastAPI):
#     print("lifespan startup")
#     start_excel_service()
#     yield
#     print("lifespan shutdown")
#     shutdown_excel_service()

app = FastAPI(title="Drivepoint Raptor Service", version=settings.get("version"))

app.include_router(ChatGPTController.router)
app.include_router(DriveItemController.router)
app.include_router(GroupController.router)
app.include_router(TestController.router)

@app.on_event("startup")
def startup_event():
    start_excel_service()

@app.on_event("shutdown")
def shutdown_event():
    shutdown_excel_service()