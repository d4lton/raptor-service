#
# Copyright ©2024 Dana Basken
#

from fastapi import APIRouter, Path
from services import SharepointGroupService

router = APIRouter(prefix="/api/v1/group/{group_id}/item/{item_id}", tags=["Drive Items"])

@router.get("/", summary="Get a drive item")
async def read(group_id: str = Path(..., description="Sharepoint Group ID"), item_id: str = Path(..., description="Sharepoint Item ID")):
    return await SharepointGroupService.get_drive_item(group_id, item_id)

@router.get("/list", summary="Get drive items under a folder drive item")
async def list(group_id: str = Path(..., description="Sharepoint Group ID"), item_id: str = Path(..., description="Sharepoint Item ID")):
    return await SharepointGroupService.get_folder_items(group_id, item_id)
