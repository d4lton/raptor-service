#
# Copyright ©2024 Dana Basken
#

from fastapi import APIRouter
from services import SharepointGroupService

router = APIRouter(prefix="/api/v1/group", tags=["Groups"])

@router.get("/", summary="Get groups")
async def get_groups():
    return await SharepointGroupService.get_groups()
