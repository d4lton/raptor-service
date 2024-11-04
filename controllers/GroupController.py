#
# Copyright ©2024 Dana Basken
#

from fastapi import APIRouter
from services import SharepointGroupService

router = APIRouter(prefix="/api/v1/group", tags=["Groups"])

@router.get("/", summary="Get groups")
def get_groups():
    return SharepointGroupService.get_groups()
