#
# Copyright ©2024 Dana Basken
#

from services import Microsoft
import urllib.parse

async def get_groups():
    return Microsoft.request("GET", "/groups").get("value")

async def get_drive_item(group_id: str, item_id: str):
    return Microsoft.request("GET", f"/groups/{group_id}/drive/items/{item_id}")

async def get_folder_items(group_id: str, item_id: str):
    group_id = urllib.parse.quote(group_id)
    item_id = urllib.parse.quote(item_id)
    items = Microsoft.request("GET", f"/groups/{group_id}/drive/items/{item_id}/children?$expand=listItem($expand=fields)")
    return items.get("value")
