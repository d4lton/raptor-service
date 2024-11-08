#
# Copyright ©2024 Dana Basken
#

import urllib.parse
from tenacity import retry, stop_after_attempt, wait_random_exponential
from services import MicrosoftGraphQLClient

def get_groups():
    return MicrosoftGraphQLClient.request("GET", "/groups").get("value")

@retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
def get_drive_item(group_id: str, item_id: str):
    return MicrosoftGraphQLClient.request("GET", f"/groups/{group_id}/drive/items/{item_id}")

@retry(wait=wait_random_exponential(multiplier=1, max=10), stop=stop_after_attempt(5))
def get_folder_items(group_id: str, item_id: str):
    group_id = urllib.parse.quote(group_id)
    item_id = urllib.parse.quote(item_id)
    items = MicrosoftGraphQLClient.request("GET", f"/groups/{group_id}/drive/items/{item_id}/children?$expand=listItem($expand=fields)")
    return items.get("value")
