#
# Copyright ©2024 Dana Basken
#

import requests
from msal import ConfidentialClientApplication
from config import settings

GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0/"
SCOPES = ["https://graph.microsoft.com/.default"]

app = ConfidentialClientApplication(
    settings.get("microsoft.client_id"),
    authority=f"https://login.microsoftonline.com/{settings.get("microsoft.tenant_id")}",
    client_credential=settings.get("microsoft.client_secret"),
)

result = app.acquire_token_for_client(scopes=SCOPES)
if "access_token" in result:
    access_token = result["access_token"]
else:
    raise Exception("failed to acquire token")

def request(method, endpoint):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
        "Prefer": "respond-async",
        "Consistency-Level": "eventual",
        "User-Agent": f"NONISV|bainbridgegrowth.com|Drivepoint/{settings.get("version")}"
    }
    url = f"{GRAPH_API_ENDPOINT}{endpoint}"
    # probably swap this out for something like tornado - https://www.tornadoweb.org/
    response = requests.request(method, url, headers=headers).json()
    return response
