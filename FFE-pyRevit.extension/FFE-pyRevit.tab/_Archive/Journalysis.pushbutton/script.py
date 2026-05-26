#! python3
__title__ = "Jurnalysis"
__author__ = "Kyle Guggenheim"
__doc__ = "Lists Model Sync Events in ACC"

import os
import requests

CLIENT_ID = "6sYYxvNS2Xun1B7rYgAkzQGoAqxWaHPQ27sQGBGSW7tJIFv0"
CLIENT_SECRET = "ziXoOwBCl9CIblrRXbM73g7Hrm8FtJzTe10IrXymp6ZDGNTPGquudZdBS1lAFFmD"
PROJECT_ID = "c1b3fecf-d650-42d9-b63d-ab66821d1e83"
FOLDER_ID = "<YOUR_FOLDER_ID>"
WEBHOOK_URL = "https://acc-sync-logger.kylegug.workers.dev"

# 1. Get token
resp = requests.post(
    "https://developer.api.autodesk.com/authentication/v1/authenticate",
    data={
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials",
        "scope": "data:read webhook:read webhook:write"
    }
)
resp.raise_for_status()
token = resp.json()["access_token"]

# 2. Register webhook
hook_payload = {
    "callbackUrl": WEBHOOK_URL,
    "scope": {
        "folder": FOLDER_ID
    },
    "event": "adsk.c4r:model.sync"
}
r = requests.post(
    f"https://developer.api.autodesk.com/webhooks/v1/systems/c4r/events/adsk.c4r:model.sync/hooks",
    json=hook_payload,
    headers={"Authorization": f"Bearer {token}"}
)
print(r.status_code, r.text)
