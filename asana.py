import json
import re
import pdfplumber
import requests
import os
import tools.util as util
from domain.adapters.anixter import AnixterParser

# 你的 Asana Personal Access Token
ACCESS_TOKEN = ""
DOWNLOAD_FOLDER = "vendor_invoice"
util.check_folder(DOWNLOAD_FOLDER)

pattern = { 'WESTBURNE' : r"Invoice\s+#\d{7}\s+from\s+Rexel" }

headers = {
    "Authorization": f"Bearer {ACCESS_TOKEN}"
}


user_resp = requests.get("https://app.asana.com/api/1.0/users/me", headers=headers)
user_resp.raise_for_status()
user_data = user_resp.json()['data']
user_gid = user_data['gid']
workspace_gid = user_data['workspaces'][0]['gid']  # 默认取第一个 workspace

# 2. 获取分配给当前用户的未完成任务
params = {
    "assignee": user_gid,
    "workspace": workspace_gid,    # 必填
    "completed_since": "now",
    "opt_fields": "name"
}

tasks_resp = requests.get("https://app.asana.com/api/1.0/tasks", headers=headers, params=params)
tasks_resp.raise_for_status()

tasks = tasks_resp.json().get('data', [])
print("Incompleted Tasks：")
for task in tasks:
    if not task["name"].startswith("Acct No. AXC"):
        continue
    if "Your Invoice From Anixter is Attached" not in task["name"]:
        continue
    task_gid = task["gid"]
    
    attachments_url = f"https://app.asana.com/api/1.0/tasks/{task_gid}/attachments"
    resp = requests.get(attachments_url, headers=headers)
    resp.raise_for_status()
    attachments = resp.json().get("data", [])

    for attach in attachments:
        attach_name = attach["name"]
        if (attach_name.split('.')[-1].lower() != 'pdf'):
            continue
        attach_gid = attach['gid']

        attach_info_url = f"https://app.asana.com/api/1.0/attachments/{attach_gid}"
        attach_resp = requests.get(attach_info_url, headers=headers)
        attach_resp.raise_for_status()
        download_url = attach_resp.json()["data"]["download_url"]

        r = requests.get(download_url, headers=headers, stream=True)
        r.raise_for_status()

        filepath = os.path.join(DOWNLOAD_FOLDER, attach_name)
        with open(filepath, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)

        print(f"Saved {attach_name} to {DOWNLOAD_FOLDER}")

        parser = AnixterParser()

        invoice = parser.parse(os.path.join(DOWNLOAD_FOLDER, attach_name))
        print(json.dumps(invoice.to_dict(),indent=4,ensure_ascii=False))