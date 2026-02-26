import re
import pdfplumber
import requests
import os
import tools.util as util

# 你的 Asana Personal Access Token
ACCESS_TOKEN = "2/1210190277961957/1211118508618456:b77803d1519340d597e0befc7d423c10"
DOWNLOAD_FOLDER = "vendor_invoice"
util.check_folder(DOWNLOAD_FOLDER)

pattern = { 'WESTBURNE' : r"Invoice\s+#\d{7}\s+from\s+Rexel" }

headers = {
    "Authorization": f"Bearer {ACCESS_TOKEN}"
}

# 1. 获取当前用户信息
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
print("未完成任务列表：")
for task in tasks:
    if not re.search(pattern["WESTBURNE"], task['name']):
        continue
    task_gid = task["gid"]
    # 获取任务的附件列表
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

        with pdfplumber.open(os.path.join(DOWNLOAD_FOLDER, attach_name)) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables()
                for table_index, table in enumerate(tables, start=1):
                    fixed_table = []  # 存放拆分后的新表格
                    header = table[0]
                    fixed_table.append(header)
                    print("-----------------------------")
                    for row in table[1:]:
                        # 找到包含换行符的单元格
                        max_splits = max([len(str(cell).split('\n')) if cell else 1 for cell in row])

                        # 为每个换行分裂成多行
                        new_rows = []
                        for i in range(max_splits):
                            new_row = []
                            for cell in row:
                                parts = str(cell).split('\n') if cell else ['']
                                # 如果当前行没有对应部分，用空字符串填充
                                new_row.append(parts[i] if i < len(parts) else '')
                            new_rows.append(new_row)

                        fixed_table.extend(new_rows)

                    # 输出修正后的表格
                    for r in fixed_table:
                        print(r)