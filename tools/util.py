import json
from pypdf import PdfReader, PdfWriter
from pathlib import Path
import os
import requests
from openpyxl.worksheet.table import Table
from datetime import datetime
import pandas as pd
import subprocess
import glob
import sys

def show_welcome_banner():
    banner = rf"""
         ██████╗    ██████    █████╗ 
        ██╔═══██╗  ██        ██╔══██╗
        ██║   ██║  ██        ███████║
        ██║▄▄ ██║  ██        ██╔══██║
        ╚██████╔╝  ╚██████   ██║  ██║
         ╚══▀▀═╝    ╚═════   ╚═╝  ╚═╝
    ──────────────────────────────────────────────────────────
            🛠️  QCA Accounting Team Automation Script
            🏷️  Version: {get_config(["VERSION"])}
    ──────────────────────────────────────────────────────────
    """

    print(banner)
    print("🚀 Welcome! \n")

def show_menu():
    MENU = """
======================================================
                   MAIN MENU
======================================================
  1) AP - Remittance
  2) AR - Statements
  3) Report Preparation - Project Status
  4) Report Preparation - Bridge Report
  5) Shipping Monitor
  6) Daily Receiving Notification
  7) Merge PDF

  0) Exit
======================================================
"""
    print(MENU)

def get_runtime_dir() -> Path:
    """Get the runtime directory. In frozen mode, returns exe directory."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

def get_config(klist):
    """Get config value. In frozen mode, uses exe directory for persistence."""
    base_dir = get_runtime_dir()
    
    # In frozen mode (packaged exe), config should be in exe directory
    # This allows config to be editable and persistent
    if getattr(sys, "frozen", False):
        # Try exe directory/config/config.json first (for persistence)
        config_path = base_dir / "config" / "config.json"
        if not config_path.exists():
            # If not found, try to copy from bundled resource (one-time setup)
            try:
                if hasattr(sys, '_MEIPASS'):
                    # PyInstaller temporary folder
                    bundled_config = Path(sys._MEIPASS) / "config" / "config.json"
                    if bundled_config.exists():
                        # Create config directory in exe folder
                        (base_dir / "config").mkdir(exist_ok=True)
                        import shutil
                        shutil.copy2(bundled_config, config_path)
            except Exception:
                pass
    else:
        # Development mode: use standard paths
        config_path = (base_dir / "config.json"
            if (base_dir / "config.json").exists()
            else base_dir.parent / "config" / "config.json")
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        node = config
        for key in klist:
            if key not in node: 
                node[key] = ""
            node = node[key]

        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
            
        return node
    
    except Exception as e:
        print(f"⚠️ Get config failed")
        return None

def set_config(key, value):
    """Set config value. In frozen mode, uses exe directory for persistence."""
    base_dir = get_runtime_dir()
    
    # In frozen mode (packaged exe), config should be in exe directory
    if getattr(sys, "frozen", False):
        config_path = base_dir / "config" / "config.json"
        # Ensure config directory exists
        config_path.parent.mkdir(exist_ok=True)
        # If config doesn't exist, try to copy from bundled resource
        if not config_path.exists():
            try:
                if hasattr(sys, '_MEIPASS'):
                    bundled_config = Path(sys._MEIPASS) / "config" / "config.json"
                    if bundled_config.exists():
                        import shutil
                        shutil.copy2(bundled_config, config_path)
            except Exception:
                pass
    else:
        # Development mode: use standard paths
        config_path = (base_dir / "config.json"
            if (base_dir / "config.json").exists()
            else base_dir.parent / "config" / "config.json")
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        
        config[key] = value

        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
            f.flush()

        print(f"💾 {key} been updated to {value}. ")

    except Exception as e:
        print(f"⚠️ {key} updated failed: ", e)

def check_login():
    try: 
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/visionservices.asmx/GetIAccessConfiguration"
        payload = {"sessionID": get_config(["TOKEN"])}
        response = requests.post(url, headers = set_headers(), json = payload)
        cookies = get_config(["COOKIES"])
        if response.status_code == 200 and "ASP.NET_SessionId" in cookies:
            data = response.json()
            print("✅ Login Success, User: " + data["d"]["UserInfo"]["EMail"])
            print()
            return True
        else: 
            return False
    except:
        set_config("TOKEN", "")
        set_config("WWWBEARER", "")
        set_config("COOKIES", "")
        print("❌ Error in function check_login()")
        return False

def merge_amazon_invoices():
    path = input("Please input the folder path: ")
    input_folder = Path(path)  
    output_file = input_folder / "combined_output.pdf"

    combined_writer = PdfWriter()

    for pdf_file in input_folder.glob("*.pdf"):
        try:
            reader = PdfReader(str(pdf_file))
            if len(reader.pages) >= 1:
                combined_writer.add_page(reader.pages[0])
            if len(reader.pages) == 2:
                combined_writer.add_page(reader.pages[1])
            if len(reader.pages) >= 3:
                combined_writer.add_page(reader.pages[2])
            print(f"✅ Done: {pdf_file.name}")
        except Exception as e:
            print(f"⚠️ Error {pdf_file.name}: {e}")

    with open(output_file, "wb") as f:
        combined_writer.write(f)

    print(f"\n🎉 Success: {output_file}")

def merge_pdfs():
    folder_path = input("📂 Please input folder path of the pdf files: ").strip()
    output_filename = "merged_output.pdf"

    if not os.path.isdir(folder_path):
        print("❌ Invalid folder path")

    pdf_writer = PdfWriter()

    pdf_files = [
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.lower().endswith('.pdf')
    ]

    # Sort by created time
    pdf_files.sort(key=lambda f: os.path.getctime(f))

    if not pdf_files:
        print("❌ No pdf files")
        return

    for pdf_path in pdf_files:
        try:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                pdf_writer.add_page(page)
            print(f"✅ Add: {os.path.basename(pdf_path)}")
        except Exception as e:
            print(f"⚠️ Skip {pdf_path}: {e}")

    output_path = os.path.join(folder_path, output_filename)
    with open(output_path, "wb") as out_file:
        pdf_writer.write(out_file)

    print(f"\n🎉 Success the merged pdf file: {output_path}")

def set_headers():
    WWWBEARER = get_config(["WWWBEARER"])
    TOKEN = get_config(["TOKEN"])
    COOKIES = get_config(["COOKIES"])
    headers = {
        "accept": "application/json, text/javascript, */*; q=0.01",
        "Content-Type": "application/json; charset=UTF-8",
        "www-bearer": WWWBEARER,
        "Token": TOKEN,
        "Cookie": COOKIES
    }
    return headers

def init_workdir():
    current_path = os.getcwd()
    pattern = os.path.join(".", "job*.xlsx")

    for file_path in glob.glob(pattern):
        try:
            os.remove(file_path)
        except Exception as e:
            print(f"Failed to delete {file_path}: {e}")
    
    before, _, after = current_path.partition("OneDrive - QCA Systems Ltd")
    if not after:
        print(f'⚠️ OneDrive - QCA Systems Ltd not found in current path: {current_path}, the power automate can not been used.')
        print()
        set_config("ONEDRIVEDIR", str(current_path))
        return

    ONEDRIVEDIR = before + "OneDrive - QCA Systems Ltd"
    WORKDIR =  after.lstrip("\\")

    set_config("ONEDRIVEDIR", ONEDRIVEDIR)
    set_config("WORKDIR", WORKDIR)


def assamble_projects(projects):
    searchOptions = [{"name":"Status","value":"[IS_EMPTY]","type":"dropdown","seq":1,"tableName":"PR","opp":"!=","condition":"and","searchLevel":1,"valueDescription":""}]
    for project in projects:
        searchOptions.append({"name":"selectedResultIds","value":project,"type":"wbs1","seq":2,"searchLevel":0,"valueDescription":"Monthly Maintenance"})
    
    return searchOptions

def check_folder(folderName):
    if not os.path.exists(folderName):
        os.makedirs(folderName)
        print(f"📁 Folder created: {folderName}")

def clear_folder(folderName):
    onedrivedir = get_config(["ONEDRIVEDIR"])
    workdir = get_config(["WORKDIR"])

    folder_path = os.path.join(onedrivedir, workdir, folderName)
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path):
            os.remove(file_path)

def get_vendor_email(clientID):
    headers = set_headers()
    try:
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Client/"+ clientID +"/Address/"
        response = requests.get(url, headers=headers  )
        data = response.json()
        email = ""
        if (len(data) == 1):
            email = data[0]["Email"]
        else:
            for d in data:
                if d["Email"] != "" and (not d["Email"] in email): 
                    if "AP Automation" in d['Address'] or "AP Mailing" in d['Address']: 
                        email = d["Email"]+";"
                        break
                    else:
                        email += d["Email"]+";"
        return email
    except Exception as e:
        print("❌ Error in function get_vendor_email()")

def get_clientID(clientName):
    try:
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/visionservices.asmx/GetLookupHash"
        headers = set_headers()
        payload = {
            "sessionID": get_config(["TOKEN"]),
            "hash":{
                "filter": clientName,
                "lookuptype": "clientvendor","vendorOnly":"N"
            },
            "page": 1,
            "pagesize": 20,
            "order": "name"
        }
        response = requests.post(url, headers=headers, json=payload)
        records = response.json()
        for record in records['d']:
            if record['IsClient'] == "Y":
                return record['Key']
    except Exception as e:
        print("❌ Error in function get_clientID() with input: " + clientName)

def save_excel(wb, records):
    try:
        ws = wb.active
        table_range = f"A1:I{records}"
        if "Table1" in ws.tables:
            del ws.tables["Table1"]
        tab = Table(displayName = "Table1", ref = table_range)
        ws.add_table(tab)
        timestamp = datetime.now().strftime("%Y%m%d")
        excelName = os.path.join( f"Job_{timestamp}.xlsx")
        wb.save(excelName)

        return excelName
    except Exception as e:
        print("❌Error in function save_excel()")

def download_with_progress(url, save_path, latest_version):
    import requests
    r = requests.get(url, stream=True)
    total = int(r.headers.get('content-length', 0))
    downloaded = 0
    chunk_size = 8192

    print("\nDownloading update...\n")  # 只终端显示

    with open(save_path, "wb") as f:
        for chunk in r.iter_content(chunk_size):
            if chunk:
                f.write(chunk)
                downloaded += len(chunk)

                percent = downloaded / total * 100 if total else 0
                bar = "█" * int(percent / 2)
                space = " " * (50 - len(bar))

                print(f"\r[{bar}{space}] {percent:6.2f}%  ({downloaded/1024/1024:.2f} MB / {total/1024/1024:.2f} MB)", end="", flush=True)

    print(f"Downloaded {downloaded/1024/1024:.2f} MB / {total/1024/1024:.2f} MB")
    print("Download complete!")
    set_config("VERSION", latest_version)


def check_update():
    GITHUB_REPO = "Coding-Lou/Vantagepoint_Automation"
    EXE_NAME = "start.exe" 
    VERSION = get_config(["VERSION"])

    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        r = requests.get(url, timeout=10)
        data = r.json()

        latest_version = data["tag_name"].replace("v_", "")
        asset = data["assets"][0]
        download_url = asset["browser_download_url"]

        if latest_version == VERSION:
            print("Already latest version.")
            return

        print(f"New version found: {latest_version}. Downloading update...")

        temp_new_exe = os.path.join(os.getenv("TEMP"), EXE_NAME)

        download_with_progress(download_url, temp_new_exe, latest_version)
        
        updater = os.path.join(os.path.dirname(sys.argv[0]), "updater.exe")
        subprocess.Popen([updater, sys.argv[0], temp_new_exe])

        print("Update started. Exiting old program...")
        sys.exit(0)
    
    except Exception as e:
        print(f"Failed to check/update version: {e}")
        return

def change_period(period):
    headers = set_headers()
    url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/PeriodSetup/ActivePeriod/" + period
    response = requests.put(url, headers = headers) 

def csv_to_xlsx(csv_path, output_file, sheet_name, need_skip, left, right):
    if need_skip:
        df = pd.read_csv(csv_path, skiprows=3)
    else:
        df = pd.read_csv(csv_path)

    if isinstance(left, int) and isinstance(right, int):
        df = df.iloc[:, left:right+1] 
    else:
        df = df.loc[:, left:right] 

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Copied to the {output_file} / {sheet_name}")

def cleanup_projectID(project_id: str) -> str:
    if not project_id:
        return project_id
    replace_map = {
        "/": "[_$2F_]",
        "&": "[_$26_]",
    }
    for k, v in replace_map.items():
        project_id = project_id.replace(k, v)
    return project_id

def cleanup_projectName(projName: str) -> str:
    if not projName:
        return projName
    replace_map = {
        " ": "%20",
        "/": "%2f",
        "&": "%26",
    }
    for k, v in replace_map.items():
        projName = projName.replace(k, v)
    return projName