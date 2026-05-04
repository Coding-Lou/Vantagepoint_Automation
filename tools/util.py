import json
import shutil
import time
import threading
import tempfile
from pypdf import PdfReader, PdfWriter
from pathlib import Path
import os
import requests
from openpyxl.worksheet.table import Table
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
import subprocess
import glob
import sys
import win32com.client
import pythoncom
import winreg

# Shared lock — prevents concurrent reads from seeing a half-written config file.
# config_manager.py imports this same lock so all config I/O is serialized.
_config_lock = threading.Lock()


def _atomic_write_config(config_path: Path, data: dict) -> None:
    """Write config atomically: write to a temp file then os.replace."""
    dir_ = config_path.parent
    fd, tmp_path = tempfile.mkstemp(dir=dir_, suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_path, config_path)
    except Exception:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise


def _read_config_with_retry(config_path: Path, retries: int = 3, delay: float = 0.05) -> dict:
    """Read and parse config.json, retrying on empty/invalid content."""
    for attempt in range(retries):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                content = f.read()
            if content.strip():
                return json.loads(content)
        except (json.JSONDecodeError, OSError):
            pass
        if attempt < retries - 1:
            time.sleep(delay)
    raise ValueError(f"Invalid JSON format in config file: Expecting value : line 1 column 1 (char 0)")


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
    """Get config value. Config file is in the main directory (same level as main.py or exe)."""
    base_dir = get_runtime_dir()
    
    # Config is in the main directory (same level as main.py or exe)
    config_path = base_dir / "config.json"
    
    # In frozen mode, if config doesn't exist, try to copy from bundled resource
    if getattr(sys, "frozen", False) and not config_path.exists():
        try:
            if hasattr(sys, '_MEIPASS'):
                # PyInstaller temporary folder
                bundled_config = Path(sys._MEIPASS) / "config.json"
                if bundled_config.exists():
                    import shutil
                    shutil.copy2(bundled_config, config_path)
        except Exception:
            pass
    try:
        with _config_lock:
            config = _read_config_with_retry(config_path)
            node = config
            for key in klist:
                if key not in node:
                    node[key] = ""
                node = node[key]
            _atomic_write_config(config_path, config)
        return node
    except Exception as e:
        print(f"⚠️ Get config failed")
        return None

def set_config(key, value):
    """Set config value. Config file is in the main directory (same level as main.py or exe)."""
    base_dir = get_runtime_dir()
    
    # Config is in the main directory (same level as main.py or exe)
    config_path = base_dir / "config.json"
    
    # In frozen mode, if config doesn't exist, try to copy from bundled resource
    if getattr(sys, "frozen", False) and not config_path.exists():
        try:
            if hasattr(sys, '_MEIPASS'):
                bundled_config = Path(sys._MEIPASS) / "config.json"
                if bundled_config.exists():
                    import shutil
                    shutil.copy2(bundled_config, config_path)
        except Exception:
            pass
    try:
        with _config_lock:
            config = _read_config_with_retry(config_path)
            config[key] = value
            _atomic_write_config(config_path, config)
        print(f"💾 {key} been updated to {value}. ")
    except Exception as e:
        print(f"⚠️ {key} updated failed: ", e)

def check_login():
    from pathlib import Path
    import json
    import time
    DEBUG_LOG_PATH = Path(r"c:\cursor\.cursor\debug.log")
    def _agent_log(hypothesis_id: str, location: str, message: str, data: dict = None):
        payload = {
            "sessionId": "debug-session",
            "runId": "pre-fix",
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": data or {},
            "timestamp": int(time.time() * 1000),
        }
        try:
            DEBUG_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
            with DEBUG_LOG_PATH.open("a", encoding="utf-8") as f:
                f.write(json.dumps(payload, ensure_ascii=False) + "\n")
        except Exception:
            pass
    #region agent log
    _agent_log("H2", "util.check_login", "check_login() entry", {})
    #endregion
    try: 
        token = get_config(["TOKEN"])
        #region agent log
        _agent_log("H2", "util.check_login", "check_login() before request", {
            "has_token": bool(token),
            "token_length": len(token) if token else 0,
        })
        #endregion
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/visionservices.asmx/GetIAccessConfiguration"
        payload = {"sessionID": token}
        headers = set_headers()
        #region agent log
        _agent_log("H2", "util.check_login", "check_login() headers from set_headers", {
            "has_token_in_headers": bool(headers.get("Token")),
            "has_cookie_in_headers": bool(headers.get("Cookie")),
        })
        #endregion
        response = requests.post(url, headers = headers, json = payload)
        #region agent log
        _agent_log("H2", "util.check_login", "check_login() after request", {
            "status_code": response.status_code,
            "response_length": len(response.text) if response.text else 0,
        })
        #endregion
        cookies = get_config(["COOKIES"])
        #region agent log
        _agent_log("H2", "util.check_login", "check_login() cookie check", {
            "has_cookies": bool(cookies),
            "has_aspnet": "ASP.NET_SessionId" in (cookies or ""),
        })
        #endregion
        if response.status_code == 200 and "ASP.NET_SessionId" in cookies:
            data = response.json()
            #region agent log
            _agent_log("H2", "util.check_login", "check_login() success", {
                "user_email": data.get("d", {}).get("UserInfo", {}).get("EMail") if isinstance(data, dict) else None,
            })
            #endregion
            #print("✅ Login Success, User: " + data["d"]["UserInfo"]["EMail"])
            #print()
            return data["d"]["UserInfo"]["EMail"]
        else: 
            #region agent log
            _agent_log("H2", "util.check_login", "check_login() failed", {
                "status_code": response.status_code,
                "has_aspnet": "ASP.NET_SessionId" in (cookies or ""),
            })
            #endregion
            return None
    except Exception as e:
        #region agent log
        _agent_log("H2", "util.check_login", "check_login() exception", {
            "error": str(e),
        })
        #endregion
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
    from pathlib import Path
    import json
    import time
    DEBUG_LOG_PATH = Path(r"c:\cursor\.cursor\debug.log")
    def _agent_log(hypothesis_id: str, location: str, message: str, data: dict = None):
        payload = {
            "sessionId": "debug-session",
            "runId": "pre-fix",
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": data or {},
            "timestamp": int(time.time() * 1000),
        }
        try:
            DEBUG_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
            with DEBUG_LOG_PATH.open("a", encoding="utf-8") as f:
                f.write(json.dumps(payload, ensure_ascii=False) + "\n")
        except Exception:
            pass
    #region agent log
    _agent_log("H3", "util.set_headers", "set_headers() entry", {})
    #endregion
    WWWBEARER = get_config(["WWWBEARER"])
    TOKEN = get_config(["TOKEN"])
    COOKIES = get_config(["COOKIES"])
    #region agent log
    _agent_log("H3", "util.set_headers", "set_headers() after get_config", {
        "has_wwwbearer": bool(WWWBEARER),
        "has_token": bool(TOKEN),
        "has_cookies": bool(COOKIES),
        "token_length": len(TOKEN) if TOKEN else 0,
        "cookie_length": len(COOKIES) if COOKIES else 0,
        "cookie_has_aspnet": "ASP.NET_SessionId" in (COOKIES or ""),
    })
    #endregion
    headers = {
        "accept": "application/json, text/javascript, */*; q=0.01",
        "Content-Type": "application/json; charset=UTF-8",
        "www-bearer": WWWBEARER,
        "Token": TOKEN,
        "Cookie": COOKIES
    }
    #region agent log
    _agent_log("H3", "util.set_headers", "set_headers() return", {
        "headers_keys": list(headers.keys()),
    })
    #endregion
    return headers


def init_workdir():
    # 1. Locate OneDrive Root
    onedrive_root_str = os.getenv("OneDriveCommercial") or os.getenv("OneDrive")
    
    if not onedrive_root_str:
        # Fallback to manual check if environment variables are missing
        current_path = Path.cwd()
        target_name = "OneDrive - QCA Systems Ltd"
        if target_name in current_path.parts:
            idx = current_path.parts.index(target_name)
            onedrive_root = Path(*current_path.parts[:idx + 1])
        else:
            print("❌ OneDrive not found.")
            return
    else:
        onedrive_root = Path(onedrive_root_str)

    # 2. Build the path to the 'automation' folder inside 'Documents'
    # Pathlib automatically handles the slashes for your OS
    automation_path = onedrive_root / "Documents" / "automation"

    try:
        # 3. Create the folder if it doesn't exist
        # parents=True ensures 'Documents' is created if it's somehow missing
        automation_path.mkdir(parents=True, exist_ok=True)
        
        # 4. Set configurations
        # relative_to() provides the string "Documents/automation"
        work_dir_relative = automation_path.relative_to(onedrive_root)
        
        set_config("ONEDRIVEDIR", str(onedrive_root))
        set_config("WORKDIR", str(work_dir_relative))
        
        print(f"✅ Target Workspace: {automation_path}")
        
    except Exception as e:
        print(f"❌ Could not create directory: {e}")


def assamble_projects(projects):
    searchOptions = [{"name":"Status","value":"[IS_EMPTY]","type":"dropdown","seq":1,"tableName":"PR","opp":"!=","condition":"and","searchLevel":1,"valueDescription":""}]
    for project in projects:
        searchOptions.append({"name":"selectedResultIds","value":project,"type":"wbs1","seq":2,"searchLevel":0,"valueDescription":""})
    
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

def save_excel(wb, records, folder_path=None):
    try:
        ws = wb.active
        table_range = f"A1:I{records}"
        if "Table1" in ws.tables:
            del ws.tables["Table1"]
        tab = Table(displayName = "Table1", ref = table_range)
        ws.add_table(tab)
        timestamp = datetime.now().strftime("%Y%m%d")
        if folder_path is None:
            folder_path = os.path.join(get_config(["ONEDRIVEDIR"]), get_config(["WORKDIR"]))
        excelName = os.path.join(folder_path, f"Job_{timestamp}.xlsx")
        wb.save(excelName)

        return excelName
    except Exception as e:
        print(f"❌ Error in function save_excel(): {e}")

def download_with_progress(url, save_path, latest_version):
    import requests
    r = requests.get(url, stream=True)
    total = int(r.headers.get('content-length', 0))
    downloaded = 0
    chunk_size = 8192

    print("\nDownloading update...\n") 

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

def csv_to_xlsx(csv_path, output_file, sheet_name, need_skip, left = None, right = None):
    if need_skip:
        df = pd.read_csv(csv_path, skiprows=3)
    else:
        df = pd.read_csv(csv_path)

    if left is not None and right is not None:
        df = df.iloc[:, left:right+1] 
    elif left is not None:
        df = df.iloc[:, left:]
    elif right is not None:
        df = df.iloc[:, :right+1] 

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


def excel_full_copy(inputFile, inputSheet, targetFile, targetSheet, onlyValue,
                    targetCell='A1', refreshAll=True,
                    excel_instance=None, target_wb_instance=None):

    _com_initialized = False
    if excel_instance is None:
        pythoncom.CoInitialize()
        _com_initialized = True
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AlertBeforeOverwriting = False
    else:
        excel = excel_instance

    inputFile = os.path.abspath(inputFile)

    CALLEE_BUSY     = -2147418111
    SERVER_DISC     = -2147220995
    MAX_RETRIES     = 20
    RETRY_SLEEP     = 3

    def robust_call(func, *args, **kwargs):
        for attempt in range(MAX_RETRIES):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                hresult = getattr(e, 'hresult', None)
                if hresult == SERVER_DISC:
                    raise                        # unrecoverable
                if hresult == CALLEE_BUSY:
                    if attempt < MAX_RETRIES - 1:
                        time.sleep(RETRY_SLEEP)
                        continue
                    raise Exception("Excel still busy after max retries") from e
                raise                            # any other error → propagate immediately
        raise Exception("Excel Busy: max retries exceeded")

    def safe_close_wb(wb, save: bool):
        """Close a workbook, retrying on CALLEE_BUSY, swallowing everything else."""
        try:
            robust_call(wb.Close, save)
        except Exception:
            pass

    wb_input  = None
    wb_target = None

    try:
        wb_input = robust_call(excel.Workbooks.Open, inputFile)
        if inputFile.lower().endswith(".csv"):
            ws_input = robust_call(lambda: wb_input.Worksheets(1))
        else:
            ws_input = robust_call(lambda: wb_input.Worksheets(inputSheet))

        wb_target = (target_wb_instance
                     or robust_call(excel.Workbooks.Open, os.path.abspath(targetFile)))

        robust_call(lambda: wb_target.Name)  # sanity-check the COM object is alive

        try:
            ws_target = robust_call(lambda: wb_target.Worksheets(targetSheet))
        except Exception:
            ws_target = robust_call(lambda: wb_target.Worksheets.Add())
            robust_call(lambda: setattr(ws_target, 'Name', targetSheet))

        robust_call(lambda: ws_input.UsedRange.Copy())
        print(f"{inputFile} / {inputSheet} copied to {targetFile} / {targetSheet}")

        paste_type = -4163 if onlyValue else -4104
        robust_call(lambda: ws_target.Range(targetCell).PasteSpecial(Paste=paste_type))
        robust_call(lambda: setattr(excel, 'CutCopyMode', False))

        if refreshAll:
            robust_call(lambda: wb_target.RefreshAll())
            robust_call(lambda: excel.CalculateUntilAsyncQueriesDone())

        if target_wb_instance is None:
            robust_call(lambda: wb_target.Save())
            print(f"✅ Saved: {targetSheet}")

    finally:
        # Always close the input workbook (we opened it, we close it)
        if wb_input is not None:
            safe_close_wb(wb_input, False)

        # Only close / quit things we ourselves opened
        if target_wb_instance is None:
            if wb_target is not None:
                safe_close_wb(wb_target, True)
            try:
                robust_call(excel.Quit)
            except Exception:
                pass
            if _com_initialized:
                pythoncom.CoUninitialize()

def download_from_gdrive(file_id, file_name, save_dir):
    base_url = "https://drive.google.com/uc?export=download"
    session = requests.Session()
    # Step 1: Initial request
    response = session.get(base_url, params={"id": file_id}, stream=True)
    # Step 2: Extract confirmation token if it exists
    # Google Drive uses a "confirm" parameter for large files
    token = None
    for key, value in response.cookies.items():
        if key.startswith("download_warning"):
            token = value
            break
    # If token not in cookies, try to find it in the response text (for very large files)
    if not token:
        # We only check the first few bytes to avoid loading the whole file into memory
        content = response.text
        if 'confirm=' in content:
            token = content.split('confirm=')[1].split('&')[0].split('"')[0]

    if token:
        params = {"id": file_id, "confirm": token}
        response = session.get(base_url, params=params, stream=True)

    # Step 3: Ensure directory exists using Path for better compatibility
    save_path = Path(save_dir)
    save_path.mkdir(parents=True, exist_ok=True)
    full_path = save_path / file_name

    # Step 4: Write file in chunks and check for success
    if response.status_code == 200:
        with open(full_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=32768):
                if chunk:
                    f.write(chunk)
        print(f"✅ {file_name} successfully downloaded to {save_dir}")
    else:
        print(f"❌ Failed to download {file_name}. Status code: {response.status_code}")

def get_currency_rate(period, currency):
    year = int(period[:4])
    period = int(period[4:])
    base_date = datetime(year, 3, 1)
    month_offset = period - 12
    real_date = base_date + relativedelta(months=month_offset)
    url = f"https://www.bankofcanada.ca/valet/observations/group/FX_RATES_MONTHLY/json?start_date={real_date.strftime('%Y-%m')}-01"
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
    except Exception as e:
        print(f"Error fetching data: {e}, return rate 1.0")
        return 1.0
    key = "FXM" + currency + "CAD"
    target_date = real_date.strftime('%Y-%m') + "-01"

    for obs in data.get("observations", []):
        if obs.get("d") == target_date:
            return float(obs.get(key, {}).get("v"))
        
    return 1.0

def move_and_replace_files(source_dir, target_dir):
    src_path = Path(source_dir).resolve()
    dst_path = Path(target_dir).resolve()

    if not src_path.exists():
        print(f"❌ Source directory does not exist: {src_path}")
        return

    dst_path.mkdir(parents=True, exist_ok=True)

    print(f"🚚 Moving files from {src_path} to {dst_path}...")

    for item in src_path.iterdir():
        if item.is_file():
            dest_file = dst_path / item.name
            
            try:
                if dest_file.exists():
                    dest_file.unlink()
                
                shutil.move(str(item), str(dest_file))
                print(f"✅ Moved: {item.name}")
                
            except PermissionError:
                print(f"❌ Permission Denied: Could not move {item.name}. File might be open.")
            except Exception as e:
                print(f"❌ Error moving {item.name}: {e}")

    print("🏁 All files processed.")

def get_onedrive_path():
    for key_path in [
        r"Software\Microsoft\OneDrive\Accounts\Business1",
        r"Software\Microsoft\OneDrive",
    ]:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                path, _ = winreg.QueryValueEx(key, "UserFolder")
                return Path(path)
        except FileNotFoundError:
            continue
            
    env_path = os.environ.get("OneDrive") or os.environ.get("OneDriveCommercial")
    if env_path:
        return Path(env_path)
    
    return None