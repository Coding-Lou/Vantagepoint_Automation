import login
import time
import util
import subprocess
import requests
import os
import sys
import project_status
from openpyxl import load_workbook
import ap
import ar

GITHUB_REPO = "Coding-Lou/Vantagepoint_Automation"
EXE_NAME = "start.exe" 
VERSION = util.get_config(["VERSION"])

def check_update():
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

        util.download_with_progress(download_url, temp_new_exe)
        
        updater = os.path.join(os.path.dirname(sys.argv[0]), "updater.exe")
        subprocess.Popen([updater, sys.argv[0], temp_new_exe])

        print("Update started. Exiting old program...")
        sys.exit(0)
    
    except Exception as e:
        print(f"Failed to check/update version: {e}")
        return

def main():
    LOGIN = False
    userInput = ""
    while userInput != "0":
        print("Input your choice: ")
        print("1 - AP Remittance")
        print("2 - AR Noticement")
        print("3 - Export project status report")
        print("4 - Merging PDF")
        #print("6 - Bridge Report (careful use not finished yet)")
        print()
        print("0 - Exit the program")
        print("--------------------------------------------------------")
        userInput = input()
        
        if (userInput == "0"): return

        if (not LOGIN and userInput in ["1", "2", "3", "6"]):
            while not LOGIN:
                login.sso_login()
                LOGIN = util.check_login()

        if (userInput == "1"): ap.ap_main()
        if (userInput == "2"): ar.ar_main()
        if (userInput == "3"): project_status.main()
        if userInput == "4":
            option = input("Merge Amazon invoices? (Y/N): ")
            if option.strip().upper().startswith("Y"):
                util.merge_amazon_invoices()
            else:
                util.merge_pdfs()
        #if (userInput == "6"): labour_process()    

if __name__=="__main__":
    print("Checking if the program is latest version ...")
    check_update()
    util.show_welcome_banner()
    start_time = time.perf_counter()
    main()
    end_time = time.perf_counter()
    total_seconds = end_time - start_time
    minutes, seconds = divmod(int(total_seconds), 60)
    print(f"Total execution time {minutes} minutes, {seconds} seconds")
    print()
    print("🎉Done, Have a good day.")
    input()
