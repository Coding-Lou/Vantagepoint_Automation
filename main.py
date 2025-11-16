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

        util.download_with_progress(download_url, temp_new_exe, latest_version)
        
        updater = os.path.join(os.path.dirname(sys.argv[0]), "updater.exe")
        subprocess.Popen([updater, sys.argv[0], temp_new_exe])

        print("Update started. Exiting old program...")
        sys.exit(0)
    
    except Exception as e:
        print(f"Failed to check/update version: {e}")
        return

def main():
    LOGIN = util.check_login()

    MENU = """
======================================================
                   MAIN MENU
======================================================
  1) AP Remittance
  2) AR Noticement
  3) Export Project Status Report
  4) Merge PDF

  0) Exit
======================================================
"""

    
    while True:
        print(MENU)
        userInput = input("Enter your choice: ").strip()

        if userInput == "0":
            print("Exiting program...")
            return

        # Actions that require login
        login_required_actions = {"1", "2", "3", "6"}
        if userInput in login_required_actions and not LOGIN:
            while not LOGIN :
                login.sso_login()
                LOGIN = util.check_login()

        # Handle menu selection
        if userInput == "1":
            ap.ap_main()

        elif userInput == "2":
            ar.ar_main()

        elif userInput == "3":
            project_status.main()

        elif userInput == "4":
            option = input("Merge Amazon invoices? (Y/N): ").strip().upper()
            if option.startswith("Y"):
                util.merge_amazon_invoices()
            else:
                util.merge_pdfs()

        # elif userInput == "6":
        #     labour_process()

        else:
            print("Invalid option. Please try again.")

        print("\n" + "-" * 55 + "\n")
        print()

if __name__=="__main__":
    #print("Checking if the program is latest version ...")
    #check_update()
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
