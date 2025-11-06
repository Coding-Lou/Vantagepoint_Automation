import login
import time
import util
import subprocess
import requests
import os
import sys

GITHUB_REPO = "Coding-Lou/Vantagepoint_Automation"
EXE_NAME = "start.exe" 
LOGIN = False
VERSION = util.get_config("VERSION")

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

        temp_path = os.path.join(os.getenv("TEMP"), EXE_NAME)

        with requests.get(download_url, stream=True) as download:
           with open(temp_path, "wb") as f:
               for chunk in download.iter_content(chunk_size=8192):
                   if chunk:
                       f.write(chunk)

        subprocess.Popen([temp_path, sys.argv[0]])
        print("Update started, exiting current program...")
        sys.exit(0)
    
    except Exception as e:
        print(f"Failed to check/update version: {e}")
        return
    
def main():
    login.sso_login()
    #userInput = ""
    #while userInput != "0":
    #    print("Input your choice: ")
    #    #print("1 - AP Remittance")
    #    #print("2 - AR Noticement")
    #    print("3 - Export project status report")
    #    #print("4 - Merging Amazon Invoice PDF")
    #    #print("5 - Merging PDF")
    #    #print("6 - month-end revenue audit (careful use not finished yet)")
    #    print()
    #    print("0 - Exit the program")
    #    print("--------------------------------------------------------")
    #    userInput = input()
    #    
    #    if (userInput == "0"): return

    #    if (not LOGIN and userInput in ["1", "2", "3", "6"]):
    #        LOGIN = True
    #        login.sso_login()

        #if (userInput == "1"): ap_main()
        #if (userInput == "2"): ar_main()
        #if (userInput == "3"): util.merge_amazon_invoices()
        #if (userInput == "4"): util.merge_pdfs()
        #if (userInput == "4"): projectStatus_main()
        #if (userInput == "5"): labour_process()    

        #projectStatus_main()
        #amazon_print()
    

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
    input("🎉Done, Have a good day. Press Enter to exit the script.")
