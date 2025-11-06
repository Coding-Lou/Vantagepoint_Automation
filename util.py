import json
from pypdf import PdfReader, PdfWriter
from pathlib import Path
import os
import requests

def show_welcome_banner():
    banner = r"""
     ██████╗    ██████    █████╗ 
    ██╔═══██╗  ██        ██╔══██╗
    ██║   ██║  ██        ███████║
    ██║▄▄ ██║  ██        ██╔══██║
    ╚██████╔╝  ╚██████   ██║  ██║
     ╚══▀▀═╝    ╚═════   ╚═╝  ╚═╝
        🔧 QCA Accounting Team Automation Script
                                            --- Made by Mr.Jay
    """
    print(banner)
    print("🚀 Welcome! \n")

def get_config(key):
    try:
        with open("config.json", "r", encoding="utf-8") as f:
            config = json.load(f)
        return config[key]
    
    except Exception as e:
        print(f"⚠️ Get {key} failed: ", e)

def set_config(key, value):
    try:
        with open("config.json", "r", encoding="utf-8") as f:
            config = json.load(f)
        
        config[key] = value
        
        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
            f.flush()

        #print(f"💾 {key} been updated to {value}. ")
        #print()

    except Exception as e:
        print(f"⚠️ {key} updated failed: ", e)
        #print()

def check_login():
    try: 
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/visionservices.asmx/GetIAccessConfiguration"
        payload = {"sessionID": get_config("TOKEN")}
        response = requests.post(url, headers = set_headers(), json = payload)
        if response.status_code == 200:
            data = response.json()
            print("✅ Login Success, User: " + data["d"]["UserInfo"]["EMail"])
            print()
            return True
        else: 
            print("❌ Login Failed, try again.")
            return False
    except:
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

def merge_pdfs(folder_path, output_filename):
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

    print(f"\n🎉 Success the merged pdf file：{output_path}")

def set_headers():
    WWWBEARER = get_config("WWWBEARER")
    TOKEN = get_config("TOKEN")
    COOKIES = get_config("COOKIES")
    headers = {
        "accept": "application/json, text/javascript, */*; q=0.01",
        "Content-Type": "application/json; charset=UTF-8",
        "www-bearer": WWWBEARER,
        "Token": TOKEN,
        "Cookie": COOKIES
    }
    return headers

def assamble_projects(projects):
    searchOptions = [{"name":"Status","value":"[IS_EMPTY]","type":"dropdown","seq":1,"tableName":"PR","opp":"!=","condition":"and","searchLevel":1,"valueDescription":""}]
    for project in projects:
        searchOptions.append({"name":"selectedResultIds","value":project,"type":"wbs1","seq":2,"searchLevel":0,"valueDescription":"Monthly Maintenance"})
    
    return searchOptions