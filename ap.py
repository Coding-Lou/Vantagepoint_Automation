import sys
import util
from openpyxl import Workbook
import requests
from tqdm import tqdm
import local_log
import re
import os

global HEADERS
HEADERS = util.set_headers()
# Email
global EXCLUDE
EXCLUDE = util.get_config(["AP", "EXCLUDE"])
global MAIL_FROM
MAIL_FROM = util.get_config(["AP", "FROM"])
global CC
CC = util.get_config(["AP", "CC"])
global SUBJECT
SUBJECT = util.get_config(["AP", "SUBJECT"])
global BODY
BODY = util.get_config(["AP", "BODY"])
# Output location
global ONEDRIVEDIR
ONEDRIVEDIR = util.get_config(['ONEDRIVEDIR'])
global WORKDIR
WORKDIR = util.get_config(['WORKDIR'])
# Log Conifg
global CONSOLE_OUTPUT
CONSOLE_OUTPUT = local_log.DualOutput("runtime_log.txt")
global RECORDS
RECORDS = 1

def init_output():
    global wb
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["ClientID","From","To","CC","Subject","AttachmentName","AttachmentContent","Body","Payee"])
    util.clear_folder("ap_export")

def ap_setup_time():
    date = input("Input the Remittance Date (format 2025-05-01): ")
    
    url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/UserSettings"

    payload = {"NavigatorUserSettings":[{"SettingName":"CHECK_REVIEW_CheckDate","SettingValue":"C","_transType":"U"},{"SettingName":"CHECK_REVIEW_CheckDateFrom","SettingValue":date,"_transType":"U"},{"SettingName":"CHECK_REVIEW_CheckDateTo","SettingValue":date,"_transType":"U"},{"SettingName":"CHECK_REVIEW_PostingDate","SettingValue":"","_transType":"U"},{"SettingName":"CHECK_REVIEW_BankCode","SettingValue":"","_transType":"U"},{"SettingName":"CHECK_REVIEW_Vendor","SettingValue":"","_transType":"U"},{"SettingName":"CHECK_REVIEW_Employee","SettingValue":"","_transType":"U"},{"SettingName":"CHECK_REVIEW_Project","SettingValue":"","_transType":"U"},{"SettingName":"CHECK_REVIEW_CheckAmtFrom","SettingValue":0,"_transType":"U"},{"SettingName":"CHECK_REVIEW_CheckAmtTo","SettingValue":0,"_transType":"U"},{"SettingName":"CHECK_REVIEW_TransType","SettingValue":"PP","_transType":"U"}]}

    response = requests.post(url, headers=HEADERS, json=payload  )

def ap_get_remittance():
    try:
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/CheckReview/?startRow=1&pageSize=1000&initialLoad=true&sortDef=CheckDate_A&offset=0"
        response = requests.get(url, headers=HEADERS )

        paymentsData = response.json()

        CONSOLE_OUTPUT.write("\nTotal Remittances Records (Includes voided) : " + str(len(paymentsData)))
        CONSOLE_OUTPUT.write("-------------------------------------------")
        return paymentsData
    
    except:
        CONSOLE_OUTPUT.tqdm_write("❌ Error in function ap_get_remittance")

def ap_process_remittance(paymentsData):
    i = 0
    for i, payment in enumerate(tqdm(paymentsData, desc="Processing Progress: ", unit="payment"), start=1):
        i = i+1
        if (not payment["ClientID"] in EXCLUDE and payment["VoidPostSeq"] == 0 and payment["BankCode"] != "1107"):
            email = util.get_vendor_email(payment["ClientID"])
            ap_download_remittance(payment)
            ap_create_record(payment, email)
        else:
            if payment["ClientID"] in EXCLUDE:
                CONSOLE_OUTPUT.tqdm_write(f"⚠️ {i} {payment['Payee']} is excluded")
            if payment["BankCode"] == "1107":
                CONSOLE_OUTPUT.tqdm_write(f"⚠️ {i} {payment['Payee']} This remittance is fX")
            if payment["VoidPostSeq"] != 0:
                CONSOLE_OUTPUT.tqdm_write(f"⚠️ {i} Payment Number: {payment['CheckNumber']} has been Voided")

def ap_create_record(payment, email):
    ws = wb.active
    ws.title = "Sheet1"
    fileName = payment["Payee"] + "_" + payment["CheckNumber"] + "_" +payment["CheckDate"][0:10] + ".pdf"
    fileName = fileName.replace("/", "")
    row = [payment["ClientID"], MAIL_FROM, email,CC, SUBJECT, fileName, WORKDIR+"\\ap_export\\"+fileName, BODY, payment["Payee"]]
    ws.append(row)
    global RECORDS
    RECORDS += 1

def ap_download_remittance(payment):
    dual_out = sys.stdout
    try:
        # Step 1: Build
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Reporting/Build"
        payload = {
            "reportPath": "/Standard/AccountingGeneral/Remittance Advice",
            "reportOptions": {
                "BankCode": payment["BankCode"],
                "period": payment["Period"],
                "PostSeq": payment["PostSeq"],
                "ShowSSN": "N",
                "checkpayee": "2",
                "vendor": payment["Vendor"],
                "Employee": payment["Employee"],
                "CheckNo": f"'{str(payment['CheckNumber'])}'"
            }
        }

        response = requests.post(url, headers=HEADERS, json=payload)
        data = response.json()
        report_path_raw = data["return"]["ReportPath"]
        report_path = report_path_raw.replace(" ", "%20")

        # Step 2: Get Nonce
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Security/Nonce"
        payload = {}
        response = requests.post(url, headers=HEADERS, json=payload  )
        nonce = response.json()

        # Step 3: Get Viewer
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&reportPath="+report_path+"&allowSchedule=N&reportName=Remittance%20Advice"

        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("Error")

        # Step 5: Download PDF
        fileName = payment["Payee"] + "_" + payment["CheckNumber"] + "_" +payment["CheckDate"][0:10] + ".pdf"
        fileName = fileName.replace("/", "")
        url_pdf = ( "https://qcadeltek03.qcasystems.com"
            "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
            f"?ReportSession={report_session.group(1)}"
            "&Culture=1033&CultureOverrides=True"
            "&UICulture=2057&UICultureOverrides=True"
            "&ReportStack=1"
            f"&ControlID={control_id.group(1)}"
            "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer"
            "&OpType=Export"
            f"&FileName={fileName}"
            "&ContentDisposition=OnlyHtmlInline&Format=PDF" )
        
        response = requests.get(url_pdf, headers=HEADERS)

        if response.status_code == 200 and response.headers.get("Content-Type") == "application/pdf":
            pdfName = os.path.join("ap_export", fileName)
            with open(pdfName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            CONSOLE_OUTPUT.tqdm_write("✅ "+ fileName+" Downloaded")
        else:
            CONSOLE_OUTPUT.tqdm_write(f"❌ Error, status code: {response.status_code}")
            CONSOLE_OUTPUT.tqdm_write("Msg: ", response.content[:500])

    except Exception as e:
        CONSOLE_OUTPUT.tqdm_write("Error in download remittence")


def ap_main():
    util.check_folder("ap_export")
    init_output()
    ap_setup_time()
    paymentsData = ap_get_remittance()
    ap_process_remittance(paymentsData)
    util.save_excel(wb, RECORDS)

if __name__ == '__main__':
    ap_main()