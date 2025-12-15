import util
import login
import requests
import os
import re
from openpyxl import Workbook
import csv
import local_log
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from openpyxl.worksheet.table import Table
from collections import defaultdict

global HEADERS
HEADERS = util.set_headers()
# Log Conifg
global CONSOLE_OUTPUT
CONSOLE_OUTPUT = local_log.DualOutput("runtime_log.txt")
FolderName = "C:\\Users\\jlou\\OneDrive - QCA Systems Ltd\\Documents\\Automation\\daily_receving"

def get_report_date():
    today = datetime.now(ZoneInfo("America/Vancouver"))

    weekday = today.weekday()  # Monday=0, Sunday=6

    if weekday == 0:
        last_workday = today - timedelta(days=3)
    elif weekday == 6:
        last_workday = today - timedelta(days=2)
    elif weekday == 5:
        last_workday = today - timedelta(days=1)
    else:
        last_workday = today - timedelta(days=1)

    date = last_workday.strftime("%Y-%m-%d")
    return date

def download_receive_report():
    util.check_folder(FolderName)
    util.clear_folder(FolderName)    
    try:
        # Step 1: Build
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/Purchasing/Received Purchase Order Items","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"yyyy-MM-dd","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.1,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Received Purchase Order Items","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":0.1,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[],"ReportColumns":[{"heading":"PO Number","width":0.7,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"PONumber","username":"","customGridColumnSort":""},{"heading":"Vendor Name","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"VendorName","username":"","customGridColumnSort":""},{"heading":"Item","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Item","username":"","customGridColumnSort":""},{"heading":"Description","width":1.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Description","username":"","customGridColumnSort":""},{"heading":"Order Date","width":0.625,"format":"yyyy-MM-dd","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"OrderDate","username":"","customGridColumnSort":""},{"heading":"Receive Date","width":0.65,"format":"yyyy-MM-dd h:mm:ss tt","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ReceiveDate","username":"","customGridColumnSort":""},{"heading":"Qty Ordered","width":0.65,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"OrderQty","username":"","customGridColumnSort":""},{"heading":"Qty Prev Accepted","width":0.65,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"PrevAcceptedQty","username":"","customGridColumnSort":""},{"heading":"Qty Accepted","width":0.65,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"AcceptedQty","username":"","customGridColumnSort":""},{"heading":"Open Qty","width":0.65,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"OpenQty","username":"","customGridColumnSort":""},{"heading":"Amount Ordered","width":0.8,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"OrderAmount","username":"","customGridColumnSort":""},{"heading":"Accepted Amount","width":0.8,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"AcceptedAmount","username":"","customGridColumnSort":""},{"heading":"Requestor","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Requestor","username":"","customGridColumnSort":""},{"heading":"Requestor Name","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"RequestorName","username":"","customGridColumnSort":""}],"ReportSections":[],"baseShowDetail":"Y","ReceiveDateGroup":"2","chkRejected":"N","baseRecordSelection":"","SRecDate":DATE + "T00:00:00.000","ERecDate":DATE + "T00:00:00.000","saveOptionRole":"[CREATOR_USERNAME]","baseOriginalFavoriteId":"3b66a581c3054792b0ab0bdb4d8d2f35"}}
    
        response = requests.post(url, headers=HEADERS, json=payload  )
        data = response.json()
        report_path_raw = data["return"]["ReportPath"]
        report_path = report_path_raw.replace(" ", "%20")

        # Step 2: Get Nonce
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Security/Nonce"
        payload = {}
        response = requests.post(url, headers=HEADERS, json=payload  )
        nonce = response.json()

        # Step 3: Get Viewer
        url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/reporting/viewer.aspx?&nonce={nonce}&ResetReportViewerOnPreview=Y&reportPath={report_path}&allowSchedule=Y&origReportPath=/Standard/Purchasing/Received%20Purchase%20Order%20Items&reportName=Received%20Purchase%20Order%20Items"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("Error")
        
        exportFileName = f"receving_export.csv"

        # Step 5: Download the csv report
        url =  f"https://qcadeltek03.qcasystems.com/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd?ReportSession={report_session.group(1)}&Culture=1033&CultureOverrides=True&UICulture=2057&UICultureOverrides=True&ReportStack=1&ControlID={control_id.group(1)}&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer&OpType=Export&FileName=Received+Purchase+Order+Items&ContentDisposition=OnlyHtmlInline&Format=CSV" 
        
        response = requests.get(url, headers = HEADERS,stream=True  )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(FolderName, exportFileName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            CONSOLE_OUTPUT.tqdm_write("✅ "+ csvName+" Downloaded")
        
        with open(csvName, "r", encoding="utf-8") as f:
            lines = f.readlines()
        lines = lines[3:]
    
        with open(csvName, "w", encoding="utf-8") as f:
            f.writelines(lines)
    
    except Exception as e:
        print(f"Error in downloading receive report {e}")
        return ""

def get_master_key(po):
    try:
        url = f"https://qcadeltek03.qcasystems.com/vantagepoint/vision/PurchaseReceive/POMaster/?searchType=ALL&filter={po}&page=1&pagesize=100&order=name"

        response = requests.get(url, headers=HEADERS )
        data = response.json()
        if len(data) == 0:
            return ""
        return data[0]["MasterPKey"]
    except Exception as e:
        print(f"Get master key of {po} had error {e}")
        return ""

def get_receiving_list(masterKey):
    try:
        url = f"https://qcadeltek03.qcasystems.com/vantagepoint/vision/PurchaseReceive/POMaster/{masterKey}/ReceiveMaster"

        response = requests.get(url, headers=HEADERS)
        data = response.json()
        return data
    except Exception as e:
        print(f"Get receiving list from masterkey {masterKey} had error {e}")
        return None
    
def get_receiving_detail(masterKey, pkey):
    try:
        url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/PurchaseReceive/POMaster/{masterKey}/ReceiveDetail"
        response = requests.get(url, headers=HEADERS)

        data = response.json()
        return list(filter(lambda x: x["PKey"] == pkey and x["AcceptedQty"] > 0, data))
    
    except Exception as e:
        print(f"Get receiving list from masterkey {masterKey} had error {e}")
        return None

def get_distribution(masterKey):
    try:
        url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/PurchaseOrder/POMaster/DefaultDistribution/{masterKey}"
        response = requests.get(url, headers=HEADERS)

        data = response.json()[0]
        return data["WBS1"]
    except Exception as e:
        print(f"Get receiving list from masterkey {masterKey} had error {e}")
        return ""

def get_pm_details(projectNum):
    try:
        url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/project/{projectNum}?meta=channel%2Caccess"
        response = requests.get(url, headers=HEADERS)

        data = response.json()[0]
        return data
    except Exception as e:
        print(f"Get receiving list from masterkey {projectNum} had error {e}")
        return ""

def get_requester_email(shortName):
    try:
        url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/employee/{shortName}"
        response = requests.get(url, headers=HEADERS)
        data = response.json()[0]
        return data["EMail"]
    except Exception as e:
        print(f"Get requester email {shortName} had error {e}")
        return ""

def iteratePO():
    try:
        inputFileName = f"receving_export.csv"
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(["ClientID","From","To","CC","Subject","AttachmentName","AttachmentContent","Body"])

        bodyDict = defaultdict(lambda: {"cc": set(), "body": ""})
        poDict = {}
        pmDict = {}

        FROM = util.get_config(["RECEIVE_NOTIFICATION", "FROM"])
        CC = util.get_config(["RECEIVE_NOTIFICATION", "CC"])
        SUBJECT = util.get_config(["RECEIVE_NOTIFICATION", "SUBJECT"])
        GREETING = util.get_config(["RECEIVE_NOTIFICATION", "GREETING"])
        BODY = util.get_config(["RECEIVE_NOTIFICATION", "BODY"])
        records = 1

        with open( f"{FolderName}/{inputFileName}", mode='r', encoding='utf-8') as file:
            reader = csv.DictReader(file)

            for row in reader:
                po = row["detail_PONumber"][:5]

                if po not in poDict:
                    masterKey = get_master_key(po)
                    projectNum = get_distribution(masterKey)
                    projectDetails = get_pm_details(projectNum)

                    projMgrName = projectDetails["ProjMgrName"].split(" ")[0]
                    projMgrEmail = projectDetails["ProjMgrEmail"]

                    poDict[po] = {
                        "projMgrName": projMgrName,
                        "projMgrEmail": projMgrEmail,
                        "projectNum": projectNum
                    }
                    pmDict[projMgrName] = projMgrEmail
                else:
                    projMgrName = poDict[po]["projMgrName"]
                    projectNum = poDict[po]["projectNum"]

                pmData = bodyDict[projMgrName]

                if not pmData["body"]:
                    pmData["body"] = (
                        f"<p>Hi {projMgrName}:</p>"
                        f"<p>{GREETING}</p>"
                        "<table border='1' cellspacing='0' cellpadding='6'>"
                        "<thead><tr>"
                        "<th>PO</th><th>Project Number</th><th>Item</th>"
                        "<th>Ordered<br/>Qty</th><th>Previous Received<br/>Qty</th>"
                        "<th>Today Received<br/>Qty</th><th>Remaining<br/>Qty</th>"
                        "<th>Unit Amount<br/>$CAD</th><th>Vendor</th><th>Requestor</th>"
                        "</tr></thead><tbody>"
                    )

                item_or_desc = row['detail_Item'] if row['detail_Item'] else row['detail_Description']
                open_qty = float(row['detail_OrderQty']) - float(row['detail_PrevAcceptedQty']) - float(row['detail_AcceptedQty'])
                accepted_qty = float(row["detail_AcceptedQty"])
                unitPrice = round( float(row["detail_AcceptedAmount"]) / accepted_qty, 2 ) if accepted_qty else 0
                vendorName = row["detail_VendorName"].split(" ")[0]

                html_row = f"""<tr><td>{po}</td><td>{projectNum}</td><td>{item_or_desc}</td><td>{row['detail_OrderQty']}</td><td>{row['detail_PrevAcceptedQty']}</td><td>{row['detail_AcceptedQty']}</td><td>{open_qty}</td><td>{unitPrice}</td><td>{vendorName}</td><td>{row['detail_RequestorName']}</td></tr>"""

                bodyDict[poDict[po]["projMgrName"]]['body'] += html_row
                if row["detail_Requestor"]:
                    bodyDict[projMgrName]['cc'].add(get_requester_email(row["detail_Requestor"]))

        for pm, data in bodyDict.items():
            ccList = CC
            for email in data['cc']:
            # ["PM_Name","From","To","CC","Subject","AttachmentName","AttachmentContent","Body"]
                ccList += ";"+ email
            row = [pm, FROM, pmDict[pm] , ccList, f"{SUBJECT}{DATE}", inputFileName, f"{FolderName}\\{inputFileName}" , data['body'] + BODY]
            ws.append(row)
            records += 1
    except Exception as e:
        CONSOLE_OUTPUT.tqdm_write("❌Error in iterator the PO")
    
    try:
        ws = wb.active
        table_range = f"A1:H{records}"
        if "Table1" in ws.tables:
            del ws.tables["Table1"]
        tab = Table(displayName = "Table1", ref = table_range)
        ws.add_table(tab)
        wb.save(f"{FolderName}/output.xlsx")

        print(f"Output file is output.xlsx")
    except Exception as e:
        CONSOLE_OUTPUT.tqdm_write("❌Error in function save_excel()")

def main():
    global DATE
    DATE = get_report_date()

    LOGIN = util.check_login()
    if not LOGIN:
        while not LOGIN :
            login.sso_login()
            LOGIN = util.check_login()

    download_receive_report()
    iteratePO()

if __name__ == "__main__":
    main()
