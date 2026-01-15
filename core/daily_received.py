import tools.util as util
import tools.form_notifier as form_notifier
import requests
import os
import re
from openpyxl import Workbook
import csv
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from openpyxl.worksheet.table import Table
from collections import defaultdict
import time

FolderName = util.get_config(["RECEIVE_NOTIFICATION", "FOLDER"])

def get_report_date():
    today = datetime.now(ZoneInfo("America/Vancouver"))

    weekday = today.weekday()  # Monday=0, Sunday=6

    if weekday == 4:
        next_workday = today + timedelta(days=3)
    elif weekday == 5:
        next_workday = today + timedelta(days=2)
    elif weekday == 6:
        next_workday = today + timedelta(days=1)
    else:
        next_workday = today + timedelta(days=1)

    # date = today.strftime("%Y-%m-%d")
    return next_workday.strftime("%Y-%m-%d")

def download_receive_report():
    util.check_folder(FolderName)
    try:
        # Step 1: Build
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/Purchasing/Received Purchase Order Items","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"yyyy-MM-dd","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.1,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Received Purchase Order Items","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":0.1,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[],"ReportColumns":[{"heading":"PO Number","width":0.7,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"PONumber","username":"","customGridColumnSort":""},{"heading":"Vendor Name","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"VendorName","username":"","customGridColumnSort":""},{"heading":"Item","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Item","username":"","customGridColumnSort":""},{"heading":"Description","width":1.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Description","username":"","customGridColumnSort":""},{"heading":"Order Date","width":0.625,"format":"yyyy-MM-dd","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"OrderDate","username":"","customGridColumnSort":""},{"heading":"Receive Date","width":0.65,"format":"yyyy-MM-dd h:mm:ss tt","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ReceiveDate","username":"","customGridColumnSort":""},{"heading":"Qty Ordered","width":0.65,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"OrderQty","username":"","customGridColumnSort":""},{"heading":"Qty Prev Accepted","width":0.65,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"PrevAcceptedQty","username":"","customGridColumnSort":""},{"heading":"Qty Accepted","width":0.65,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"AcceptedQty","username":"","customGridColumnSort":""},{"heading":"Open Qty","width":0.65,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"OpenQty","username":"","customGridColumnSort":""},{"heading":"Amount Ordered","width":0.8,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"OrderAmount","username":"","customGridColumnSort":""},{"heading":"Accepted Amount","width":0.8,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"AcceptedAmount","username":"","customGridColumnSort":""},{"heading":"Requestor","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Requestor","username":"","customGridColumnSort":""},{"heading":"Requestor Name","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"RequestorName","username":"","customGridColumnSort":""}],"ReportSections":[],"baseShowDetail":"Y","ReceiveDateGroup":"2","chkRejected":"N","baseRecordSelection":"","SRecDate":DATE + "T00:00:00.000","ERecDate":get_report_date() + "T00:00:00.000","saveOptionRole":"[CREATOR_USERNAME]","baseOriginalFavoriteId":"3b66a581c3054792b0ab0bdb4d8d2f35"}}
    
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
            print("Error")
        
        exportFileName = f"receving_export {DATE}.csv "

        # Step 5: Download the csv report
        url =  f"https://qcadeltek03.qcasystems.com/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd?ReportSession={report_session.group(1)}&Culture=1033&CultureOverrides=True&UICulture=2057&UICultureOverrides=True&ReportStack=1&ControlID={control_id.group(1)}&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer&OpType=Export&FileName=Received+Purchase+Order+Items&ContentDisposition=OnlyHtmlInline&Format=CSV" 
        
        response = requests.get(url, headers = HEADERS,stream=True  )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(FolderName, exportFileName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            print("✅ "+ csvName+" Downloaded")
        
        with open(csvName, "r", encoding="utf-8") as f:
            lines = f.readlines()
        lines = lines[3:]
    
        with open(csvName, "w", encoding="utf-8") as f:
            f.writelines(lines)
    
    except Exception as e:
        print(f"Error in downloading receive report {e}")

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


def get_distribution(masterKey):
    try:
        url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/PurchaseOrder/POMaster/DefaultDistribution/{masterKey}"
        response = requests.get(url, headers=HEADERS)

        data = response.json()[0]
        return data["WBS1"]
    except Exception as e:
        print(f"Get receiving list from masterkey {masterKey} had error {e}")

def get_pm_details(projectNum):
    try:
        url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/project/{projectNum}?meta=channel%2Caccess"
        response = requests.get(url, headers=HEADERS)

        data = response.json()[0]
        return data
    except Exception as e:
        print(f"Get receiving list from masterkey {projectNum} had error {e}")

def get_requester_email(shortName):
    try:
        url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/employee/{shortName}"
        response = requests.get(url, headers=HEADERS)
        data = response.json()[0]
        return data["EMail"]
    except Exception as e:
        print(f"Get requester email {shortName} had error {e}")

def iterateLineItems():
    
    try:
        inputFileName = f"receving_export {DATE}.csv"
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(["ClientID","From","To","CC","Subject","AttachmentName","AttachmentContent","Body"])

        bodyDict = defaultdict(lambda: {"cc": set(),"body": {},"openQty": {},"head": ""})
        poInfoDict = {}
        pmDict = {}
        poLineDict = set()

        FROM = util.get_config(["RECEIVE_NOTIFICATION", "FROM"])
        CC = util.get_config(["RECEIVE_NOTIFICATION", "CC"])
        SUBJECT = util.get_config(["RECEIVE_NOTIFICATION", "SUBJECT"])
        GREETING = util.get_config(["RECEIVE_NOTIFICATION", "GREETING"])
        BODY = util.get_config(["RECEIVE_NOTIFICATION", "BODY"])
        STYLE = util.get_config(["RECEIVE_NOTIFICATION", "STYLE"])
        records = 1

        with open( f"{FolderName}/{inputFileName}", mode='r', encoding='utf-8') as file:
            reader = csv.DictReader(file)

            for row in reader:
                po = row["detail_PONumber"][:5]
                
                if po not in poInfoDict:
                    masterKey = get_master_key(po)
                    projectNum = get_distribution(masterKey)
                    projectDetails = get_pm_details(projectNum)

                    projMgrName = projectDetails["ProjMgrName"].split(" ")[0]
                    projMgrEmail = projectDetails["ProjMgrEmail"]

                    poInfoDict[po] = {
                        "projMgrName": projMgrName,
                        "projMgrEmail": projMgrEmail,
                        "projectNum": projectNum
                    }
                    pmDict[projMgrName] = projMgrEmail
                else:
                    projMgrName = poInfoDict[po]["projMgrName"]
                    projectNum = poInfoDict[po]["projectNum"]

                pmData = bodyDict[projMgrName]
                if not pmData["body"]:
                    pmData["body"] = {}
                    pmData["open_qty"] = {}
                    pmData["order_qty"] ={}
                    pmData["item_or_desc"] = {}
                    pmData["prev_qty"] = {}
                    pmData["accepted_qty"] = {}
                    pmData["unitPrice"] = {}
                    pmData["vendorName"] = {}
                    pmData["head"] = f"{STYLE}<p>Hi {projMgrName}:</p><p>{GREETING}</p><table border='1' cellspacing='0' cellpadding='6' style='border-collapse:collapse;'><thead><tr><th class='c'>PO</th><th class='c'>Project<br/>Number</th><th class='c'>Item</th><th class='c'>Ordered<br/>Qty</th><th class='c'>Previously<br/>Received<br/>Qty</th><th class='c'>{DATE}<br/>Received<br/>Qty</th><th class='c'>Remaining<br/>Qty</th><th class='c'>Unit Amount<br/>$CAD</th><th class='c'>Vendor</th><th class='c'>Requestor</th></tr></thead><tbody>"
               
                key = po+row["detail_Item"]+row["detail_Description"]
                poLineDict.add(key)

                if key in pmData["order_qty"]:
                    pmData["accepted_qty"][key] += float(row["detail_AcceptedQty"] or 0)
                    pmData["order_qty"][key] += float(row["detail_OrderQty"] or 0)
                    pmData["open_qty"][key] = pmData["order_qty"][key] - pmData["prev_qty"][key] - pmData["accepted_qty"][key]
                else:
                    pmData["item_or_desc"][key] = row['detail_Item'] if row['detail_Item'] else row['detail_Description']
                    pmData["order_qty"][key] = float(row["detail_OrderQty"] or 0)
                    pmData["prev_qty"][key] = float(row["detail_PrevAcceptedQty"] or 0)
                    pmData["accepted_qty"][key] = float(row["detail_AcceptedQty"] or 0)
                    pmData["open_qty"][key] = pmData["order_qty"][key] - pmData["prev_qty"][key] - pmData["accepted_qty"][key]
                    pmData["accepted_qty"][key] = float(row["detail_AcceptedQty"])
                    pmData["unitPrice"][key] = round( float(row["detail_AcceptedAmount"]) / pmData["accepted_qty"][key], 2 ) if pmData["accepted_qty"][key] else 0
                    pmData["vendorName"][key] = row["detail_VendorName"].split(" ")[0]
                
                html_row = f"<tr><td class='l'>{po}</td><td class='l'>{projectNum}</td><td class='l'>{pmData['item_or_desc'][key]}</td><td class='r'>{pmData['order_qty'][key]:,.2f}</td><td class='r'>{pmData['prev_qty'][key]:,.2f}</td><td class='r'>{pmData['accepted_qty'][key]:,.2f}</td><td class='r'>{pmData['open_qty'][key]:,.2f}</td><td class='r'>{pmData['unitPrice'][key]:,.2f}</td><td class='l'>{pmData['vendorName'][key]}</td><td class='l'>{row['detail_RequestorName']}</td></tr>"

                pmData["body"][key] = html_row
            
                if row["detail_Requestor"]:
                    pmData['cc'].add(get_requester_email(row["detail_Requestor"]))

        for pm, data in bodyDict.items():
            ccList = str(CC)
            for email in data["cc"]:
                ccList += email + ";"

            body_html = (
                data["head"]+ "".join(data["body"].values())+ BODY
            )

            ws.append([
                pm,FROM,pmDict.get(pm, ""),ccList,f"{SUBJECT}{DATE}","","",body_html
            ])
            records += 1

    except Exception as e:
        print(f"❌Error in iterator the PO {e}")
    
    try:
        ws = wb.active

        if records == 1:
            ws.append(["Jay Lou", "jlou@qcasystems.com", "jlou@qcasystems.com", "", "Daily Receiving Report", "", "", "<p>No items received on last business day.</p>"])
            records += 1

        table_range = f"A1:H{records}"
        if "Table1" in ws.tables:
            del ws.tables["Table1"]
        tab = Table(displayName = "Table1", ref = table_range)
        ws.add_table(tab)
        report_date = get_report_date()
        wb.save(f"{FolderName}/output {report_date}.xlsx")
        print(f"Output file is output {report_date}.xlsx")

    except Exception as e:
        print("❌Error in function save_excel()")

def init():
    global DATE
    DATE = datetime.now(ZoneInfo("America/Vancouver")).strftime("%Y-%m-%d")
    global HEADERS
    HEADERS = util.set_headers()

def main():
    try:
        init()
        download_receive_report()
        iterateLineItems()
        form_notifier.send_task_execution_status(
            task_name= f"Daily Receiving Report {DATE}",
            executor= "Jay Lou",
            status= "Success",
            message= f"Daily receiving report task completed successfully."
        )
    except Exception as e:
        print(f"Daily receiving report task failed with error: {e}")
        form_notifier.send_task_execution_status(
            task_name= f"Daily Receiving Report {DATE}",
            executor= "Jay Lou",
            status= "Failed",
            message= f"Daily receiving report task failed with error: {e}"
        )


if __name__ == "__main__":
    main()
