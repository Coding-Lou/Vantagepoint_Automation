import csv
from datetime import datetime
from tools import login
import tools.util as util
import requests
import re
import os
from datetime import date
from pathlib import Path
import win32com.client
from io import StringIO
import os
import random
import string
import hashlib
import shutil
import pythoncom
import time
import getpass

TEMPLATE0 = "0 JTD Billed Invoice Summary.xlsx"
TEMPLATE1 = "1 2 3 JTD Billing.xlsx"
TEMPLATE2 = "4 5 Budget _Complete.xlsx"
TEMPLATE3 = "6 New Model_Earned Revenue Accrual.xlsx"

def download_template(base_folder, file_name):
    onedrive_root = util.get_onedrive_path()

    if onedrive_root:
        source_path = onedrive_root / "QCA Accounting Dept - Documents/03 Accounting/400 Process Improvement/Automation/template/Rev-Gen"
        print(f"Success get the OneDrive path: {onedrive_root}")
    else:
        print("OneDrive path not found. Please ensure OneDrive is installed and configured correctly.")

    save_path = Path(base_folder)
    save_path.mkdir(parents=True, exist_ok=True)

    source_file = source_path / file_name
    target_file = save_path / file_name

    if not source_file.exists():
        print(f"❌ Source does not exist: {source_file}")
        return

    try:
        if target_file.exists():
            target_file.unlink()

        shutil.copy2(source_file, target_file)
        print(f"✅ Copied template to: {target_file}")

    except PermissionError:
        print("❌ Permission denied: file may be open or locked (Excel/OneDrive).")
    except Exception as e:
        print(f"❌ Copy failed: {e}")


def get_search_list():
    url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/SaveSearchOptions/?order=name&type=wbs1&excludeCrossHubSearch=false&searchType=ALL&WBSType=WBS1&AccessGroupBy=WBS1&isLevelLock=false"
    data = requests.get(url, headers=HEADERS).json()
    search_list = []
    for item in data:
        search_list.append({
            "PKey": item["PKey"],
            "Name": item["Name"]
        })
    search_list.sort(key=lambda x: x["Name"])
    print("Search List:", [f"{item['PKey']}: {item['Name']}" for item in search_list])
    return search_list


def get_search_option_project_total(pkey):
    url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/project/?count&searchType=ALL&WBSType=WBS1&AccessGroupBy=WBS1&isLevelLock=false&filter=&page=1&pagesize=20&order=name&applicationId=&savedSearchPKey={pkey}"
    data = requests.get(url, headers=HEADERS).json()[0]["_ResultCount"]
    return data


def append_to_project_list(url, columnName, needFilter = False):
    global project_list
    response = requests.get(url, headers = HEADERS,stream=True)
    if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
        csv_text = response.text
        lines = csv_text.splitlines()
        start_index = None
        for i, line in enumerate(lines):
            if columnName in line:
                start_index = i
                break
        if start_index is None:
            raise ValueError("detail_WBS1 header not found")
        data_part = "\n".join(lines[start_index:])
        reader = csv.DictReader(StringIO(data_part))
        for row in reader:
            value = row.get(columnName)
            preCheck = value is not None
            if needFilter:
                preCheck = preCheck and (row.get('detail_Name'))
            if preCheck:
                project_list.add(value)


def download_project_list(base_folder):
    try:
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        baseRecordSelection = {"pKey":None,"name":"Records Selected","type":"wbs1","whereClauseSearch":"N","isLegacy":"N","searchOptions":[{"name":"ChargeType","value":"R","type":"dropdown","seq":1,"tableName":"PR","condition":"and","searchLevel":1,"valueDescription":"Regular"}]}
        payload = {"reportPath":"/Standard/Project/Project List","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseChart3D":"N","baseChartColumn":"","baseChartDivisor":"1","baseChartFontSize":8,"baseChartHeight":3,"baseChartLabelLines":"N","baseChartLabels":"none","baseChartLeft":1,"baseChartLegendPosition":"righttop","baseChartSeriesColumn2":"","baseChartSeriesColumn3":"","baseChartShowPosition":"1","baseChartTitle":"","baseChartTop":0.5,"baseChartType":"none","baseChartWidth":6,"baseChartXTitle":"Project Number","baseChartYTitle":"","baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"Activity","baseGroupIndent":0.1,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Project List","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":0.1,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"Project Number","sort":"ASC","color":"000080","subTotal":"N","showHeading":"N","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"projectNumber","customGridColumnSort":"","groupWBSLevel":"1"}],"ReportColumns":[{"heading":"Project","width":1.25,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"WBS1","username":"","customGridColumnSort":""},{"heading":"Status","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"StatusDesc","username":"","customGridColumnSort":""},{"heading":"Biller","width":1.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"billerName","username":"","customGridColumnSort":""},{"heading":"Name","width":2.2,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Name","username":"","customGridColumnSort":""},{"heading":"Primary Client","width":1.8,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ClientName","username":"","customGridColumnSort":""},{"heading":"Project Manager","width":1.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"prgName","username":"","customGridColumnSort":""},{"heading":"Principal","width":1.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"prinName","username":"","customGridColumnSort":""},{"heading":"Charge Type","width":0.55,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ChargeType","username":"","customGridColumnSort":""},{"heading":"USD Project","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustUSDProject","username":"","customGridColumnSort":"N"},{"heading":"Large Projects","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustMilestoneBillingProjects400k","username":"","customGridColumnSort":"N"},{"heading":"Materials Only","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustMaterialsOnly","username":"","customGridColumnSort":"N"},{"heading":"Time & Material","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustTimeMaterial","username":"","customGridColumnSort":"N"},{"heading":"Long-term Contract","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustLongtermFixedPrice","username":"","customGridColumnSort":"N"},{"heading":"No Revenue","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustNoRevenue","username":"","customGridColumnSort":"N"},{"heading":"Fixed Fee","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustLumpSumBilling","username":"","customGridColumnSort":"N"},{"heading":"Defined Scope/deliverables","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustDefinedScopeValue","username":"","customGridColumnSort":"N"}],"ReportSections":[],"baseRecordSelection":baseRecordSelection,"baseCreateActivity":"N","baseShowDetail":"Y","CurrentWBSActivityActiveWBS1Only":"N","CurrentWBSActivityActiveWBS2Only":"N","CurrentWBSActivityActiveWBS3Only":"N","CurrentWBSActivityActivityRange":"1","CurrentWBSActivityInclInvoiceActivity":"N","CurrentWBSActivityCheckLabor":"Y","CurrentWBSActivityCheckExpense":"Y","_desc_saveOptionRole":["","","","",""],"saveOptionRole":["[CREATOR_USERNAME]","CONTROLLER-RO","ACCOUNTANT","CONTROLLER","ACCOUNTING"],"baseOriginalFavoriteId":"6af3e2e5be034e1cbebec9fda1d74c98","baseSelectionRows":10430}}

        response = requests.post(url, headers=HEADERS, json=payload  )
        data = response.json()
        report_path_raw = data["return"]["ReportPath"]
        report_path = report_path_raw.replace(" ", "%20")

        # Step 2: Get Nonce
        nonceUrl = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Security/Nonce"
        payload = {}
        response = requests.post(nonceUrl, headers=HEADERS, json=payload  )
        nonce = response.json()

        # Step 3: Get Viewer
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/Project/Project%20List&reportName=Project%20List"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("Error")
        
        exportFileName_csv = "Project List_"+ date.today().strftime("%Y-%m-%d") + ".csv"

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer"
                "&OpType=Export&FileName=Project+List&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        response = requests.get(url, headers = HEADERS,stream=True  )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(base_folder, exportFileName_csv)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            print("✅ "+ csvName+" Downloaded")

        # Trim the csv file
        temp_file = csvName + ".temp"
        lines_to_skip = 4

        with open(csvName, 'r', encoding='utf-8', newline='') as f_in, \
            open(temp_file, 'w', encoding='utf-8', newline='') as f_out:
            reader = csv.reader(f_in)
            writer = csv.writer(f_out)
            for _ in range(lines_to_skip):
                next(reader, None)
            for row in reader:
                writer.writerow(row)

        os.replace(temp_file, csvName)
        
        return csvName 

    except Exception as e:
        print("⚠️ Failed to update the project list:", e)


def download_labour_details_YTD():
    global project_list
    try:
        # Step 1: Build
        baseRecordSelection = {"pKey":None,"name":"Records Selected","type":"wbs1","whereClauseSearch":"N","isLegacy":"N","searchOptions":[{"name":"ChargeType","value":"R","type":"dropdown","seq":1,"tableName":"PR","condition":"and","searchLevel":1,"valueDescription":"Regular"}]}
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/Project/Labor Detail","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseChart3D":"N","baseChartColumn":"","baseChartDivisor":"1","baseChartFontSize":8,"baseChartHeight":3,"baseChartLabelLines":"N","baseChartLabels":"none","baseChartLeft":1,"baseChartLegendPosition":"righttop","baseChartSeriesColumn2":"","baseChartSeriesColumn3":"","baseChartShowPosition":"1","baseChartTitle":"","baseChartTop":0.5,"baseChartType":"none","baseChartWidth":6,"baseChartXTitle":"Project Number","baseChartYTitle":"","baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"yyyy-MM-dd","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.05,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"Y","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Jay_Labour Detail_Rev","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":3.5,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"Project Number","sort":"ASC","color":"000080","subTotal":"N","showHeading":"N","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"projectNumber","customGridColumnSort":"","groupWBSLevel":"1"}],"ReportColumns":[{"heading":"Project","width":1.25,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"WBS1","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":baseRecordSelection,"baseCreateActivity":"N","baseShowDetail":"Y","baseLeft1":0,"baseRight1":21,"baseLeft2":0,"baseRight2":8,"baseLeft3":0,"baseRight3":8,"baseSub":"1","rollType":"Project","CurrentWBSActivityActiveWBS1Only":"N","CurrentWBSActivityActiveWBS2Only":"N","CurrentWBSActivityActiveWBS3Only":"N","CurrentWBSActivityActivityRange":"1","CurrentWBSActivityInclInvoiceActivity":"N","timeframeBackup":"YTD","timeFrame":"radio1","periodType":"2","transDetail":"1","curSelection":"2","showComments":"N","showUnposted":"N","atCost":"2","CurrentWBSActivityCheckLabor":"Y","CurrentWBSActivityCheckExpense":"Y","CurrentWBSActivityUnpostedLabor":"N","employeeWhere":"","emOwnerfl":"","dateStart":"2026-02-01T00:00:00","dateEnd":"2026-02-01T00:00:00.000","periodStart":"202601","periodEnd":202611,"baseSelectionRows":None,"saveOptionRole":"[CREATOR_USERNAME]","baseOriginalFavoriteId":"005225c84127485fa272f83ae828f266","LaborPostingRun":[]}}

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
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/Project/Labor%20Detail&reportName="

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("⚠️ Failed to download the YTD invoices")

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer"
                "&OpType=Export&FileName=Jay_Labour+Detail_Rev&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        append_to_project_list(url, "detail_WBS1", True)

    except Exception as e:
        print("⚠️ Failed to download the labour details:", e) 


def download_invoice_YTD():
    global project_list
    try:
        # Step 1: Build
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/AccountsReceivable/Invoice Register","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseChart3D":"N","baseChartColumn":"Other","baseChartDivisor":"1","baseChartFontSize":8,"baseChartHeight":3,"baseChartLabelLines":"N","baseChartLabels":"none","baseChartLeft":1,"baseChartLegendPosition":"righttop","baseChartSeriesColumn2":"","baseChartSeriesColumn3":"","baseChartShowPosition":"1","baseChartTitle":"","baseChartTop":0.5,"baseChartType":"none","baseChartWidth":6,"baseChartXTitle":"Primary Client Name","baseChartYTitle":"Other","baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"yyyy-MM-dd","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.1,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Invoice Register","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"Y","baseStartColumnPosition":0.5,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"Primary Client Name","sort":"ASC","color":"000000","subTotal":"Y","showHeading":"Y","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"clientName","customGridColumnSort":"","groupWBSLevel":"1"},{"label":"Project Number","sort":"ASC","color":"000000","subTotal":"N","showHeading":"N","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"projectNumber","customGridColumnSort":"","groupWBSLevel":"1"}],"ReportColumns":[{"heading":"Project","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"WBS1","username":"","customGridColumnSort":""},{"heading":"Date","width":0.7,"format":"yyyy-MM-dd","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"TransDate","username":"","customGridColumnSort":""},{"heading":"Invoice","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"InvoiceNumber","username":"","customGridColumnSort":""},{"heading":"Total","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"TotalAmt","username":"","customGridColumnSort":""},{"heading":"Prof Fees","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col1","username":"","customGridColumnSort":""},{"heading":"H/W Sales","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col2","username":"","customGridColumnSort":""},{"heading":"S/W Sales","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col3","username":"","customGridColumnSort":""},{"heading":"O/S Services","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col4","username":"","customGridColumnSort":""},{"heading":"Reimbursable","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col5","username":"","customGridColumnSort":""},{"heading":"EHF","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col6","username":"","customGridColumnSort":""},{"heading":"Other","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Other","username":"","customGridColumnSort":""},{"heading":"Taxes","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col7","username":"","customGridColumnSort":""},{"heading":"Net Revenue","width":1,"format":"###T###T###T###D##;(###T###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Net Revenue","header1":"Net Revenue","header2":"","detailExpression":"[TotalAmt]-[Col7]","groupExpression":"[TotalAmt]-[Col7]","queryJoin":"","checkSecurity":"N","altHeader1":"","altHeader2":"","queryColumn":"","calculatedColumnType":"ALLFRAMES","username":"DPALACIO","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":"","baseShowDetail":"Y","baseLeft1":0,"baseRight1":21,"baseLeft2":0,"baseRight2":8,"baseLeft3":0,"baseRight3":8,"baseSub":"1","rollType":"Project","timeframe":"YTD","radioTF":"radio1","tfCYJ":"Y","clientInfo":"None","InterestCol":"0","txtShowLink":"N","_desc_saveOptionRole":["","",""],"saveOptionRole":["ACCOUNTANT","ACCOUNTING","[CREATOR_USERNAME]"],"baseOriginalFavoriteId":"1589925385074d90a125301801f1c864","baseSelectionRows":11386}}

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
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/AccountsReceivable/Invoice%20Register&reportName=Invoice%20Register"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("⚠️ Failed to download the YTD invoices")

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer"
                "&OpType=Export"
                f"&FileName=Invoice+Register"
                "&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        append_to_project_list(url, "detail_WBS1")

    except Exception as e:
        print("⚠️ Failed to download the invoices register:", e)


def download_GL(base_folder, startPeriod, endPeriod, needDownload, baseRecordSelection, fileName = None):
    global project_list

    try:
        # Step 1: Build
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/DataExport/GeneralLedgerDataSource","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"General Ledger Export","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":0,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[],"ReportColumns":[{"heading":"Account Numbers g","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Account","username":"","customGridColumnSort":""},{"heading":"Account Code","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Account_Code","username":"","customGridColumnSort":""},{"heading":"Account Description","width":1.375,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Account_Description","username":"","customGridColumnSort":""},{"heading":"Trans Type","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Transaction_Type","username":"","customGridColumnSort":""},{"heading":"Trans SubType","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Transaction_SubType","username":"","customGridColumnSort":""},{"heading":"Date","width":0.875,"format":"M/d/yyyy","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Transaction_Date","username":"","customGridColumnSort":""},{"heading":"Ref. No","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Reference_Number","username":"","customGridColumnSort":""},{"heading":"Period","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Period","username":"","customGridColumnSort":""},{"heading":"Post Seq","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Posting_Sequence","username":"","customGridColumnSort":""},{"heading":"Description 1","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Description_1","username":"","customGridColumnSort":""},{"heading":"Description 2","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Description_2","username":"","customGridColumnSort":""},{"heading":"Project","width":1.25,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"wbs1","username":"","customGridColumnSort":""},{"heading":"Task","width":0.875,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"wbs2","username":"","customGridColumnSort":""},{"heading":"Subtask","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"wbs3","username":"","customGridColumnSort":""},{"heading":"Amount","width":1,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"amount","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":baseRecordSelection,"timeframe":"PeriodRange","ReportingPeriod":"1","W2Year":"","W2Qtr":"","W2Quarter":"","startPeriod":startPeriod,"endPeriod":endPeriod,"PeriodStart":startPeriod,"PeriodEnd":endPeriod,"baseSelectionRows":None,"_desc_saveOptionRole":["",""],"saveOptionRole":["ACCOUNTANT","[CREATOR_USERNAME]"],"baseOriginalFavoriteId":"1C4980783B8643039DF4DD0EADCE26AD"}}

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
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/DataExport/GeneralLedgerDataSource&reportName=General%20Ledger%20Export"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("⚠️ Failed to retrieve the General Ledger")

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer"
                "&OpType=Export&FileName=General+Ledger+Export&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        if needDownload:
            exportFileName_csv = "General Ledger_"+ fileName + date.today().strftime("%Y-%m-%d") + ".csv"
            response = requests.get(url, headers = HEADERS,stream=True  )
            if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
                csvName = os.path.join(base_folder, exportFileName_csv)
                if os.path.exists(csvName):
                    os.remove(csvName)
                with open(csvName, "wb") as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)

                print("✅ "+ csvName+" Downloaded")
            return os.path.abspath(csvName)
        else:
            append_to_project_list(url,  "detail_wbs1")
            return None

    except Exception as e:
        print("⚠️ Failed to download the General Ledger:", e)


def download_vp_revgen(update = False):
    global project_list
    try:
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        baseRecordSelection = {"pKey":None,"name":"Records Selected","type":"wbs1","whereClauseSearch":"N","isLegacy":"N","searchOptions":[{"name":"ChargeType","value":"R","type":"dropdown","seq":1,"tableName":"PR","condition":"and","searchLevel":1,"valueDescription":"Regular"}]}
        payload = {"reportPath":"/Standard/Project/Office Earnings","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseChart3D":"N","baseChartColumn":"unb1","baseChartDivisor":"1","baseChartFontSize":8,"baseChartHeight":3,"baseChartLabelLines":"N","baseChartLabels":"none","baseChartLeft":1,"baseChartLegendPosition":"righttop","baseChartSeriesColumn2":"","baseChartSeriesColumn3":"","baseChartShowPosition":"1","baseChartTitle":"","baseChartTop":0.5,"baseChartType":"none","baseChartWidth":6,"baseChartXTitle":"Biller Number","baseChartYTitle":"Labour Unbilled Amount","baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.1,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Office Earnings","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"Y","baseShowTotalsOnHeader":"Y","baseStartColumnPosition":2,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"Biller Number","sort":"ASC","color":"000000","subTotal":"N","showHeading":"Y","pageHeading":"N","collapseExpand":"E","line":"None","pageBreak":"N","groupID":"biller","customGridColumnSort":"","groupWBSLevel":"1"},{"label":"Status","sort":"ASC","color":"000000","subTotal":"N","showHeading":"Y","pageHeading":"N","collapseExpand":"E","line":"None","pageBreak":"N","groupID":"projectStatus","customGridColumnSort":"","groupWBSLevel":"1"},{"label":"Primary Client Name","sort":"ASC","color":"000000","subTotal":"N","showHeading":"Y","pageHeading":"N","collapseExpand":"E","line":"None","pageBreak":"N","groupID":"clientName","customGridColumnSort":"","groupWBSLevel":"1"},{"label":"Project Number","sort":"ASC","color":"000080","subTotal":"Y","showHeading":"Y","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"projectNumber","customGridColumnSort":"","groupWBSLevel":"1"}],"ReportColumns":[{"heading":"Labour Unbilled","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"unb1","username":"","customGridColumnSort":""},{"heading":"Hardware Unbilled","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"unb2","username":"","customGridColumnSort":""},{"heading":"Software Unbilled","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"unb3","username":"","customGridColumnSort":""},{"heading":"Outside Services Unbilled","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"unb4","username":"","customGridColumnSort":""},{"heading":"Others (Expenses) Unbilled","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"unb5","username":"","customGridColumnSort":""},{"heading":"Other Unbilled","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"unbOther","username":"","customGridColumnSort":""},{"heading":"Unbilled","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"unb","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":baseRecordSelection,"baseCreateActivity":"N","baseShowDetail":"Y","baseLeft1":0,"baseRight1":21,"baseLeft2":0,"baseRight2":8,"baseLeft3":0,"baseRight3":8,"baseSub":"1","rollType":"Project","CurrentWBSActivityActiveWBS1Only":"N","CurrentWBSActivityActiveWBS2Only":"N","CurrentWBSActivityActiveWBS3Only":"N","CurrentWBSActivityActivityRange":"1","CurrentWBSActivityInclInvoiceActivity":"N","drillDownSort":"1","budgetSelection":"1","ETCDRadioChecked":"radioETCD1","atCost":"2","labDrillDown":"1","expDrillDown":"1","showCur":"N","showYTD":"N","showJTD":"Y","custDate":"N","LabelCustDate":"Period Range","showOverhead":"N","estimateOverhead":"N","showUnposted":"N","chkIncludeCommitPO":"N","useSummaryTable":"N","excludeContractsNotinFees":"N","CurrentWBSActivityCheckLabor":"Y","CurrentWBSActivityCheckExpense":"Y","PrintDirects":"Y","ETCDate":"5/16/2024 1:38:16 PM","baseChartColumnDisplayTimeframe":"JTD","SummaryTableLastUpdate":"","CurrentWBSActivityUnpostedLabor":"N","CurrentWBSActivityCommittedExp":"N","_desc_saveOptionRole":["","","","",""],"saveOptionRole":["ACCOUNTING","CONTROLLER-RO","[CREATOR_USERNAME]","ACCOUNTANT","CONTROLLER"],"baseOriginalFavoriteId":"98fdd0394b294d66bbdbf1c4592e1ad4"}}

        response = requests.post(url, headers=HEADERS, json=payload  )
        data = response.json()
        report_path_raw = data["return"]["ReportPath"]
        report_path = report_path_raw.replace(" ", "%20")

        # Step 2: Get Nonce
        nonceUrl = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Security/Nonce"
        payload = {}
        response = requests.post(nonceUrl, headers=HEADERS, json=payload  )
        nonce = response.json()

        # Step 3: Get Viewer
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/Project/Office%20Earnings&reportName=Office%20Earnings"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("Error")
        
        exportFileName_csv = "VP Rev_gen_"+ date.today().strftime("%Y-%m-%d") + ".csv"
        if update:
            exportFileName_csv = "Updated VP Rev_gen_"+ date.today().strftime("%Y-%m-%d") + ".csv"

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer"
                "&OpType=Export&FileName=Office+Earnings&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        response = requests.get(url, headers = HEADERS,stream=True  )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(r"C:\\temp\\revenue_accrual", exportFileName_csv)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            print("✅ "+ csvName+" Downloaded")
        
        # Trim the csv file
        temp_file =  os.path.join(r"C:\\temp\\revenue_accrual", "temp_"+exportFileName_csv)
        lines_to_skip = 4

        with open(csvName, 'r', encoding='utf-8', newline='') as f_in, \
            open(temp_file, 'w', encoding='utf-8', newline='') as f_out:
                reader = csv.reader(f_in)
                writer = csv.writer(f_out)
                for _ in range(lines_to_skip):
                    next(reader, None)
                for row in reader:
                    if len(row) <= 15 or row[15] == "":
                        continue
                    if len(row) <= 7 or "Project Number:" not in row[7]:
                        continue
                    projectNum = row[7].split("Project Number:")[1].strip().split()[0]
                    row[7] = projectNum
                    filtered_row = [row[7]] + row[9:16]
                    writer.writerow(filtered_row)

                    if project_list is not None:
                        project_list.add(projectNum)
        
        os.replace(temp_file, csvName)
    except Exception as e:
        print("⚠️", e)
    finally:
        pass


def search_options(project_list):
    name = f"Rev_{datetime.now().strftime('%Y-%m-%d_%H_%M')}"
    pkey =  hashlib.md5(name.encode('utf-8')).hexdigest()
    try:
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/SaveSearchOptions/"
        
        savedOptionsDetail = []
        for projectNum in project_list:
            savedOptionsDetail.append({"Seq":1,"ParentKey":pkey,"OptionName":"WBS1","Type":"wbs1","Operator":"=","Value":projectNum,"ValueDescription":"","ReportOption":"N","Condition":"and","TableName":"PR","CrossHubField":" ","CrossHubFieldType":None,"SearchLevel":1,"PKey":"","_originalValues":{},"_transType":"I"})
        
        savedOptionsDetail.append({"Seq":-100,"ParentKey":pkey,"OptionName":"saveOptionRole","Type":"role","Operator":"=","Value":"[CREATOR_USERNAME]","ValueDescription":"Myself","ReportOption":"N","Condition":"","TableName":"","CrossHubField":" ","CrossHubFieldType":None,"SearchLevel":0,"PKey":"","_originalValues":{},"_transType":"I"})
  
        payload = {"Name":name,"Type":"wbs1","Private":"Y","Folder":"","LinkedPKey":"","WhereClauseSearch":"N","ResultsToDisplay":"","ListViewDisplay":"","SavedOptionsDetail":savedOptionsDetail, "PKey":pkey,"Username":""}
        response = requests.post(url, headers=HEADERS, json=payload)

        return pkey, name

    except Exception as e:
        print(f"Error in save search_options: {e}")


def generate_search_option(base_folder, start_period, end_period):
    print("Step 1: Generate the project list for the search options")
    download_invoice_YTD()
    print(f"Checking invoice register, now total {len(project_list)} projects touched.")

    download_GL(base_folder=base_folder, startPeriod = start_period, endPeriod = end_period, needDownload=False, baseRecordSelection={"pKey":None,"name":"Records Selected","type":"CA","whereClauseSearch":"N","isLegacy":"N","searchOptions":[{"name":"Account","value":"4","type":"account","seq":1,"tableName":"CA","opp":"startsWith","condition":"or","searchLevel":0,"valueDescription":"4"},{"name":"Account","value":"5","type":"account","seq":2,"tableName":"CA","opp":"startsWith","condition":"and","searchLevel":0,"valueDescription":"5"}]})
    print(f"Checking GL with 4*** and 5***, now total {len(project_list)} projects touched.")
    download_labour_details_YTD()
    print(f"Checking Labour hours, now total {len(project_list)} projects touched.")
    download_vp_revgen()
    print(f"Checking VantagePoint Revgen, now total {len(project_list)} projects touched.\n")

    print(f"unique project number: {len(project_list)}")

    pkey, option_name = search_options(project_list)

    print(f"{option_name} {pkey}")

    return pkey, option_name


def download_serch_options_list(pkey):
    base_folder = Path.home() / "Downloads"

    url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/SaveSearchOptions/{pkey}"
    response = requests.get(url, headers=HEADERS)
    option_name = response.json()[0]["Name"]

    url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/project/?lookuptype=wbs1&searchType=ALL&pagesize=1000&offset=0&page=1&isLevelLock=false&order=name&applicationId=&excludeSelectedResultIdsOption=true&savedSearchPKey={pkey}&timeout=Project_Search&WBSType=WBS1&AccessGroupBy=WBS1"
    data = requests.get(url, headers=HEADERS).json()

    csvName = os.path.join(base_folder, f"search_option_{option_name}.csv")
    
    with open(csvName, 'w', encoding='utf-8', newline='') as f_out:
        writer = csv.writer(f_out)
        writer.writerow(["Q-Number", "ProjectName"])
        for item in data:
            projectNum = item["key"]
            projectName = item["Name"]
            writer.writerow([projectNum, projectName])

    print(f"✅ Search option result downloaded: {csvName}")


def create_new_search_option(period, base_folder):
    print("--------------------------------------")
    print("Step 1: Generate the project list for the search options")
    download_invoice_YTD()
    print(f"Checking invoice register, now total {len(project_list)} projects touched.")
    download_GL(base_folder=base_folder, startPeriod = '202601', endPeriod = period, needDownload=False, baseRecordSelection={"pKey":None,"name":"Records Selected","type":"CA","whereClauseSearch":"N","isLegacy":"N","searchOptions":[{"name":"Account","value":"4","type":"account","seq":1,"tableName":"CA","opp":"startsWith","condition":"or","searchLevel":0,"valueDescription":"4"},{"name":"Account","value":"5","type":"account","seq":2,"tableName":"CA","opp":"startsWith","condition":"and","searchLevel":0,"valueDescription":"5"}]})
    print(f"Checking GL with 4*** and 5***, now total {len(project_list)} projects touched.")
    download_labour_details_YTD()
    print(f"Checking Labour hours, now total {len(project_list)} projects touched.")
    download_vp_revgen()
    print(f"Checking VantagePoint Revgen, now total {len(project_list)} projects touched.\n")

    print(f"unique project number: {len(project_list)}")
    pkey, option_name = search_options(project_list)
    return pkey, option_name


def download_total_project_list():
    try:
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        baseRecordSelection = {"pKey":None,"name":"Records Selected","type":"wbs1","whereClauseSearch":"N","isLegacy":"N","searchOptions":[{"name":"ChargeType","value":"R","type":"dropdown","seq":1,"tableName":"PR","condition":"and","searchLevel":1,"valueDescription":"Regular"}]}
        payload = {"reportPath":"/Standard/Project/Project List","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseChart3D":"N","baseChartColumn":"","baseChartDivisor":"1","baseChartFontSize":8,"baseChartHeight":3,"baseChartLabelLines":"N","baseChartLabels":"none","baseChartLeft":1,"baseChartLegendPosition":"righttop","baseChartSeriesColumn2":"","baseChartSeriesColumn3":"","baseChartShowPosition":"1","baseChartTitle":"","baseChartTop":0.5,"baseChartType":"none","baseChartWidth":6,"baseChartXTitle":"Project Number","baseChartYTitle":"","baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"Activity","baseGroupIndent":0.1,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Project List","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":0.1,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"Project Number","sort":"ASC","color":"000080","subTotal":"N","showHeading":"N","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"projectNumber","customGridColumnSort":"","groupWBSLevel":"1"}],"ReportColumns":[{"heading":"Project","width":1.25,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"WBS1","username":"","customGridColumnSort":""},{"heading":"Status","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"StatusDesc","username":"","customGridColumnSort":""},{"heading":"Biller","width":1.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"billerName","username":"","customGridColumnSort":""},{"heading":"Name","width":2.2,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Name","username":"","customGridColumnSort":""},{"heading":"Primary Client","width":1.8,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ClientName","username":"","customGridColumnSort":""},{"heading":"Project Manager","width":1.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"prgName","username":"","customGridColumnSort":""},{"heading":"Principal","width":1.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"prinName","username":"","customGridColumnSort":""},{"heading":"Charge Type","width":0.55,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ChargeType","username":"","customGridColumnSort":""},{"heading":"USD Project","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustUSDProject","username":"","customGridColumnSort":"N"},{"heading":"Large Projects","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustMilestoneBillingProjects400k","username":"","customGridColumnSort":"N"},{"heading":"Materials Only","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustMaterialsOnly","username":"","customGridColumnSort":"N"},{"heading":"Time & Material","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustTimeMaterial","username":"","customGridColumnSort":"N"},{"heading":"Long-term Contract","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustLongtermFixedPrice","username":"","customGridColumnSort":"N"},{"heading":"No Revenue","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustNoRevenue","username":"","customGridColumnSort":"N"},{"heading":"Fixed Fee","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustLumpSumBilling","username":"","customGridColumnSort":"N"},{"heading":"Defined Scope/deliverables","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustDefinedScopeValue","username":"","customGridColumnSort":"N"}],"ReportSections":[],"baseRecordSelection":baseRecordSelection,"baseCreateActivity":"N","baseShowDetail":"Y","CurrentWBSActivityActiveWBS1Only":"N","CurrentWBSActivityActiveWBS2Only":"N","CurrentWBSActivityActiveWBS3Only":"N","CurrentWBSActivityActivityRange":"1","CurrentWBSActivityInclInvoiceActivity":"N","CurrentWBSActivityCheckLabor":"Y","CurrentWBSActivityCheckExpense":"Y","_desc_saveOptionRole":["","","","",""],"saveOptionRole":["[CREATOR_USERNAME]","CONTROLLER-RO","ACCOUNTANT","CONTROLLER","ACCOUNTING"],"baseOriginalFavoriteId":"6af3e2e5be034e1cbebec9fda1d74c98","baseSelectionRows":10430}}

        response = requests.post(url, headers=HEADERS, json=payload  )
        data = response.json()
        report_path_raw = data["return"]["ReportPath"]
        report_path = report_path_raw.replace(" ", "%20")

        # Step 2: Get Nonce
        nonceUrl = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Security/Nonce"
        payload = {}
        response = requests.post(nonceUrl, headers=HEADERS, json=payload  )
        nonce = response.json()

        # Step 3: Get Viewer
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/Project/Project%20List&reportName=Project%20List"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("Error")
        
        exportFileName_csv = "Project List_"+ date.today().strftime("%Y-%m-%d") + ".csv"

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer"
                "&OpType=Export&FileName=Project+List&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        response = requests.get(url, headers = HEADERS,stream=True  )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(r"C:\\temp\\revenue_accrual", exportFileName_csv)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            print("✅ "+ csvName+" Downloaded")

        # Trim the csv file
        temp_file = csvName + ".temp"
        lines_to_skip = 4

        with open(csvName, 'r', encoding='utf-8', newline='') as f_in, \
            open(temp_file, 'w', encoding='utf-8', newline='') as f_out:
            reader = csv.reader(f_in)
            writer = csv.writer(f_out)
            for _ in range(lines_to_skip):
                next(reader, None)
            for row in reader:
                writer.writerow(row)

        os.replace(temp_file, csvName)
        
        return os.path.abspath(csvName)

    except Exception as e:
        print("⚠️ Failed to update the project list:", e)


def download_project_list(base_folder, need_trim = False):
    try:
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/Project/Project List","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseChart3D":"N","baseChartColumn":"","baseChartDivisor":"1","baseChartFontSize":8,"baseChartHeight":3,"baseChartLabelLines":"N","baseChartLabels":"none","baseChartLeft":1,"baseChartLegendPosition":"righttop","baseChartSeriesColumn2":"","baseChartSeriesColumn3":"","baseChartShowPosition":"1","baseChartTitle":"","baseChartTop":0.5,"baseChartType":"none","baseChartWidth":6,"baseChartXTitle":"Project Number","baseChartYTitle":"","baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"Activity","baseGroupIndent":0.1,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Project List","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":0.1,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"Project Number","sort":"ASC","color":"000080","subTotal":"N","showHeading":"N","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"projectNumber","customGridColumnSort":"","groupWBSLevel":"1"}],"ReportColumns":[{"heading":"Project","width":1.25,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"WBS1","username":"","customGridColumnSort":""},{"heading":"Name","width":2.2,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Name","username":"","customGridColumnSort":""},{"heading":"Primary Client","width":1.8,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ClientName","username":"","customGridColumnSort":""},{"heading":"Charge Type","width":0.55,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ChargeType","username":"","customGridColumnSort":""},{"heading":"Defined Scope/deliverables","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustDefinedScopeValue","username":"","customGridColumnSort":"N"},{"heading":"Biller","width":1.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"billerName","username":"","customGridColumnSort":""},{"heading":"Principal","width":1.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"prinName","username":"","customGridColumnSort":""},{"heading":"Project Manager","width":1.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"prgName","username":"","customGridColumnSort":""},{"heading":"USD Project","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustUSDProject","username":"","customGridColumnSort":"N"},{"heading":"Project Country","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"PRCountry","username":"","customGridColumnSort":""},{"heading":"Industry","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustIndustry","username":"","customGridColumnSort":"N"},{"heading":"Status","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"StatusDesc","username":"","customGridColumnSort":""},{"heading":"Large Projects","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustMilestoneBillingProjects400k","username":"","customGridColumnSort":"N"},{"heading":"Materials Only","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustMaterialsOnly","username":"","customGridColumnSort":"N"},{"heading":"Time & Material","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustTimeMaterial","username":"","customGridColumnSort":"N"},{"heading":"Long-term Contract","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustLongtermFixedPrice","username":"","customGridColumnSort":"N"},{"heading":"No Revenue","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustNoRevenue","username":"","customGridColumnSort":"N"},{"heading":"Fixed Fee","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustLumpSumBilling","username":"","customGridColumnSort":"N"}],"ReportSections":[],"baseRecordSelection":{"pKey":"","name":"Records Selected","type":"wbs1","whereClauseSearch":"N","isLegacy":"N","searchOptions":[{"name":"ChargeType","value":"[IS_EMPTY]","type":"dropdown","seq":1,"tableName":"PR","opp":"!=","condition":"and","searchLevel":1,"valueDescription":""}]},"baseCreateActivity":"N","baseShowDetail":"Y","CurrentWBSActivityActiveWBS1Only":"N","CurrentWBSActivityActiveWBS2Only":"N","CurrentWBSActivityActiveWBS3Only":"N","CurrentWBSActivityActivityRange":"1","CurrentWBSActivityInclInvoiceActivity":"N","CurrentWBSActivityCheckLabor":"Y","CurrentWBSActivityCheckExpense":"Y","_desc_saveOptionRole":["","","","",""],"saveOptionRole":["CONTROLLER","CONTROLLER-RO","[CREATOR_USERNAME]","ACCOUNTING","ACCOUNTANT"],"baseOriginalFavoriteId":"","baseSelectionRows":1}}

        response = requests.post(url, headers=HEADERS, json=payload  )
        data = response.json()
        report_path_raw = data["return"]["ReportPath"]
        report_path = report_path_raw.replace(" ", "%20")

        # Step 2: Get Nonce
        nonceUrl = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Security/Nonce"
        payload = {}
        response = requests.post(nonceUrl, headers=HEADERS, json=payload  )
        nonce = response.json()

        # Step 3: Get Viewer
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/Project/Project%20List&reportName=Project%20List"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("Error")
        
        exportFileName_csv = "Project List_"+ date.today().strftime("%Y-%m-%d") + ".csv"

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer"
                "&OpType=Export&FileName=Project+List&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        response = requests.get(url, headers = HEADERS,stream=True  )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(base_folder, exportFileName_csv)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            print("✅ "+ csvName+" Downloaded")
        
        if need_trim:
            # Trim the csv file used for copy paste
            temp_file = csvName + ".temp"
            lines_to_skip = 4

            with open(csvName, 'r', encoding='utf-8', newline='') as f_in, \
                open(temp_file, 'w', encoding='utf-8', newline='') as f_out:
                reader = csv.reader(f_in)
                writer = csv.writer(f_out)
                for _ in range(lines_to_skip):
                    next(reader, None)
                for row in reader:
                    writer.writerow(row)

            os.replace(temp_file, csvName)
        
        return os.path.abspath(csvName)

    except Exception as e:
        print("⚠️ Failed to update the project list:", e)


def download_estimate(base_folder):
    try:
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/Project/Project Earnings","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseChart3D":"N","baseChartColumn":"UDCol_CustLaborEstimate","baseChartDivisor":"1","baseChartFontSize":8,"baseChartHeight":3,"baseChartLabelLines":"N","baseChartLabels":"none","baseChartLeft":1,"baseChartLegendPosition":"righttop","baseChartSeriesColumn2":"","baseChartSeriesColumn3":"","baseChartShowPosition":"1","baseChartTitle":"","baseChartTop":0.5,"baseChartType":"none","baseChartWidth":6,"baseChartXTitle":"Project Number","baseChartYTitle":"Labor Estimate (No. of Hours)","baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.14,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Estimate","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":2.85,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"Project Number","sort":"ASC","color":"000080","subTotal":"Y","showHeading":"Y","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"projectNumber","customGridColumnSort":"","groupWBSLevel":"1"}],"ReportColumns":[{"heading":"Charge Type","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"chargeType","username":"","customGridColumnSort":""},{"heading":"Project Manager","width":1.25,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"prgName","username":"","customGridColumnSort":""},{"heading":"Labor Estimate (No. of Hours)","width":0.85,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustLaborEstimate","username":"","customGridColumnSort":"N"},{"heading":"HW Estimate (Cost in CAD)","width":0.85,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustHWEstimate","username":"","customGridColumnSort":"N"},{"heading":"SW Estimate (Cost in CAD)","width":0.85,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustSWEstimate","username":"","customGridColumnSort":"N"},{"heading":"Outside Services Estimate (Cost in CAD)","width":0.85,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustOutsideServicesEstimate","username":"","customGridColumnSort":"N"},{"heading":"Other Expenses Estimate (Cost in CAD)","width":0.85,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustOtherExpensesEstimate","username":"","customGridColumnSort":"N"},{"heading":"Defined Scope/deliverables","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustDefinedScopeValue","username":"","customGridColumnSort":"N"}],"ReportSections":[],"baseRecordSelection":"{\"pKey\":null,\"name\":\"Records Selected\",\"type\":\"wbs1\",\"whereClauseSearch\":\"N\",\"isLegacy\":\"N\",\"searchOptions\":[{\"name\":\"CustDefinedScopeValue\",\"value\":\"Y\",\"type\":\"checkbox\",\"seq\":1,\"tableName\":\"PR\",\"condition\":\"and\",\"searchLevel\":1,\"valueDescription\":\"\"}]}","baseCreateActivity":"N","baseShowDetail":"Y","baseLeft1":0,"baseRight1":21,"baseLeft2":0,"baseRight2":8,"baseLeft3":0,"baseRight3":8,"baseSub":"1","rollType":"Project","CurrentWBSActivityActiveWBS1Only":"N","CurrentWBSActivityActiveWBS2Only":"N","CurrentWBSActivityActiveWBS3Only":"N","CurrentWBSActivityActivityRange":"1","CurrentWBSActivityInclInvoiceActivity":"N","ReportFormat":"VOS","includeType":"All","hoursandAmounts":"1","labDrillDown":"1","expDrillDown":"1","atCost":"1","useSummaryTable":"N","CurrentWBSActivityCheckLabor":"Y","CurrentWBSActivityCheckExpense":"Y","CurrentWBSUnpostsedLabor":"Y","budgetSelection":"1","ETCDRadioChecked":"radioETCD1","showOverhead":"Y","estimateOverhead":"Y","currencyType":"project","targetDate":"6/6/2025 2:24:25 PM","currencyTypeARFee":"billing","currencyTypeBill":"billing","startPeriod":"03/2026","endPeriod":"03/2026","SummaryTableLastUpdate":"","_desc_saveOptionRole":["","",""],"saveOptionRole":["ACCOUNTANT","ACCOUNTING","[CREATOR_USERNAME]"],"baseOriginalFavoriteId":"","baseSelectionRows":0}}

        response = requests.post(url, headers=HEADERS, json=payload  )
        data = response.json()
        report_path_raw = data["return"]["ReportPath"]
        report_path = report_path_raw.replace(" ", "%20")

        # Step 2: Get Nonce
        nonceUrl = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Security/Nonce"
        payload = {}
        response = requests.post(nonceUrl, headers=HEADERS, json=payload  )
        nonce = response.json()

        # Step 3: Get Viewer
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/Project/Project%20Earnings&reportName=Estimate"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("Error")
        
        exportFileName_csv = "Estimate_"+ date.today().strftime("%Y-%m-%d") + ".csv"

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer&OpType=Export&FileName=Estimate&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        response = requests.get(url, headers = HEADERS, stream=True )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(base_folder, exportFileName_csv)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            print("✅ "+ csvName+" Downloaded")

    except Exception as e:
        print("⚠️ Failed to update the project list:", e)


def init(base_folder):
    LOGIN = util.check_login()
    while not LOGIN:
        login.sso_login()
        LOGIN = util.check_login()
    global HEADERS
    HEADERS = util.set_headers()

    if not Path.exists(base_folder):
        os.makedirs(base_folder)
    else:
        shutil.rmtree(base_folder)
        os.makedirs(base_folder, exist_ok=True)


def update_rate(base_folder, rate, file_name):
    pythoncom.CoInitialize()
    app = None
    wb = None
    target_file = os.path.join(base_folder, file_name)
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        for _ in range(3):
            try:
                app.Visible = False
                app.DisplayAlerts = False
                break
            except:
                time.sleep(1)
        max_retries = 5
        for i in range(max_retries):
            try:
                wb = app.Workbooks.Open(target_file)
                break
            except Exception as e:
                if i < max_retries - 1:
                    print(f"Excel busy, retrying in 10s... ({i+1}/{max_retries})")
                    time.sleep(10)
                else:
                    raise e
        if wb is not None:
            names = wb.Names
            found = False
            for i in range(1, names.Count + 1):
                try:
                    n = names.Item(i)
                    if n.Name == "USRATE":
                        target_range = n.RefersToRange
                        if target_range is not None:
                            target_range.Value = rate
                            found = True
                            break
                except:
                    continue
            if found:
                wb.Save()
                print(f"Successfully updated the rate in {file_name}")
            else:
                print(f"Warning: 'USRATE' not found in {file_name}")

    except Exception as e:
        print(f"Critical error processing {file_name}: {e}")
    
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=True)
            except:
                pass
        if app is not None:
            try:
                app.Quit()
            except:
                pass
        del wb
        del app
        pythoncom.CoUninitialize()


def download_contract(base_folder, end_period, pkey = "", option_name = ""):
    try:
        download_template(base_folder, TEMPLATE2)
        targetFile = os.path.join(base_folder, TEMPLATE2)

        print("Fetch the USD/CAD rate")
        rate = util.get_currency_rate(end_period, "USD")
        print(f"USD/CAD rate for {end_period}: {rate}")
        update_rate(base_folder, rate, TEMPLATE2)

        csvName = download_project_list(base_folder)
        util.excel_full_copy(inputFile=csvName, inputSheet=None, targetFile=targetFile, targetSheet="Project List", onlyValue=True, targetCell='A2')

        os.remove(csvName)        

        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        baseRecordSelection = {"pKey": pkey ,"name":option_name ,"type":"wbs1","whereClauseSearch":"N","isLegacy":"N"}
        payload = {"reportPath":"/Standard/Project/Contract Management","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.2,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"Y","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Contract Management","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"Y","baseStartColumnPosition":3.2,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"","sort":"ASC","color":"000080","subTotal":"Y","showHeading":"Y","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"projectNumber","customGridColumnSort":"","groupWBSLevel":"1"}],"ReportColumns":[{"heading":"Contract  Direct Labour","width":0.95,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ContractFeeDirectLab","username":"","customGridColumnSort":""},{"heading":"Contract  Direct Expense","width":0.95001,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ContractFeeDirectExp","username":"","customGridColumnSort":""},{"heading":"Contract Direct Consultant","width":0.95001,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ContractConsultFee","username":"","customGridColumnSort":""},{"heading":"Contract Reimb Allow Exp","width":1,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ContractReimbAllowExp","username":"","customGridColumnSort":""},{"heading":"Contract Reimb Allow Consultant","width":1,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ContractReimbAllowCons","username":"","customGridColumnSort":""},{"heading":"Total Contract","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ContractTotal","username":"","customGridColumnSort":""},{"heading":"Contract Number","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ContractNumber","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":baseRecordSelection,"baseCreateActivity":"N","baseShowDetail":"Y","radReportingPeriod":"1","_desc_saveOptionRole":["","","","",""],"saveOptionRole":["CONTROLLER","CONTROLLER-RO","ACCOUNTING","ACCOUNTANT","[CREATOR_USERNAME]"],"baseOriginalFavoriteId":"","baseSelectionRows":0}}

        response = requests.post(url, headers=HEADERS, json=payload  )
        data = response.json()
        report_path_raw = data["return"]["ReportPath"]
        report_path = report_path_raw.replace(" ", "%20")

        # Step 2: Get Nonce
        nonceUrl = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Security/Nonce"
        payload = {}
        response = requests.post(nonceUrl, headers=HEADERS, json=payload  )
        nonce = response.json()

        # Step 3: Get Viewer
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/Project/Contract%20Management&reportName="

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("Error")
        
        exportFileName_csv = "Contract Manger_"+ date.today().strftime("%Y-%m-%d") + ".csv"

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer"
                "&OpType=Export&FileName=Contract+Management&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        response = requests.get(url, headers = HEADERS,stream=True )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(base_folder, exportFileName_csv)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            print("✅ "+ csvName+" Downloaded")
        
        # Trim the csv file
        temp_file = csvName + ".temp"
        lines_to_skip = 4

        with open(csvName, 'r', encoding='utf-8', newline='') as f_in, \
            open(temp_file, 'w', encoding='utf-8', newline='') as f_out:
            reader = csv.reader(f_in)
            writer = csv.writer(f_out)
            for _ in range(lines_to_skip):
                next(reader, None)
            for row in reader:
                writer.writerow(row)

        os.replace(temp_file, csvName)

        
        util.excel_full_copy(inputFile=csvName, inputSheet=None, targetFile=targetFile, targetSheet="Contract export", onlyValue=True, targetCell='B3')

        os.remove(csvName)
    except Exception as e:
        print("⚠️ Failed to download the contract:", e)


def download_labour_details_JTD(base_folder, pkey = "", option_name = ""):
    try:
        # Step 1: Build
        baseRecordSelection = {"pKey": pkey ,"name":option_name ,"type":"wbs1","whereClauseSearch":"N","isLegacy":"N"}
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/Project/Labor Detail","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseChart3D":"N","baseChartColumn":"regHrs","baseChartDivisor":"1","baseChartFontSize":8,"baseChartHeight":3,"baseChartLabelLines":"N","baseChartLabels":"none","baseChartLeft":1,"baseChartLegendPosition":"righttop","baseChartSeriesColumn2":"","baseChartSeriesColumn3":"","baseChartShowPosition":"1","baseChartTitle":"","baseChartTop":0.5,"baseChartType":"none","baseChartWidth":6,"baseChartXTitle":"Project Number","baseChartYTitle":"Regular Hours","baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"yyyy-MM-dd","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.05,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"Y","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Labour Detail for Rev_Gen","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":3.5,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"Project Number","sort":"ASC","color":"000080","subTotal":"N","showHeading":"N","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"projectNumber","customGridColumnSort":"","groupWBSLevel":"1"},{"label":"Task Number","sort":"ASC","color":"228B22","subTotal":"N","showHeading":"N","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"wbs2Number","customGridColumnSort":"","groupWBSLevel":"2"},{"label":"Subtask Number","sort":"ASC","color":"000000","subTotal":"N","showHeading":"N","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"wbs3Number","customGridColumnSort":"","groupWBSLevel":"3"},{"label":"Labour Code Level 1","sort":"ASC","color":"000000","subTotal":"N","showHeading":"N","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"labcd1","customGridColumnSort":"","groupWBSLevel":"1"}],"ReportColumns":[{"heading":"Project","width":1.25,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"WBS1","username":"","customGridColumnSort":""},{"heading":"Task","width":0.8,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"WBS2","username":"","customGridColumnSort":""},{"heading":"Subtask","width":0.8,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"WBS3","username":"","customGridColumnSort":""},{"heading":"Date","width":0.75,"format":"yyyy-MM-dd","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"transDate","username":"","customGridColumnSort":""},{"heading":"Regular Hours","width":0.55,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"regHrs","username":"","customGridColumnSort":""},{"heading":"Regular Billing","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"regAmt","username":"","customGridColumnSort":""},{"heading":"Overtime Hours","width":0.55,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ovtHrs","username":"","customGridColumnSort":""},{"heading":"Overtime Billing","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ovtAmt","username":"","customGridColumnSort":""},{"heading":"Ovt-2 Hours","width":0.55,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"specialOvtHrs","username":"","customGridColumnSort":""},{"heading":"Ovt-2 Billing","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"specialOvtAmt","username":"","customGridColumnSort":""},{"heading":"Total Hours","width":0.55,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"totHrs","username":"","customGridColumnSort":""},{"heading":"Total Billing","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"totAmt","username":"","customGridColumnSort":""},{"heading":"Total Ovt Hrs","width":0.55,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"totOvtHrs","username":"","customGridColumnSort":""},{"heading":"Tot. Ovt Billing","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"totOvtAmt","username":"","customGridColumnSort":""},{"heading":"Bill Status","width":0.55,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"billStatus","username":"","customGridColumnSort":""},{"heading":"Period","width":0.55,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"period","username":"","customGridColumnSort":""},{"heading":"Post Seq","width":0.55,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"postSeq","username":"","customGridColumnSort":""},{"heading":"Charge Type","width":0.55,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"chargeType","username":"","customGridColumnSort":""},{"heading":"Trans Type","width":0.55,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"transType","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":baseRecordSelection,"baseCreateActivity":"N","baseShowDetail":"Y","baseLeft1":0,"baseRight1":21,"baseLeft2":0,"baseRight2":8,"baseLeft3":0,"baseRight3":8,"baseSub":"1","rollType":"Project","CurrentWBSActivityActiveWBS1Only":"N","CurrentWBSActivityActiveWBS2Only":"N","CurrentWBSActivityActiveWBS3Only":"N","CurrentWBSActivityActivityRange":"1","CurrentWBSActivityInclInvoiceActivity":"N","timeframeBackup":"JTD","timeFrame":"radio1","periodType":"3","transDetail":"1","curSelection":"2","showComments":"N","showUnposted":"N","atCost":"2","CurrentWBSActivityCheckLabor":"Y","CurrentWBSActivityCheckExpense":"Y","CurrentWBSActivityUnpostedLabor":"N","employeeWhere":"","emOwnerfl":"","dateStart":"2026-04-01T00:00:00","dateEnd":"2026-04-01T00:00:00","periodStart":"202601","periodEnd":"202604","baseSelectionRows":1,"baseOriginalFavoriteId":"","_desc_saveOptionRole":["","","",""],"saveOptionRole":["ACCOUNTING","ACCOUNTANT","[CREATOR_USERNAME]","CONTROLLER"],"LaborPostingRun":[]}}

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
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/Project/Labor%20Detail&reportName="

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("⚠️ Failed to download the YTD invoices")

        exportFileName_csv = "Labour Details JTD_"+ date.today().strftime("%Y-%m-%d") + ".csv"

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer"
                "&OpType=Export&FileName=Jay_Labour+Detail_Rev&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        response = requests.get(url, headers = HEADERS,stream=True )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(base_folder, exportFileName_csv)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            print("✅ "+ csvName+" Downloaded")

    except Exception as e:
        print("⚠️ Failed to download the labour details:", e) 


def download_cost_GL(base_folder, end_period):
    try:
        # Step 1: Build
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/DataExport/GeneralLedgerDataSource","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"General Ledger Export","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":0,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[],"ReportColumns":[{"heading":"Account Numbers g","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Account","username":"","customGridColumnSort":""},{"heading":"Account Code","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Account_Code","username":"","customGridColumnSort":""},{"heading":"Account Description","width":1.375,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Account_Description","username":"","customGridColumnSort":""},{"heading":"Trans Type","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Transaction_Type","username":"","customGridColumnSort":""},{"heading":"Trans SubType","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Transaction_SubType","username":"","customGridColumnSort":""},{"heading":"Date","width":0.875,"format":"M/d/yyyy","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Transaction_Date","username":"","customGridColumnSort":""},{"heading":"Ref. No","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Reference_Number","username":"","customGridColumnSort":""},{"heading":"Period","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Period","username":"","customGridColumnSort":""},{"heading":"Post Seq","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Posting_Sequence","username":"","customGridColumnSort":""},{"heading":"Description 1","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Description_1","username":"","customGridColumnSort":""},{"heading":"Description 2","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Description_2","username":"","customGridColumnSort":""},{"heading":"Project","width":1.25,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"wbs1","username":"","customGridColumnSort":""},{"heading":"Task","width":0.875,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"wbs2","username":"","customGridColumnSort":""},{"heading":"Subtask","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"wbs3","username":"","customGridColumnSort":""},{"heading":"Amount","width":1,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"amount","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":{"pKey":"e84aeb53bd854183b3c38b91e23bbecd","name":"5*** Accounts","type":"CA","whereClauseSearch":"N","isLegacy":"N"},"timeframe":"PeriodRange","ReportingPeriod":"1","W2Year":"","W2Qtr":"","W2Quarter":"","startPeriod":201901,"endPeriod":end_period,"PeriodStart":201901,"PeriodEnd":end_period,"baseSelectionRows":42,"_desc_saveOptionRole":["",""],"saveOptionRole":["ACCOUNTANT","[CREATOR_USERNAME]"],"baseOriginalFavoriteId":"1C4980783B8643039DF4DD0EADCE26AD"}}

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
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/DataExport/GeneralLedgerDataSource&reportName=General%20Ledger%20Export"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("⚠️ Failed to download the JTD cost details")

        exportFileName_csv = "Cost Details JTD_"+ date.today().strftime("%Y-%m-%d") + ".csv"

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer&OpType=Export&FileName=General+Ledger+Export&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        response = requests.get(url, headers = HEADERS,stream=True )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(base_folder, exportFileName_csv)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            print("✅ "+ csvName+" Downloaded")

    except Exception as e:
        print("⚠️ Failed to download the cost details JTD:", e)


def download_project_labor_hrs(base_folder, pkey, option_name):
    try:
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        baseRecordSelection = {"pKey": pkey ,"name":option_name ,"type":"wbs1","whereClauseSearch":"N","isLegacy":"N"}
        payload = {"reportPath":"/Standard/DataExport/ProjectLaborDataSource","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[username] - [options]","baseGridTable":"","baseGroupIndent":0,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Project Labour Export","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":0,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[],"ReportColumns":[{"heading":"Project","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Project","username":"","customGridColumnSort":""},{"heading":"Task","width":0.625,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Phase","username":"","customGridColumnSort":""},{"heading":"Subtask","width":0.625,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Task","username":"","customGridColumnSort":""},{"heading":"Employee Number","width":0.875,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Employee","username":"","customGridColumnSort":""},{"heading":"Employee Full Name","width":1.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Employee_Name","username":"","customGridColumnSort":""},{"heading":"Date","width":0.75,"format":"M/d/yyyy","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_TransDate","username":"","customGridColumnSort":""},{"heading":"Regular Hours","width":0.625,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Regular_Hours","username":"","customGridColumnSort":""},{"heading":"Overtime Hours","width":0.75,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Overtime_Hours","username":"","customGridColumnSort":""},{"heading":"Special Overtime Hours","width":1,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Special_Overtime_Hours","username":"","customGridColumnSort":""},{"heading":"Status","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Status","username":"","customGridColumnSort":""},{"heading":"Bill Status","width":0.875,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_BillStatus","username":"","customGridColumnSort":""},{"heading":"Period","width":0.55,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Period","username":"","customGridColumnSort":""},{"heading":"Post Seq","width":0.55,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_PostSeq","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":baseRecordSelection,"ReportingPeriod":"1","W2Year":"","W2Qtr":"","W2Quarter":"","baseSelectionRows":1,"_desc_saveOptionRole":["","","","",""],"saveOptionRole":["CONTROLLER-RO","[CREATOR_USERNAME]","CONTROLLER","ACCOUNTING","ACCOUNTANT"],"baseOriginalFavoriteId":""}}

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
        url = f"https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce={nonce}&ResetReportViewerOnPreview=Y&reportPath={report_path}&allowSchedule=Y&origReportPath=/Standard/DataExport/ProjectLaborDataSource&reportName=Project%20Labour%20Export"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("⚠️ Failed to download the JTD cost details")

        exportFileName_csv = "Project_Labour_Hours_"+ date.today().strftime("%Y-%m-%d") + ".csv"

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer&OpType=Export&FileName=Project+Labour+Export&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        response = requests.get(url, headers = HEADERS,stream=True )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(base_folder, exportFileName_csv)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            print("✅ "+ csvName+" Downloaded")
        
    except Exception as e:
        print("⚠️ Failed to download the project labor hours:", e)


def download_project_expense(base_folder, pkey, option_name):
    try:
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        baseRecordSelection = {"pKey": pkey ,"name":option_name ,"type":"wbs1","whereClauseSearch":"N","isLegacy":"N"}
        payload = {"reportPath":"/Standard/DataExport/ProjectExpenseDataSource","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[username] - [options]","baseGridTable":"","baseGroupIndent":0,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Project Expense Export","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":0,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[],"ReportColumns":[{"heading":"Project","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Project","username":"","customGridColumnSort":""},{"heading":"Task","width":0.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Phase","username":"","customGridColumnSort":""},{"heading":"Subtask","width":0.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Task","username":"","customGridColumnSort":""},{"heading":"Expense Account","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_Account","username":"","customGridColumnSort":""},{"heading":"TransType","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_TransType","username":"","customGridColumnSort":""},{"heading":"Reference Number","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_RefNo","username":"","customGridColumnSort":""},{"heading":"Date","width":0.75,"format":"M/d/yyyy","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_transDate","username":"","customGridColumnSort":""},{"heading":"Cost Amount","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_costAmount","username":"","customGridColumnSort":""},{"heading":"Description 1","width":2,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_Desc1","username":"","customGridColumnSort":""},{"heading":"Status","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Status","username":"","customGridColumnSort":""},{"heading":"Bill Status","width":0.55,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_BillStatus","username":"","customGridColumnSort":""},{"heading":"Period","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_period","username":"","customGridColumnSort":""},{"heading":"Post Seq","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_PostSeq","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":baseRecordSelection,"ReportingPeriod":"1","W2Year":"","W2Qtr":"","W2Quarter":"","baseSelectionRows":595,"_desc_saveOptionRole":["","","","",""],"saveOptionRole":["ACCOUNTING","CONTROLLER","CONTROLLER-RO","[CREATOR_USERNAME]","ACCOUNTANT"],"baseOriginalFavoriteId":"bc8b936ed21b4d51940a3182450f4401"}}

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
        url = f"https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce={nonce}&ResetReportViewerOnPreview=Y&reportPath={report_path}&allowSchedule=Y&origReportPath=/Standard/DataExport/ProjectExpenseDataSource&reportName=Project%20Expense%20Export"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("⚠️ Failed to download the JTD cost details")

        exportFileName_csv = "Project_Expense_"+ date.today().strftime("%Y-%m-%d") + ".csv"

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer&OpType=Export&FileName=Project+Expense+Export&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        response = requests.get(url, headers = HEADERS,stream=True )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(base_folder, exportFileName_csv)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

        print("✅ "+ csvName+" Downloaded")
        
    except Exception as e:
        print("⚠️ Failed to download the project expense:", e)


def download_JTD_Billing(base_folder, pkey = "", option_name = "", end_period = ""):
    HEADERS = util.set_headers()
    
    download_template(base_folder, TEMPLATE1)
    targetFile = os.path.join(base_folder, TEMPLATE1)
    csvName = download_total_project_list()
    util.excel_full_copy(inputFile=csvName, inputSheet=None, targetFile=targetFile, targetSheet="Project List", onlyValue=True,  targetCell='A2')


    csvName = download_GL(base_folder=base_folder, startPeriod='200301', endPeriod=end_period, needDownload=True, baseRecordSelection={"pKey":None,"name":"Records Selected","type":"CA","whereClauseSearch":"N","isLegacy":"N","searchOptions":[{"name":"Name","value":"USD","type":"string","seq":1,"tableName":"CA","opp":"LIKE","condition":"or","searchLevel":0,"valueDescription":"USD"},{"name":"Name","value":"USA","type":"string","seq":2,"tableName":"CA","opp":"LIKE","condition":"and","searchLevel":0,"valueDescription":"USA"}]}, fileName="US ")

    util.excel_full_copy(inputFile=csvName, inputSheet=None, targetFile=targetFile, targetSheet="USD Revenue GL", onlyValue=True)
    
    try:
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        baseRecordSelection = {"pKey": pkey ,"name":option_name ,"type":"wbs1","whereClauseSearch":"N","isLegacy":"N"}
        payload = {"reportPath":"/Standard/Project/Project Detail","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseChart3D":"N","baseChartColumn":"TotHrs","baseChartDivisor":"1","baseChartFontSize":8,"baseChartHeight":3,"baseChartLabelLines":"N","baseChartLabels":"none","baseChartLeft":1,"baseChartLegendPosition":"righttop","baseChartSeriesColumn2":"","baseChartSeriesColumn3":"","baseChartShowPosition":"1","baseChartTitle":"","baseChartTop":0.5,"baseChartType":"none","baseChartWidth":6,"baseChartXTitle":"Project Number","baseChartYTitle":"Total Hours","baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.1,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"Y","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Project Detail","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"Y","baseShowTotalsOnHeader":"N","baseStartColumnPosition":3.2,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"Project Number","sort":"ASC","color":"000080","subTotal":"Y","showHeading":"Y","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"projectNumber","customGridColumnSort":"","groupWBSLevel":"1"},{"label":"Task Number","sort":"ASC","color":"228B22","subTotal":"Y","showHeading":"Y","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"wbs2Number","customGridColumnSort":"","groupWBSLevel":"2"},{"label":"Subtask Number","sort":"ASC","color":"000000","subTotal":"Y","showHeading":"Y","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"wbs3Number","customGridColumnSort":"","groupWBSLevel":"3"}],"ReportColumns":[{"heading":"Total Hours","width":0.7,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"TotHrs","username":"","customGridColumnSort":""},{"heading":"Billing","width":0.75,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"TotBill","username":"","customGridColumnSort":""},{"heading":"Effective Rate","width":1,"format":"###T###T###T###D##;(###T###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Effective Rate","header1":"Effective Rate","header2":"","detailExpression":"[TotBill]/[TotHrs]","groupExpression":"[TotBill]/[TotHrs]","queryJoin":"","checkSecurity":"N","altHeader1":"","altHeader2":"","queryColumn":"","calculatedColumnType":"ALLFRAMES","username":"BLIU","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection": baseRecordSelection,"baseCreateActivity":"N","baseShowDetail":"Y","baseLeft1":0,"baseRight1":21,"baseLeft2":0,"baseRight2":8,"baseLeft3":0,"baseRight3":8,"baseSub":"1","rollType":"Project","CurrentWBSActivityActiveWBS1Only":"N","CurrentWBSActivityActiveWBS2Only":"N","CurrentWBSActivityActiveWBS3Only":"N","CurrentWBSActivityActivityRange":"3","CurrentWBSActivityInclInvoiceActivity":"N","timeframeBackup":"JTD","timeFrame":"radio1","periodType":"3","CurrentWBSActivityCheckLabor":"N","CurrentWBSActivityCheckExpense":"Y","CurrentWBSActivityUnpostedLabor":"N","CurrentWBSActivityCommittedExp":"N","atCost":"1","FACost":"1","ProjectInfo":"N","transDetail":"1","curSelection":"2","PrintExpenses":"Y","PrintLabor":"Y","currencyTypeFA":"project","LaborDetail":"1","estimateOverhead":"N","PrintComments":"N","emplSubtot":"N","showUnposted":"N","PrintReims":"Y","PrintReimCons":"Y","PrintIndirects":"Y","PrintDirects":"Y","PrintDirectCons":"Y","Sortby2":"1","LCLevels":"1","LCSort1":"1","PrintRatesFlag":"Y","chkBillingStatus":"Y","ConsultantBreakout":"N","VendorInvoice":"Y","chkIncludeCommitPO":"N","_desc_saveOptionRole":["","","","",""],"saveOptionRole":["[CREATOR_USERNAME]","ACCOUNTANT","ACCOUNTING","CONTROLLER","CONTROLLER-RO"],"baseOriginalFavoriteId":"23406FD295DC4D979A49AB43F69556EF","LaborPostingRun":[]}}

        response = requests.post(url, headers=HEADERS, json=payload  )
        data = response.json()
        report_path_raw = data["return"]["ReportPath"]
        report_path = report_path_raw.replace(" ", "%20")

        # Step 2: Get Nonce
        nonceUrl = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Security/Nonce"
        payload = {}
        response = requests.post(nonceUrl, headers=HEADERS, json=payload  )
        nonce = response.json()

        # Step 3: Get Viewer
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/Project/Project%20Detail&reportName=Project%20Detail"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("Error")
        
        exportFileName_csv = "JTD Billing_"+ date.today().strftime("%Y-%m-%d") + ".csv"

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer"
                "&OpType=Export&FileName=Project+Detail&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        response = requests.get(url, headers = HEADERS,stream=True  )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(r"C:\\temp\\revenue_accrual", exportFileName_csv)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            print("✅ "+ csvName+" Downloaded")

        # Trim the csv file
        temp_file = csvName + ".temp"
        lines_to_skip = 4

        with open(csvName, 'r', encoding='utf-8', newline='') as f_in, \
            open(temp_file, 'w', encoding='utf-8', newline='') as f_out:
            reader = csv.reader(f_in)
            writer = csv.writer(f_out)
            for _ in range(lines_to_skip):
                next(reader, None)
            for row in reader:
                filtered_row = row[1:10]
                writer.writerow(filtered_row)

        os.replace(temp_file, csvName)
        
        util.excel_full_copy(inputFile=csvName, inputSheet=None, targetFile=targetFile, targetSheet="JTD Billing", onlyValue=True, targetCell='B6')

    except Exception as e:
        print("⚠️ Failed to download the JTD Billing:", e)


def download_invoice_pretax(base_folder, pkey = "", option_name = ""):
    try:
        baseRecordSelection = {"pKey":pkey,"name":option_name,"type":"wbs1","whereClauseSearch":"N","isLegacy":"N"}
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/AccountsReceivable/Invoice Register","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseChart3D":"N","baseChartColumn":"Col1","baseChartDivisor":"1","baseChartFontSize":8,"baseChartHeight":3,"baseChartLabelLines":"N","baseChartLabels":"none","baseChartLeft":1,"baseChartLegendPosition":"righttop","baseChartSeriesColumn2":"","baseChartSeriesColumn3":"","baseChartShowPosition":"1","baseChartTitle":"","baseChartTop":0.5,"baseChartType":"none","baseChartWidth":6,"baseChartXTitle":"Primary Client Name","baseChartYTitle":"Prof Fees","baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.1,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Invoice Summary","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"Y","baseStartColumnPosition":0.5,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"Primary Client Name","sort":"ASC","color":"000000","subTotal":"Y","showHeading":"Y","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"clientName","customGridColumnSort":"","groupWBSLevel":"1"},{"label":"Project Number","sort":"ASC","color":"000000","subTotal":"N","showHeading":"N","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"projectNumber","customGridColumnSort":"","groupWBSLevel":"1"}],"ReportColumns":[{"heading":"Project","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"WBS1","username":"","customGridColumnSort":""},{"heading":"Date","width":0.7,"format":"M/d/yyyy","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"TransDate","username":"","customGridColumnSort":""},{"heading":"Invoice","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"InvoiceNumber","username":"","customGridColumnSort":""},{"heading":"Prof Fees","width":0.9,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col1","username":"","customGridColumnSort":""},{"heading":"H/W Sales","width":0.9,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col2","username":"","customGridColumnSort":""},{"heading":"S/W Sales","width":0.9,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col3","username":"","customGridColumnSort":""},{"heading":"O/S Services","width":0.9,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col4","username":"","customGridColumnSort":""},{"heading":"Reimbursable","width":0.9,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col5","username":"","customGridColumnSort":""},{"heading":"Sales Pre-Tax","width":1,"format":"###T###T###T###D##;(###T###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Sales Pre-tax","header1":"Sales Pre-Tax","header2":"","detailExpression":"[Col1]+[Col2]+[Col3]+[Col4]+[Col5]","groupExpression":"[Col1]+[Col2]+[Col3]+[Col4]+[Col5]","queryJoin":"","checkSecurity":"N","altHeader1":"","altHeader2":"","queryColumn":"","calculatedColumnType":"ALLFRAMES","username":"BLIU","customGridColumnSort":""},{"heading":"Large Projects","width":0.85,"format":"","align":"center","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustMilestoneBillingProjects400k","username":"","customGridColumnSort":"N"},{"heading":"USD Project","width":0.85,"format":"","align":"center","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustUSDProject","username":"","customGridColumnSort":"N"},{"heading":"Lump Sum Billing","width":0.85,"format":"","align":"center","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustLumpSumBilling","username":"","customGridColumnSort":"N"}],"ReportSections":[],"baseRecordSelection":baseRecordSelection,"baseShowDetail":"Y","baseLeft1":0,"baseRight1":21,"baseLeft2":0,"baseRight2":8,"baseLeft3":0,"baseRight3":8,"baseSub":"1","rollType":"Project","timeframe":"JTD","radioTF":"radio1","tfCYJ":"J","clientInfo":"Number and Name","InterestCol":"0","txtShowLink":"N","baseOriginalFavoriteId":"a7874644a4924c94b522a788dc703aec","baseSelectionRows":579,"_desc_saveOptionRole":["","","","",""],"saveOptionRole":["ACCOUNTANT","ACCOUNTING","CONTROLLER-RO","[CREATOR_USERNAME]","CONTROLLER"]}}

        response = requests.post(url, headers=HEADERS, json=payload  )
        data = response.json()
        report_path_raw = data["return"]["ReportPath"]
        report_path = report_path_raw.replace(" ", "%20")

        # Step 2: Get Nonce
        nonceUrl = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Security/Nonce"
        payload = {}
        response = requests.post(nonceUrl, headers=HEADERS, json=payload  )
        nonce = response.json()

        # Step 3: Get Viewer
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/AccountsReceivable/Invoice%20Register&reportName=Invoice%20Summary"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("Error")
        
        exportFileName_csv = "Sales Invoice Pre-Tax_"+ date.today().strftime("%Y-%m-%d") + ".csv"
        exportFileName_xlsx = "Sales Invoice Pre-Tax_"+ date.today().strftime("%Y-%m-%d") + ".xlsx"

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer"
                "&OpType=Export&FileName=Office+Earnings&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        response = requests.get(url, headers = HEADERS,stream=True  )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(base_folder, exportFileName_csv)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            print("✅ "+ csvName+" Downloaded")

        targetFile = os.path.join(base_folder, TEMPLATE0)
        util.excel_full_copy(inputFile=csvName, inputSheet=None, targetFile=targetFile, targetSheet="Invoice Summary_csv", onlyValue=True)

        # Step 6: Download the xlsx report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer"
                "&OpType=Export&FileName=Office+Earnings&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML" )
        
        response = requests.get(url, headers = HEADERS,stream=True  )
        if response.status_code == 200 and response.headers.get("Content-Type") == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
            xlsxName = os.path.join(base_folder, exportFileName_xlsx)
            if os.path.exists(xlsxName):
                os.remove(xlsxName)
            with open(xlsxName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            print("✅ "+ xlsxName+" Downloaded")
        
        util.excel_full_copy(inputFile=xlsxName, inputSheet="Invoice Summary", targetFile=targetFile, targetSheet="Invoice Summary_excel", onlyValue=False)
    except Exception as e:
        print("⚠️ Failed to download the sales invoice pre-tax JTD:", e)


def download_billed(base_folder, pkey = "", option_name = "", period = ""):
    download_template(base_folder, TEMPLATE0)
    download_invoice_pretax(base_folder, pkey, option_name)
    csvName = download_GL(base_folder, startPeriod='200301', endPeriod=period, needDownload=True, baseRecordSelection={"pKey":None,"name":"Records Selected","type":"CA","whereClauseSearch":"N","isLegacy":"N","searchOptions":[{"name":"Account","value":"4","type":"account","seq":1,"tableName":"CA","opp":"startsWith","condition":"and","searchLevel":0,"valueDescription":"4"}]}, fileName="4XXXX ")
    targetFile = os.path.join(base_folder, TEMPLATE0)
    util.excel_full_copy(inputFile=csvName, inputSheet=None, targetFile=targetFile, targetSheet="General Ledger Export", onlyValue=True)



def final_step(base_folder):
    source = base_folder
    target = os.path.join(os.getcwd(), "revenue model report export")
    util.check_folder(target)
    util.clear_folder(target)
    util.move_and_replace_files(source_dir=source, target_dir=target)


def main():
    print(util.get_onedrive_path())

    #global project_list
    #project_list = set()

    #base_folder = Path("C:\\temp\\revenue_model_export")
    #init(base_folder)
    #get_search_list()
    #download_serch_options_list(pkey="5a14b2d45288e71f3a6919c0e9f9c513")
    #download_template(base_folder, TEMPLATE1)
    #download_template(base_folder, TEMPLATE2)
    #start_period = input("Please input the start period (202601): ")
    #end_period = input("Please input the end period (202607): ")
    #util.change_period(end_period)

    #generate_search_option(base_folder, start_period, end_period)

    #print(f"Report period set to {period}")
    #print("--------------------------------------\n\n")
    ##download_project_list(base_folder = Path("C:\\temp\\revenue_model_export") )
    #get_search_list()

    #download_contract(base_folder, period, pkey = "27d0f5c8e91c7c709c014d697a71b7fc", option_name = "Rev_2026-04-13_w4")
    #download_labour_details_JTD(base_folder, pkey = "27d0f5c8e91c7c709c014d697a71b7fc", option_name = "Rev_2026-04-13_w4")
    #download_cost_GL(base_folder, period)
    #download_project_labor_hrs(base_folder, pkey = "27d0f5c8e91c7c709c014d697a71b7fc", option_name = "Rev_2026-04-13_w4")
    #download_project_expense(base_folder, pkey = "27d0f5c8e91c7c709c014d697a71b7fc", option_name = "Rev_2026-04-13_w4")
    #download_JTD_Billing(base_folder, pkey = "27d0f5c8e91c7c709c014d697a71b7fc", option_name = "Rev_2026-04-13_w4")
    # download_billed(base_folder, pkey = "27d0f5c8e91c7c709c014d697a71b7fc", option_name = "Rev_2026-04-13_w4", period=period)
    # final_step(base_folder)

if __name__ == '__main__':
    main()