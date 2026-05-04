import csv
import shutil
import tools.util as util
import requests
import re
import os
from datetime import date, time
from pathlib import Path
import logging
import time

global HEADERS
HEADERS = util.set_headers()
global searchOptions
searchOptions = None
global targetFile

def setup_logger(file_name="program_log.txt"):
    # 统一路径：确保这里和你 final_step 里的路径逻辑一致
    log_dir = r"C:\temp" 
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    log_path = os.path.join(log_dir, file_name)

    logger = logging.getLogger("MyLogger")
    logger.setLevel(logging.DEBUG)

    # 清除旧的 handler 防止重复打印
    if logger.handlers:
        logger.handlers.clear()

    # 创建 Handler
    file_handler = logging.FileHandler(log_path, encoding='utf-8')
    
    # 关键点：让日志立即写入硬盘，不要等缓存
    # 在某些环境下，如果不显式设置，程序崩溃时日志就丢了
    
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    
    return logger

def update_templete(templete_abspath, copy_to_template = True):
    # 1. Initialize the Path object
    save_path = Path(templete_abspath)
    
    # 2. If directory exists, clear it completely
    if save_path.exists():
        try:
            shutil.rmtree(save_path)
            print(f"💥 Cleared entire directory: {save_path}")
        except PermissionError:
            print(f"❌ Permission Denied: Could not clear {save_path}. Is an Excel file open?")
            return  # Stop here if we can't clean the old files
        except Exception as e:
            print(f"❌ Error during cleanup: {e}")
            return
    # 3. Create the directory fresh
    try:
        save_path.mkdir(parents=True, exist_ok=True)
        print(f"📁 Directory ready: {save_path}")
    except Exception as e:
        print(f"❌ Error creating directory: {e}")

    # Convert back to string for compatibility with older functions if needed
    if (copy_to_template) :
        save_dir = str(save_path.absolute())
        util.download_from_gdrive(file_id= "1wlCCykoDKqgrILSUlDT69oFfUiCAPMa2", file_name="Project Status.xlsx", save_dir=save_dir)
        global targetFile
        targetFile = os.path.join(save_path, "Project Status.xlsx")

def print_period():
    url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/PeriodSetup/?meta=access"
    response = requests.get(url, headers=HEADERS)
    data = response.json()[:5]
    for p in data:
        print(f"{p['Period']} | From: {p['AccountPdStart'][:10]} To: {p['AccountPdEnd'][:10]}")

    return data

def download_invoices(copy_to_template = True):
    try:
        # Step 1: Build
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/AccountsReceivable/Invoice Register","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseChart3D":"N","baseChartColumn":"Other","baseChartDivisor":"1","baseChartFontSize":8,"baseChartHeight":3,"baseChartLabelLines":"N","baseChartLabels":"none","baseChartLeft":1,"baseChartLegendPosition":"righttop","baseChartSeriesColumn2":"","baseChartSeriesColumn3":"","baseChartShowPosition":"1","baseChartTitle":"","baseChartTop":0.5,"baseChartType":"none","baseChartWidth":6,"baseChartXTitle":"Primary Client Name","baseChartYTitle":"Other","baseCulture":"default","baseDefaultCurrencyFormat":"#########D##;-#########D##;0D00","baseDefaultDateFormat":"yyyy-MM-dd","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"#########D##;-#########D##;0D00","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.1,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Invoice Register","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"Y","baseStartColumnPosition":0.5,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"Primary Client Name","sort":"ASC","color":"000000","subTotal":"N","showHeading":"N","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"clientName","customGridColumnSort":"","groupWBSLevel":"1"},{"label":"Project Number","sort":"ASC","color":"000000","subTotal":"N","showHeading":"N","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"projectNumber","customGridColumnSort":"","groupWBSLevel":"1"}],"ReportColumns":[{"heading":"Project","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"WBS1","username":"","customGridColumnSort":""},{"heading":"Date","width":0.7,"format":"yyyy-MM-dd","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"TransDate","username":"","customGridColumnSort":""},{"heading":"Invoice","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"InvoiceNumber","username":"","customGridColumnSort":""},{"heading":"Total","width":0.7,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"TotalAmt","username":"","customGridColumnSort":""},{"heading":"Prof Fees","width":0.7,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col1","username":"","customGridColumnSort":""},{"heading":"H/W Sales","width":0.7,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col2","username":"","customGridColumnSort":""},{"heading":"S/W Sales","width":0.7,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col3","username":"","customGridColumnSort":""},{"heading":"O/S Services","width":0.7,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col4","username":"","customGridColumnSort":""},{"heading":"Reimbursable","width":0.7,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col5","username":"","customGridColumnSort":""},{"heading":"EHF","width":0.7,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col6","username":"","customGridColumnSort":""},{"heading":"Other","width":0.7,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Other","username":"","customGridColumnSort":""},{"heading":"Taxes","width":0.7,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col7","username":"","customGridColumnSort":""},{"heading":"Net Revenue","width":1,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Net Revenue","header1":"Net Revenue","header2":"","detailExpression":"[TotalAmt]-[Col7]","groupExpression":"[TotalAmt]-[Col7]","queryJoin":"","checkSecurity":"N","altHeader1":"","altHeader2":"","queryColumn":"","calculatedColumnType":"ALLFRAMES","username":"DPALACIO","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":{"pKey":"","name":"","type":"wbs1","whereClauseSearch":"N","isLegacy":"N","searchOptions":searchOptions},"baseShowDetail":"Y","baseLeft1":0,"baseRight1":21,"baseLeft2":0,"baseRight2":8,"baseLeft3":0,"baseRight3":8,"baseSub":"1","rollType":"Project","timeframe":"JTD","radioTF":"radio1","tfCYJ":"J","clientInfo":"None","InterestCol":"0","txtShowLink":"N","baseSelectionRows":1,"baseOriginalFavoriteId":"385322010f1c4997b596c9017169c40b","_desc_saveOptionRole":["","",""],"saveOptionRole":["[CREATOR_USERNAME]","ACCOUNTANT","ACCOUNTING"]}}
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
            raise RuntimeError("Error")
        
        exportFileName = "Project Invoices_"+ date.today().strftime("%Y-%m-%d") + ".csv"

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
        
        response = requests.get(url, headers = HEADERS,stream=True  )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(r"C:\\temp\\project status", exportFileName)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            print("✅ "+ csvName+" Downloaded")
        
        if copy_to_template:
            temp_file = csvName + ".temp"
            lines_to_skip = 4
            with open(csvName, 'r', encoding='utf-8', newline='') as f_in, \
                    open(temp_file, 'w', encoding='utf-8', newline='') as f_out:
                    reader = csv.reader(f_in)
                    writer = csv.writer(f_out)
                    for _ in range(lines_to_skip):
                        next(reader, None)
                    for row in reader:
                        filtered_row = row[0:13]
                        writer.writerow(filtered_row)

            os.replace(temp_file, csvName)

    except Exception as e:
        print("⚠️ Failed to download the invoices register:", e)

def download_earnings(copy_to_template = True):
    try:
        # Step 1: Build
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/Project/Project Earnings","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseChart3D":"N","baseChartColumn":"contractFeeDirect_Labor","baseChartDivisor":"1","baseChartFontSize":8,"baseChartHeight":3,"baseChartLabelLines":"N","baseChartLabels":"none","baseChartLeft":1,"baseChartLegendPosition":"righttop","baseChartSeriesColumn2":"","baseChartSeriesColumn3":"","baseChartShowPosition":"1","baseChartTitle":"","baseChartTop":0.5,"baseChartType":"none","baseChartWidth":6,"baseChartXTitle":"Project Number","baseChartYTitle":"Contract Direct Labour","baseCulture":"default","baseDefaultCurrencyFormat":"#########D##;-#########D##;0D00","baseDefaultDateFormat":"yyyy-MM-dd","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"#########D##;-#########D##;0D00","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.14,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Project Earnings","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":2.85,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"Project Number","sort":"ASC","color":"000080","subTotal":"Y","showHeading":"Y","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"projectNumber","customGridColumnSort":"","groupWBSLevel":"1"}],"ReportColumns":[{"heading":"Billing Client Name","width":1.54999,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"BillClientName","username":"","customGridColumnSort":""},{"heading":"Status","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"StatusDesc","username":"","customGridColumnSort":""},{"heading":"Biller","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"billerName","username":"","customGridColumnSort":""},{"heading":"Project Manager","width":1.25,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"prgName","username":"","customGridColumnSort":""},{"heading":"Principal","width":1.25,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"prinName","username":"","customGridColumnSort":""},{"heading":"Contract Number","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustContractNumber","username":"","customGridColumnSort":"N"},{"heading":"Contract Direct Labour","width":1,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"contractFeeDirect_Labor","username":"","customGridColumnSort":""},{"heading":"Contract Direct Expense","width":1,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"contractFeeDirect_Exp","username":"","customGridColumnSort":""},{"heading":"Contract Direct Consultant","width":1,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"contractconsultFee_","username":"","customGridColumnSort":""},{"heading":"Contract Reimb. Consultant","width":1,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"contractReimbAllow_Consult","username":"","customGridColumnSort":""},{"heading":"Contract Reimb. Expense","width":1,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"contractReimbAllow_Exp","username":"","customGridColumnSort":""},{"heading":"Budget Hours","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"hrsBud","username":"","customGridColumnSort":""},{"heading":"Bud Exp Amount","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"amtBud_c_Exp","username":"","customGridColumnSort":""},{"heading":"JTD Revenue","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"revJTD_c_","username":"","customGridColumnSort":""},{"heading":"JTD Billed","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"billJTD_c_","username":"","customGridColumnSort":""},{"heading":"JTD Unbilled","width":0.85002,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"unbJTD_c_","username":"","customGridColumnSort":""},{"heading":"Create Date","width":1,"format":"yyyy-MM-dd","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"CreateDate","username":"","customGridColumnSort":""},{"heading":"JTD Mark Up","width":1,"format":"###T###T###T###D##;(###T###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"JTD Mark Up","header1":"JTD Mark Up","header2":"","detailExpression":"[amtJTD_b_Exp]/[amtJTD_c_Exp]","groupExpression":"[amtJTD_b_Exp]/[amtJTD_c_Exp]","queryJoin":"","checkSecurity":"N","altHeader1":"","altHeader2":"","queryColumn":"","calculatedColumnType":"ALLFRAMES","username":"DPALACIO","customGridColumnSort":""},{"heading":"Effective Rate Bill","width":0.84999,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"effRate_b_","username":"","customGridColumnSort":""},{"heading":"Labor Estimate (No. of Hours)","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustLaborEstimate","username":"","customGridColumnSort":"N"},{"heading":"HW Estimate (Cost in CAD)","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustHWEstimate","username":"","customGridColumnSort":"N"},{"heading":"SW Estimate (Cost in CAD)","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustSWEstimate","username":"","customGridColumnSort":"N"},{"heading":"Outside Services Estimate (Cost in CAD)","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustOutsideServicesEstimate","username":"","customGridColumnSort":"N"},{"heading":"Other Expenses Estimate (Cost in CAD)","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustOtherExpensesEstimate","username":"","customGridColumnSort":"N"},{"heading":"Materials Only","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustMaterialsOnly","username":"","customGridColumnSort":"N"},{"heading":"Time & Material","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustTimeMaterial","username":"","customGridColumnSort":"N"},{"heading":"Defined Scope/deliverables","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustDefinedScopeValue","username":"","customGridColumnSort":"N"},{"heading":"Monthly Fixed Fee","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustLumpSumBilling","username":"","customGridColumnSort":"N"},{"heading":"No Billing","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustNoRevenue","username":"","customGridColumnSort":"N"}],"ReportSections":[],"baseRecordSelection":{"pKey":"a9fec334a0f14410aa4a37c5c253706e","name":"Lantic PLC Upgrade","type":"wbs1","whereClauseSearch":"N","isLegacy":"N","searchOptions":searchOptions},"baseCreateActivity":"N","baseShowDetail":"Y","baseLeft1":0,"baseRight1":21,"baseLeft2":0,"baseRight2":8,"baseLeft3":0,"baseRight3":8,"baseSub":"1","rollType":"Project","CurrentWBSActivityActiveWBS1Only":"N","CurrentWBSActivityActiveWBS2Only":"N","CurrentWBSActivityActiveWBS3Only":"N","CurrentWBSActivityActivityRange":"1","CurrentWBSActivityInclInvoiceActivity":"N","ReportFormat":"VOS","includeType":"All","hoursandAmounts":"1","labDrillDown":"1","expDrillDown":"1","atCost":"1","useSummaryTable":"N","CurrentWBSActivityCheckLabor":"Y","CurrentWBSActivityCheckExpense":"Y","CurrentWBSUnpostsedLabor":"Y","budgetSelection":"1","ETCDRadioChecked":"radioETCD1","showOverhead":"N","estimateOverhead":"N","currencyType":"project","targetDate":"7/24/2025 9:55:07 PM","currencyTypeARFee":"billing","currencyTypeBill":"billing","startPeriod":"03/2026","endPeriod":"03/2026","SummaryTableLastUpdate":"","baseSelectionRows":1,"_desc_saveOptionRole":["","","","",""],"saveOptionRole":["CONTROLLER-RO","ACCOUNTING","ACCOUNTANT","CONTROLLER","[CREATOR_USERNAME]"],"baseOriginalFavoriteId":"cf11a471a01e4b8492889a9c4e0772af"}}
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
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/Project/Project%20Earnings&reportName=Project%20Earnings"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("Error")
        
        exportFileName = "Project Earnings_"+ date.today().strftime("%Y-%m-%d") + ".csv"

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
                f"&FileName=Project+Earnings"
                "&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        response = requests.get(url, headers = HEADERS,stream=True  )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(r"C:\\temp\\project status", exportFileName)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            print("✅ "+ csvName+" Downloaded")

        if copy_to_template:
            # Trim
            temp_file = csvName + ".temp"
            lines_to_skip = 4
            with open(csvName, 'r', encoding='utf-8', newline='') as f_in, \
                    open(temp_file, 'w', encoding='utf-8', newline='') as f_out:
                    reader = csv.reader(f_in)
                    writer = csv.writer(f_out)
                    for _ in range(lines_to_skip):
                        next(reader, None)
                    for row in reader:
                        filtered_row = row[1:31]
                        writer.writerow(filtered_row)

            os.replace(temp_file, csvName)

    except Exception as e:
        print("⚠️ Failed to download the project earnings:", e)
    
    # try:
    #     util.csv_to_xlsx(csvName, output_file, "proj export", True, 1, 30)
    # except Exception as e:
    #     print("⚠️ Failed to copy the data to output file:", e)

def download_expenses(copy_to_template = True):
    try:
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/DataExport/ProjectExpenseDataSource","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseCulture":"default","baseDefaultCurrencyFormat":"#########D##;-#########D##;0D00","baseDefaultDateFormat":"yyyy-MM-dd","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"#########D##;-#########D##;0D00","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Project Expense Export","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":0,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[],"ReportColumns":[{"heading":"Project","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Project","username":"","customGridColumnSort":""},{"heading":"Task","width":0.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Phase","username":"","customGridColumnSort":""},{"heading":"Subtask","width":0.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Task","username":"","customGridColumnSort":""},{"heading":"Expense Account","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_Account","username":"","customGridColumnSort":""},{"heading":"TransType","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_TransType","username":"","customGridColumnSort":""},{"heading":"Reference Number","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_RefNo","username":"","customGridColumnSort":""},{"heading":"Date","width":0.75,"format":"yyyy-MM-dd","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_transDate","username":"","customGridColumnSort":""},{"heading":"Description 1","width":2,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_Desc1","username":"","customGridColumnSort":""},{"heading":"Description 2","width":2,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_Desc2","username":"","customGridColumnSort":""},{"heading":"Cost Amount","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_costAmount","username":"","customGridColumnSort":""},{"heading":"Bill Amount","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_BillAmount","username":"","customGridColumnSort":""},{"heading":"Bill Status","width":0.55,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_BillStatus","username":"","customGridColumnSort":""},{"heading":"Period","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_period","username":"","customGridColumnSort":""},{"heading":"Post Seq","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_PostSeq","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":{"pKey":"","name":"Records Selected","type":"wbs1","whereClauseSearch":"N","isLegacy":"N","searchOptions":searchOptions},"ReportingPeriod":"1","W2Year":"","W2Qtr":"","W2Quarter":"","baseSelectionRows":2,"_desc_saveOptionRole":["","",""],"saveOptionRole":["[CREATOR_USERNAME]","ACCOUNTANT","ACCOUNTING"],"baseOriginalFavoriteId":""}}
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
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/DataExport/ProjectExpenseDataSource&reportName=Project%20Expense%20Export"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("Error")
        
        exportFileName = "Project Expense_"+ date.today().strftime("%Y-%m-%d") + ".csv"

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer&OpType=Export&FileName=Project+Expense+Export&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        response = requests.get(url, headers = HEADERS,stream=True  )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(r"C:\\temp\\project status", exportFileName)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            print("✅ "+ csvName+" Downloaded")
        
        if copy_to_template:
            # Trim
            temp_file = csvName + ".temp"
            lines_to_skip = 1
            with open(csvName, 'r', encoding='utf-8', newline='') as f_in, \
                    open(temp_file, 'w', encoding='utf-8', newline='') as f_out:
                    reader = csv.reader(f_in)
                    writer = csv.writer(f_out)
                    for _ in range(lines_to_skip):
                        next(reader, None)
                    for row in reader:
                        if (row == [] or row[3] == ""):
                            continue
                        writer.writerow(row)
            os.replace(temp_file, csvName)

    except Exception as e:
        print("⚠️ Failed to download the invoices register:", e)

    # try:
    #     util.csv_to_xlsx(csvName, output_file, "exp export", False, 0, 13)
    # except Exception as e:
    #     print("⚠️ Failed to copy the data to output file:", e)

def download_labor_hours(copy_to_template = True):
    try:
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/DataExport/ProjectLaborDataSource","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseCulture":"default","baseDefaultCurrencyFormat":"#########D##;-#########D##;0D00","baseDefaultDateFormat":"yyyy-MM-dd","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"#########D##;-#########D##;0D00","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Project Labour Export","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":0,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[],"ReportColumns":[{"heading":"Project","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Project","username":"","customGridColumnSort":""},{"heading":"Task","width":0.625,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Phase","username":"","customGridColumnSort":""},{"heading":"Subtask","width":0.625,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Task","username":"","customGridColumnSort":""},{"heading":"Employee Number","width":0.875,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Employee","username":"","customGridColumnSort":""},{"heading":"Employee Full Name","width":1.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Employee_Name","username":"","customGridColumnSort":""},{"heading":"Date","width":0.75,"format":"yyyy-MM-dd","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_TransDate","username":"","customGridColumnSort":""},{"heading":"Regular Hours","width":0.625,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Regular_Hours","username":"","customGridColumnSort":""},{"heading":"Overtime Hours","width":0.75,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Overtime_Hours","username":"","customGridColumnSort":""},{"heading":"Special Overtime Hours","width":1,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Special_Overtime_Hours","username":"","customGridColumnSort":""},{"heading":"Labour Bill Rate","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Bill_Rate","username":"","customGridColumnSort":""},{"heading":"Labour Bill Amount","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Bill_Amount","username":"","customGridColumnSort":""},{"heading":"Bill Status","width":0.875,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_BillStatus","username":"","customGridColumnSort":""},{"heading":"Period","width":0.55,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Period","username":"","customGridColumnSort":""},{"heading":"Task Name","width":2,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Phase_Name","username":"","customGridColumnSort":""},{"heading":"Subtask Name","width":2,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Task_Name","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":{"pKey":"","name":"Records Selected","type":"wbs1","whereClauseSearch":"N","isLegacy":"N","searchOptions":searchOptions},"ReportingPeriod":"1","W2Year":"","W2Qtr":"","W2Quarter":"","baseSelectionRows":2,"_desc_saveOptionRole":["","",""],"saveOptionRole":["ACCOUNTANT","[CREATOR_USERNAME]","ACCOUNTING"],"baseOriginalFavoriteId":""}}
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
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/DataExport/ProjectLaborDataSource&reportName=Project%20Labour%20Export"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("Error")
        
        exportFileName = "Project Hours_"+ date.today().strftime("%Y-%m-%d") + ".csv"

        # Step 5: Download the csv report
        url = ( "https://qcadeltek03.qcasystems.com"
                "/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd"
                f"?ReportSession={report_session.group(1)}"
                "&Culture=1033&CultureOverrides=True"
                "&UICulture=2057&UICultureOverrides=True"
                "&ReportStack=1"
                f"&ControlID={control_id.group(1)}"
                "&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer"
                "&OpType=Export&FileName=Project+Labour+Export&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        response = requests.get(url, headers = HEADERS,stream=True  )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(r"C:\\temp\\project status", exportFileName)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            print("✅ "+ csvName+" Downloaded")
    except Exception as e:
        print("⚠️ Failed to download the invoices register:", e)
    
    # try:
    #     util.csv_to_xlsx(csvName, output_file, "hrs export", False, 0, 14)
    # except Exception as e:
    #     print("⚠️ Failed to copy the data to output file:", e)

def downloand_open_po(projects, copy_to_template = True):
    write_header = False
    exportFileName = "Project Purchase Orders_" + date.today().strftime("%Y-%m-%d") + ".csv"
    csvName = os.path.join(r"C:\\temp\\project status", exportFileName)

    try:
        with open(csvName, "w", newline="", encoding="utf-8-sig") as f:
            writer = None
            for project in projects:
                project = util.cleanup_projectID(project)
                url = (
                    f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/"
                    f"projectreview/{project}/PurchaseOrders?Closed=N&POStatus="
                )
                response = requests.get(url, headers=HEADERS)
                po_data = response.json()

                if len(po_data) == 0:
                    #print(f"⚠️ No open purchase orders for project {project}")
                    continue

                print(project + ": " + str(len(po_data)) + " purchase orders found")
                for row in po_data:
                    row["Project_Number"] = project

                if not write_header:
                    writer = csv.DictWriter(f, fieldnames=po_data[0].keys())
                    writer.writeheader()
                    write_header = True

                writer.writerows(po_data)
        
        if (copy_to_template):
            # Trim
            temp_file = csvName + ".temp"
            lines_to_skip = 1
            with open(csvName, 'r', encoding='utf-8', newline='') as f_in, \
                    open(temp_file, 'w', encoding='utf-8', newline='') as f_out:
                    reader = csv.reader(f_in)
                    writer = csv.writer(f_out)
                    for _ in range(lines_to_skip):
                        next(reader, None)
                    for row in reader:
                        writer.writerow(row[1:17])

            os.replace(temp_file, csvName)

        print("✅ " + exportFileName + " generated")

        # try:
        #     util.csv_to_xlsx(csvName, output_file, "PO Export", False, 1, 16)
        # except Exception as e:
        #     print("⚠️ Failed to copy the purchase orders data to output file:", e)

    except Exception as e:
        print(f"⚠️ Failed to download purchase orders for project {project}:", e)

def download_office_earnings(copy_to_template = True):
    try:
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/Project/Office Earnings","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseChart3D":"N","baseChartColumn":"rev","baseChartDivisor":"1","baseChartFontSize":8,"baseChartHeight":3,"baseChartLabelLines":"N","baseChartLabels":"none","baseChartLeft":1,"baseChartLegendPosition":"righttop","baseChartSeriesColumn2":"","baseChartSeriesColumn3":"","baseChartShowPosition":"1","baseChartTitle":"","baseChartTop":0.5,"baseChartType":"none","baseChartWidth":6,"baseChartXTitle":"Biller Number","baseChartYTitle":"","baseCulture":"default","baseDefaultCurrencyFormat":"#########D##;-#########D##;0D00","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"#########D##;-#########D##;0D00","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.1,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Office Earnings","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"Y","baseStartColumnPosition":2,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"Biller Number","sort":"ASC","color":"000000","subTotal":"N","showHeading":"Y","pageHeading":"N","collapseExpand":"E","line":"None","pageBreak":"N","groupID":"biller","customGridColumnSort":"","groupWBSLevel":"1"},{"label":"Status","sort":"ASC","color":"000000","subTotal":"N","showHeading":"Y","pageHeading":"N","collapseExpand":"E","line":"None","pageBreak":"N","groupID":"projectStatus","customGridColumnSort":"","groupWBSLevel":"1"},{"label":"Primary Client Name","sort":"ASC","color":"000000","subTotal":"N","showHeading":"Y","pageHeading":"N","collapseExpand":"E","line":"None","pageBreak":"N","groupID":"clientName","customGridColumnSort":"","groupWBSLevel":"1"},{"label":"Project Number","sort":"ASC","color":"000080","subTotal":"Y","showHeading":"Y","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"projectNumber","customGridColumnSort":"","groupWBSLevel":"1"}],"ReportColumns":[{"heading":"Revenue","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"rev","username":"","customGridColumnSort":""},{"heading":"Total Billed","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"bill","username":"","customGridColumnSort":""},{"heading":"Unbilled","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"unb","username":"","customGridColumnSort":""},{"heading":"Revenue Type","width":0.95,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"wbs3revType","username":"","customGridColumnSort":""},{"heading":"Hardware Rev Type","width":0.95,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"wbs3revType2","username":"","customGridColumnSort":""},{"heading":"Software Rev Type","width":0.95,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"wbs3revType3","username":"","customGridColumnSort":""},{"heading":"Outside Services Rev Type","width":0.95,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"wbs3revType4","username":"","customGridColumnSort":""},{"heading":"Others (Expenses) Rev Type","width":0.95,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"wbs3revType5","username":"","customGridColumnSort":""},{"heading":"Labour Revenue","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"rev1","username":"","customGridColumnSort":""},{"heading":"Hardware Revenue","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"rev2","username":"","customGridColumnSort":""},{"heading":"Software Revenue","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"rev3","username":"","customGridColumnSort":""},{"heading":"Outside Services Revenue","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"rev4","username":"","customGridColumnSort":""},{"heading":"Others (Expenses) Revenue","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"rev5","username":"","customGridColumnSort":""},{"heading":"Contract Number","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustContractNumber","username":"","customGridColumnSort":"N"},{"heading":"Contract Total Compensation","width":1,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"contractTotComp","username":"","customGridColumnSort":""},{"heading":"JTD Spent Labor, cost-plus","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustJTDSpentLaborcostplus","username":"","customGridColumnSort":"N"},{"heading":"JTD Spent HW, cost-plus","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustJTDSpentHWcostplus","username":"","customGridColumnSort":"N"},{"heading":"JTD Spent SW, cost-plus","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustJTDSpentSWcostplus","username":"","customGridColumnSort":"N"},{"heading":"JTD Spent OS, cost-plus","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustJTDSpentOScostplus","username":"","customGridColumnSort":"N"},{"heading":"JTD Spent OE, cost-plus","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustJTDSpentOEcostplus","username":"","customGridColumnSort":"N"},{"heading":"JTD Write-off Labor","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustJTDWriteoffLabor","username":"","customGridColumnSort":"N"},{"heading":"JTD Write-off HW","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustJTDWriteoffHW","username":"","customGridColumnSort":"N"},{"heading":"JTD Write-off SW","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustJTDWriteoffSW","username":"","customGridColumnSort":"N"},{"heading":"JTD Write-off OS","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustJTDWriteoffOS","username":"","customGridColumnSort":"N"},{"heading":"JTD Write-off OE","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustJTDWriteoffOE","username":"","customGridColumnSort":"N"},{"heading":"Proposal Hours on Hold","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustProposalHoursonHold","username":"","customGridColumnSort":"N"},{"heading":"Labour Unbilled","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"unb1","username":"","customGridColumnSort":""},{"heading":"Hardware Unbilled","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"unb2","username":"","customGridColumnSort":""},{"heading":"Software Unbilled","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"unb3","username":"","customGridColumnSort":""},{"heading":"Outside Services Unbilled","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"unb4","username":"","customGridColumnSort":""},{"heading":"Others (Expenses) Unbilled","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"unb5","username":"","customGridColumnSort":""},{"heading":"Other Unbilled","width":0.85,"format":"#########D##;-#########D##;0D00","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"unbOther","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":{"pKey":"","name":"","type":"wbs1","whereClauseSearch":"N","isLegacy":"N","searchOptions": searchOptions},"baseCreateActivity":"N","baseShowDetail":"Y","baseLeft1":0,"baseRight1":21,"baseLeft2":0,"baseRight2":8,"baseLeft3":0,"baseRight3":8,"baseSub":"1","rollType":"Project","CurrentWBSActivityActiveWBS1Only":"N","CurrentWBSActivityActiveWBS2Only":"N","CurrentWBSActivityActiveWBS3Only":"N","CurrentWBSActivityActivityRange":"1","CurrentWBSActivityInclInvoiceActivity":"N","drillDownSort":"1","budgetSelection":"1","ETCDRadioChecked":"radioETCD1","atCost":"2","labDrillDown":"1","expDrillDown":"1","showCur":"N","showYTD":"N","showJTD":"Y","custDate":"N","LabelCustDate":"Period Range","showOverhead":"N","estimateOverhead":"N","showUnposted":"N","chkIncludeCommitPO":"N","useSummaryTable":"N","excludeContractsNotinFees":"N","CurrentWBSActivityCheckLabor":"Y","CurrentWBSActivityCheckExpense":"Y","PrintDirects":"Y","ETCDate":"5/16/2024 1:38:16 PM","baseChartColumnDisplayTimeframe":"JTD","SummaryTableLastUpdate":"","CurrentWBSActivityUnpostedLabor":"N","CurrentWBSActivityCommittedExp":"N","baseOriginalFavoriteId":"","baseSelectionRows":1,"_desc_saveOptionRole":["","",""],"saveOptionRole":["ACCOUNTING","[CREATOR_USERNAME]","ACCOUNTANT"]}}

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
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/Project/Office%20Earnings&reportName=Office%20Earnings "

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("Error")
        
        exportFileName = "Project Office Earnings_"+ date.today().strftime("%Y-%m-%d") + ".csv"

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
            csvName = os.path.join(r"C:\\temp\\project status", exportFileName)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            print("✅ "+ csvName+" Downloaded")
    except Exception as e:
        print("⚠️ Failed to download the office earnings:", e)
    
    # try:
    #     util.csv_to_xlsx(csvName, output_file, "Office earnings", need_skip=True)
    # except Exception as e:
    #     print("⚠️ Failed to copy the data to output file:", e)


def copy_to_template():
    today_str = date.today().strftime("%Y-%m-%d")
    base_path = r"C:\\temp\\project status"
    
    tasks = [
        ("Project Invoices", "invoice export", "A4", False),
        ("Project Earnings", "proj export", "B5", False),
        ("Project Expense", "exp export", "A4", False),
        ("Project Hours", "hrs export", "A3", False),
        ("Project Purchase Orders", "open PO", "A3", True),
    ]

    import win32com.client
    import pythoncom
    log = setup_logger()
    log.info("--- Start logging ---")
    log.info("--- pythoncom.CoInitialize() ---")
    pythoncom.CoInitialize()
    log.info('--- excel = win32.DispatchEx("Excel.Application") ---')
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AlertBeforeOverwriting = False
    time.sleep(5)
    
    CALLEE_BUSY = -2147418111

    def safe_com_call(func, retries=15, delay=3):
        for i in range(retries):
            try:
                return func()
            except pythoncom.com_error as e:
                if getattr(e, 'hresult', None) != CALLEE_BUSY:
                    raise
                log.info(f"⚠️ COM busy, retrying {i+1}/{retries}...")
                time.sleep(delay)
        raise RuntimeError("Excel remained busy after retries")

    try:
        global targetFile
        
        target_wb = safe_com_call(
            lambda: excel.Workbooks.Open(os.path.abspath(targetFile))
        )
        
        for file_key, sheet, cell, is_last in tasks:
            input_path = os.path.join(base_path, f"{file_key}_{today_str}.csv")
            if os.path.exists(input_path):
               log.info(f"Copy from {file_key} {sheet} {cell}")
               util.excel_full_copy(
                    inputFile=input_path, 
                    inputSheet=None, 
                    targetFile=targetFile, 
                    targetSheet=sheet, 
                    onlyValue=True, 
                    targetCell=cell, 
                    refreshAll=is_last,
                    excel_instance=excel,
                    target_wb_instance=target_wb
                )
            else:
                print(f"⚠️ File Not Found: {input_path}")

        safe_com_call(target_wb.Save)
        print(f"✅ Data copied and saved to : {targetFile}")

    finally:
        if 'target_wb' in locals(): target_wb.Close()
        excel.Quit()
        pythoncom.CoUninitialize()

def move_to_folder():
    source = r"C:\\temp\\project status"
    target = os.path.join(os.getcwd(), "project status")
    util.move_and_replace_files(source_dir=source, target_dir=target)


def main():
    global searchOptions
    #print_period()
    period = input("Please input the period (202607): ")
    util.change_period(period)
    
    currentYear = date.today().year
    userInput = input(f"Do you want to filter by charge type is regular and created after {currentYear-3}.01.01? (Y/N): ")
    if userInput == "Y":
        startDate = f"{currentYear-3}-01-01T00:00:00"
        searchOptions = [{"name":"CreateDate","value":startDate,"type":"datetime","seq":1,"tableName":"PR","opp":">","condition":"and","searchLevel":1,"valueDescription":""},{"name":"ChargeType","value":"R","type":"dropdown","seq":2,"tableName":"PR","condition":"and","searchLevel":1,"valueDescription":"Regular"}]
    else:
        userInput = input("Project Name(s) (use commas to separate multiple entries): " )
        projects = [p.strip() for p in userInput.split(",") if p.strip()]
        if not projects:
            print("Error: Please enter at least one project name.")
            return
        searchOptions = util.assamble_projects(projects)
    
    util.check_folder("project status")
    util.clear_folder("project status")
    update_templete(r"C:\\temp\\project status")
    download_invoices()
    download_earnings()
    download_expenses()
    download_labor_hours()
    download_office_earnings()
    if (projects):
        downloand_open_po(projects)

    copy_to_template()
    '''
    download_earnings()
    download_expenses()
    download_labor_hours()
    downloand_open_po(projects)
    download_office_earnings()

    print(f'The output file is: output_{date.today().strftime("%Y-%m-%d")}.xlsx')
    '''

if __name__ == '__main__':
    main()