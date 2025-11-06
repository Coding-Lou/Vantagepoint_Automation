import util
import requests
import re
import os
from datetime import date

HEADERS = util.set_headers()
searchOptions = None

def main():
    global searchOptions
    print("Project Name(s) (use commas to separate multiple entries): " )
    userInput = input()
    projects = [p.strip() for p in userInput.split(",") if p.strip()]

    if not projects:
        print("Error: Please enter at least one project name.")
        return
    
    searchOptions = util.assamble_projects(projects)
    #period = input("Please input the period (202607): ")
    #set_period(period)
    util.check_folder("project status")
    print()
    download_invoices()
    download_earnings()
    download_expenses()
    download_labor_hours()

def set_period(period):
    url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/PeriodSetup/ActivePeriod/" + period
    response = requests.put(url, headers = HEADERS) 

def download_invoices():
    try:
        # Step 1: Build
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/AccountsReceivable/Invoice Register","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseChart3D":"N","baseChartColumn":"Other","baseChartDivisor":"1","baseChartFontSize":8,"baseChartHeight":3,"baseChartLabelLines":"N","baseChartLabels":"none","baseChartLeft":1,"baseChartLegendPosition":"righttop","baseChartSeriesColumn2":"","baseChartSeriesColumn3":"","baseChartShowPosition":"1","baseChartTitle":"","baseChartTop":0.5,"baseChartType":"none","baseChartWidth":6,"baseChartXTitle":"Primary Client Name","baseChartYTitle":"Other","baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"yyyy-MM-dd","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.1,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Invoice Register","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"Y","baseStartColumnPosition":0.5,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"Primary Client Name","sort":"ASC","color":"000000","subTotal":"N","showHeading":"N","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"clientName","customGridColumnSort":"","groupWBSLevel":"1"},{"label":"Project Number","sort":"ASC","color":"000000","subTotal":"N","showHeading":"N","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"projectNumber","customGridColumnSort":"","groupWBSLevel":"1"}],"ReportColumns":[{"heading":"Project","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"WBS1","username":"","customGridColumnSort":""},{"heading":"Date","width":0.7,"format":"yyyy-MM-dd","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"TransDate","username":"","customGridColumnSort":""},{"heading":"Invoice","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"InvoiceNumber","username":"","customGridColumnSort":""},{"heading":"Total","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"TotalAmt","username":"","customGridColumnSort":""},{"heading":"Prof Fees","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col1","username":"","customGridColumnSort":""},{"heading":"H/W Sales","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col2","username":"","customGridColumnSort":""},{"heading":"S/W Sales","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col3","username":"","customGridColumnSort":""},{"heading":"O/S Services","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col4","username":"","customGridColumnSort":""},{"heading":"Reimbursable","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col5","username":"","customGridColumnSort":""},{"heading":"EHF","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col6","username":"","customGridColumnSort":""},{"heading":"Other","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Other","username":"","customGridColumnSort":""},{"heading":"Taxes","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Col7","username":"","customGridColumnSort":""},{"heading":"Net Revenue","width":1,"format":"###T###T###T###D##;(###T###T###T###D##);#","align":"right","sectionName":"","sectionRow":0,"sectionColumn":0,"columnID":"Net Revenue","header1":"Net Revenue","header2":"","detailExpression":"[TotalAmt]-[Col7]","groupExpression":"[TotalAmt]-[Col7]","queryJoin":"","checkSecurity":"N","altHeader1":"","altHeader2":"","queryColumn":"","calculatedColumnType":"ALLFRAMES","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":"","baseShowDetail":"Y","baseLeft1":0,"baseRight1":21,"baseLeft2":0,"baseRight2":8,"baseLeft3":0,"baseRight3":8,"baseSub":"1","rollType":"Project","timeframe":"JTD","radioTF":"radio1","tfCYJ":"J","clientInfo":"None","InterestCol":"0","txtShowLink":"N","baseRecordSelection":{"pKey":"","name":"Records Selected","type":"wbs1","whereClauseSearch":"N","isLegacy":"N","searchOptions":searchOptions},"_desc_saveOptionRole":["","",""],"saveOptionRole":["[CREATOR_USERNAME]","ACCOUNTANT","ACCOUNTING"],"baseOriginalFavoriteId":""}}
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
            csvName = os.path.join("project status", exportFileName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            print("✅ "+ csvName+" Downloaded")

    except Exception as e:
        print("⚠️ Failed to download the invoices register:", e)

def download_earnings():
    try:
        # Step 1: Build
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/Project/Project Earnings","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseChart3D":"N","baseChartColumn":"contractFeeDirect_Labor","baseChartDivisor":"1","baseChartFontSize":8,"baseChartHeight":3,"baseChartLabelLines":"N","baseChartLabels":"none","baseChartLeft":1,"baseChartLegendPosition":"righttop","baseChartSeriesColumn2":"","baseChartSeriesColumn3":"","baseChartShowPosition":"1","baseChartTitle":"","baseChartTop":0.5,"baseChartType":"none","baseChartWidth":6,"baseChartXTitle":"Project Number","baseChartYTitle":"Contract Direct Labour","baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"yyyy-MM-dd","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.14,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Project Earnings","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":2.85,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"Project Number","sort":"ASC","color":"000080","subTotal":"Y","showHeading":"Y","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"projectNumber","customGridColumnSort":"","groupWBSLevel":"1"}],"ReportColumns":[{"heading":"Billing Client Name","width":1.54999,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"BillClientName","username":"","customGridColumnSort":""},{"heading":"Status","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"StatusDesc","username":"","customGridColumnSort":""},{"heading":"Biller","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"billerName","username":"","customGridColumnSort":""},{"heading":"Project Manager","width":1.25,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"prgName","username":"","customGridColumnSort":""},{"heading":"Principal","width":1.25,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"prinName","username":"","customGridColumnSort":""},{"heading":"Contract Number","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustContractNumber","username":"","customGridColumnSort":"N"},{"heading":"Contract Direct Labour","width":1,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"contractFeeDirect_Labor","username":"","customGridColumnSort":""},{"heading":"Contract Direct Expense","width":1,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"contractFeeDirect_Exp","username":"","customGridColumnSort":""},{"heading":"Contract Direct Consultant","width":1,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"contractconsultFee_","username":"","customGridColumnSort":""},{"heading":"Contract Reimb. Consultant","width":1,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"contractReimbAllow_Consult","username":"","customGridColumnSort":""},{"heading":"Contract Reimb. Expense","width":1,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"contractReimbAllow_Exp","username":"","customGridColumnSort":""},{"heading":"Budget Hours","width":0.85,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"hrsBud","username":"","customGridColumnSort":""},{"heading":"Bud Exp Amount","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"amtBud_c_Exp","username":"","customGridColumnSort":""},{"heading":"JTD Revenue","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"revJTD_c_","username":"","customGridColumnSort":""},{"heading":"JTD Billed","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"billJTD_c_","username":"","customGridColumnSort":""},{"heading":"JTD Unbilled","width":0.85002,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"unbJTD_c_","username":"","customGridColumnSort":""},{"heading":"Create Date","width":1,"format":"yyyy-MM-dd","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"CreateDate","username":"","customGridColumnSort":""},{"heading":"JTD Mark Up","width":1,"format":"###T###T###T###D##;(###T###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"JTD Mark Up","header1":"JTD Mark Up","header2":"","detailExpression":"[amtJTD_b_Exp]/[amtJTD_c_Exp]","groupExpression":"[amtJTD_b_Exp]/[amtJTD_c_Exp]","queryJoin":"","checkSecurity":"N","altHeader1":"","altHeader2":"","queryColumn":"","calculatedColumnType":"ALLFRAMES","username":"DPALACIO","customGridColumnSort":""},{"heading":"Effective Rate Bill","width":0.84999,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"effRate_b_","username":"","customGridColumnSort":""},{"heading":"Labor Estimate (No. of Hours)","width":0.85,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustLaborEstimate","username":"","customGridColumnSort":"N"},{"heading":"HW Estimate (Cost in CAD)","width":0.85,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustHWEstimate","username":"","customGridColumnSort":"N"},{"heading":"SW Estimate (Cost in CAD)","width":0.85,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustSWEstimate","username":"","customGridColumnSort":"N"},{"heading":"Outside Services Estimate (Cost in CAD)","width":0.85,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustOutsideServicesEstimate","username":"","customGridColumnSort":"N"},{"heading":"Other Expenses Estimate (Cost in CAD)","width":0.85,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustOtherExpensesEstimate","username":"","customGridColumnSort":"N"},{"heading":"Materials Only","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustMaterialsOnly","username":"","customGridColumnSort":"N"},{"heading":"Time & Material","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustTimeMaterial","username":"","customGridColumnSort":"N"},{"heading":"Defined Scope/deliverables","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustDefinedScopeValue","username":"","customGridColumnSort":"N"},{"heading":"Monthly Fixed Fee","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustLumpSumBilling","username":"","customGridColumnSort":"N"},{"heading":"No Billing","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UDCol_CustNoRevenue","username":"","customGridColumnSort":"N"}],"ReportSections":[],"baseRecordSelection":{"pKey":"","name":"Records Selected","type":"wbs1","whereClauseSearch":"N","isLegacy":"N","searchOptions":searchOptions},"baseCreateActivity":"N","baseShowDetail":"Y","baseLeft1":0,"baseRight1":21,"baseLeft2":0,"baseRight2":8,"baseLeft3":0,"baseRight3":8,"baseSub":"1","rollType":"Project","CurrentWBSActivityActiveWBS1Only":"N","CurrentWBSActivityActiveWBS2Only":"N","CurrentWBSActivityActiveWBS3Only":"N","CurrentWBSActivityActivityRange":"1","CurrentWBSActivityInclInvoiceActivity":"N","ReportFormat":"VOS","includeType":"All","hoursandAmounts":"1","labDrillDown":"1","expDrillDown":"1","atCost":"1","useSummaryTable":"N","CurrentWBSActivityCheckLabor":"Y","CurrentWBSActivityCheckExpense":"Y","CurrentWBSUnpostsedLabor":"Y","budgetSelection":"1","ETCDRadioChecked":"radioETCD1","showOverhead":"N","estimateOverhead":"N","currencyType":"project","targetDate":"7/24/2025 9:55:07 PM","currencyTypeARFee":"billing","currencyTypeBill":"billing","startPeriod":"03/2026","endPeriod":"03/2026","SummaryTableLastUpdate":"","baseSelectionRows":2,"baseOriginalFavoriteId":"","_desc_saveOptionRole":["","",""],"saveOptionRole":["[CREATOR_USERNAME]","ACCOUNTANT","ACCOUNTING"]}}
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
            csvName = os.path.join("project status", exportFileName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            print("✅ "+ csvName+" Downloaded")

    except Exception as e:
        print("⚠️ Failed to download the project earnings:", e)

def download_expenses():
    try:
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/DataExport/ProjectExpenseDataSource","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"yyyy-MM-dd","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Project Expense Export","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":0,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[],"ReportColumns":[{"heading":"Project","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Project","username":"","customGridColumnSort":""},{"heading":"Task","width":0.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Phase","username":"","customGridColumnSort":""},{"heading":"Subtask","width":0.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Task","username":"","customGridColumnSort":""},{"heading":"Expense Account","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_Account","username":"","customGridColumnSort":""},{"heading":"TransType","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_TransType","username":"","customGridColumnSort":""},{"heading":"Reference Number","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_RefNo","username":"","customGridColumnSort":""},{"heading":"Date","width":0.75,"format":"yyyy-MM-dd","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_transDate","username":"","customGridColumnSort":""},{"heading":"Description 1","width":2,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_Desc1","username":"","customGridColumnSort":""},{"heading":"Description 2","width":2,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_Desc2","username":"","customGridColumnSort":""},{"heading":"Cost Amount","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_costAmount","username":"","customGridColumnSort":""},{"heading":"Bill Amount","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_BillAmount","username":"","customGridColumnSort":""},{"heading":"Bill Status","width":0.55,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_BillStatus","username":"","customGridColumnSort":""},{"heading":"Period","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_period","username":"","customGridColumnSort":""},{"heading":"Post Seq","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Expense_PostSeq","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":{"pKey":"","name":"Records Selected","type":"wbs1","whereClauseSearch":"N","isLegacy":"N","searchOptions":searchOptions},"ReportingPeriod":"1","W2Year":"","W2Qtr":"","W2Quarter":"","baseSelectionRows":2,"_desc_saveOptionRole":["","",""],"saveOptionRole":["[CREATOR_USERNAME]","ACCOUNTANT","ACCOUNTING"],"baseOriginalFavoriteId":""}}
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
            csvName = os.path.join("project status", exportFileName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            print("✅ "+ csvName+" Downloaded")

    except Exception as e:
        print("⚠️ Failed to download the invoices register:", e)

def download_labor_hours():
    try:
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/DataExport/ProjectLaborDataSource","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"yyyy-MM-dd","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Project Labour Export","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":0,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[],"ReportColumns":[{"heading":"Project","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Project","username":"","customGridColumnSort":""},{"heading":"Task","width":0.625,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Phase","username":"","customGridColumnSort":""},{"heading":"Subtask","width":0.625,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Task","username":"","customGridColumnSort":""},{"heading":"Employee Number","width":0.875,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Employee","username":"","customGridColumnSort":""},{"heading":"Employee Full Name","width":1.5,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Employee_Name","username":"","customGridColumnSort":""},{"heading":"Date","width":0.75,"format":"yyyy-MM-dd","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_TransDate","username":"","customGridColumnSort":""},{"heading":"Regular Hours","width":0.625,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Regular_Hours","username":"","customGridColumnSort":""},{"heading":"Overtime Hours","width":0.75,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Overtime_Hours","username":"","customGridColumnSort":""},{"heading":"Special Overtime Hours","width":1,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Special_Overtime_Hours","username":"","customGridColumnSort":""},{"heading":"Labour Bill Rate","width":0.85,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Bill_Rate","username":"","customGridColumnSort":""},{"heading":"Labour Bill Amount","width":0.85,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Bill_Amount","username":"","customGridColumnSort":""},{"heading":"Bill Status","width":0.875,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_BillStatus","username":"","customGridColumnSort":""},{"heading":"Period","width":0.55,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Labor_Period","username":"","customGridColumnSort":""},{"heading":"Task Name","width":2,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Phase_Name","username":"","customGridColumnSort":""},{"heading":"Subtask Name","width":2,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Task_Name","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":{"pKey":"","name":"Records Selected","type":"wbs1","whereClauseSearch":"N","isLegacy":"N","searchOptions":searchOptions},"ReportingPeriod":"1","W2Year":"","W2Qtr":"","W2Quarter":"","baseSelectionRows":2,"_desc_saveOptionRole":["","",""],"saveOptionRole":["ACCOUNTANT","[CREATOR_USERNAME]","ACCOUNTING"],"baseOriginalFavoriteId":""}}
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
            csvName = os.path.join("project status", exportFileName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            print("✅ "+ csvName+" Downloaded")

    except Exception as e:
        print("⚠️ Failed to download the invoices register:", e)