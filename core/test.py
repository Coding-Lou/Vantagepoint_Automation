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
import json

def download_AR(HEADERS):
    try:
        # Step 1: Build
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/AccountsReceivable/AR Aging","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseChart3D":"N","baseChartColumn":"InvoiceBalance","baseChartDivisor":"1","baseChartFontSize":8,"baseChartHeight":3,"baseChartLabelLines":"N","baseChartLabels":"none","baseChartLeft":1,"baseChartLegendPosition":"righttop","baseChartSeriesColumn2":"","baseChartSeriesColumn3":"","baseChartShowPosition":"1","baseChartTitle":"","baseChartTop":0.5,"baseChartType":"none","baseChartWidth":6,"baseChartXTitle":"Primary Client Name","baseChartYTitle":"Balance","baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.1,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"AR by Client Project Invoice","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"Y","baseShowTotalsOnHeader":"Y","baseStartColumnPosition":0.5,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"","sort":"ASC","color":"000000","subTotal":"Y","showHeading":"Y","pageHeading":"N","collapseExpand":"D","line":"DotLine","pageBreak":"N","groupID":"clientName","customGridColumnSort":"","groupWBSLevel":"1"}],"ReportColumns":[{"heading":"Project","width":1.8,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"WBS1","username":"","customGridColumnSort":""},{"heading":"Invoice","width":1,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"InvoiceNumber","username":"","customGridColumnSort":""},{"heading":"Date","width":0.7,"format":"yyyy/M/d","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"InvoiceDate","username":"","customGridColumnSort":""},{"heading":"Due Date","width":0.7,"format":"yyyy/M/d","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"DueDate","username":"","customGridColumnSort":""},{"heading":"Balance","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"InvoiceBalance","username":"","customGridColumnSort":""},{"heading":"Current","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Age1Amount","username":"","customGridColumnSort":"","altCol1":0,"altCol2":30},{"heading":"31-60","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Age2Amount","username":"","customGridColumnSort":"","altCol1":31,"altCol2":60},{"heading":"61-90","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Age3Amount","username":"","customGridColumnSort":"","altCol1":61,"altCol2":90},{"heading":"91 Over","width":0.7,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Age4Amount","username":"","customGridColumnSort":"","altCol1":91,"altCol2":99999}],"ReportSections":[],"baseRecordSelection":"{\"pKey\":null,\"type\":\"wbs1\",\"name\":\"Records Selected\",\"whereClauseSearch\":\"N\",\"searchOptions\":[{\"name\":\"ClientID\",\"value\":\"QCAW284\",\"type\":\"client\",\"seq\":1,\"tableName\":\"PR\",\"opp\":\"!=\",\"condition\":\"and\",\"searchLevel\":1,\"valueDescription\":\"QCA Systems Ltd.\"}]}","baseShowDetail":"Y","baseLeft1":0,"baseRight1":21,"baseLeft2":0,"baseRight2":8,"baseLeft3":0,"baseRight3":8,"baseSub":"1","rollType":"Project","timeframe":"JTD","radioTF":"radio1","tfCYJ":"J","interestCol":"0","clientBalance":"0","agedEndPeriod":"0","daysOld":"0","invoiceBalance":"0","agingDate":"2024-01-16T17:06:00.122","radioAD":"radioAD1","radioAU":"radioAU1","clientInfo":"None","includeInvDetail":"Y","chkIncInvOver":"N","chkPrintPreInvoices":"N","displayRtgInvoice":"N","chkIncInvPer":"N","excludeUReceipts":"N","ageUReceipts":"N","chkClientBalanceOver":"N","ageInterest":"Y","chkInvBalanceOver":"N","lastReceiptInfo":"N","showClientMemo":"N","finalAvgAge":"N","displayAR":"N","finalReceiptDate":"N","printRetainer":"N","txtShowLink":"N","commentSDate":"","commentEDate":"","printARComments":"N","_desc_saveOptionRole":["","","","","",""],"saveOptionRole":["[CREATOR_USERNAME]","ACCOUNTING","CONTROLLER","CONTROLLER-RO","ACCOUNTANT","PRINCIPAL"],"baseOriginalFavoriteId":"B2F3E409D153459F88D39F5D2CF46EC9","baseSelectionRows":11471}}

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
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/AccountsReceivable/AR%20Aging&reportName=AR%20by%20Client%20Project%20Invoice"

        print(url)

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("Error")
        
        exportFileName = "AR_AGing_"+ date.today().strftime("%Y-%m-%d") + ".csv"

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
                f"&FileName=AR+by+Client+Project+Invoice&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        print(url)
        
        response = requests.get(url, headers = HEADERS,stream=True  )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            csvName = os.path.join(r"C:\\temp\\", exportFileName)
            if os.path.exists(csvName):
                os.remove(csvName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            print("✅ "+ csvName+" Downloaded")
        

    except Exception as e:
        print("⚠️ Failed to download the AR Aging:", e)
