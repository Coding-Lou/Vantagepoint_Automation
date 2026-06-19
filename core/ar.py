from pathlib import Path

import tools.util as util
import requests
import re
import os
import csv
from json import JSONDecodeError
from openpyxl import Workbook
from datetime import datetime
import zipfile

global HEADERS
HEADERS = util.set_headers()
# Email
global MAIL_FROM
MAIL_FROM = util.get_config(["AR", "FROM"])
global CC
CC = util.get_config(["AR", "CC"])
global SUBJECT
SUBJECT = util.get_config(["AR", "SUBJECT"])
global BODY
BODY = util.get_config(["AR", "BODY"])
global OPTIONALMSG
OPTIONALMSG = util.get_config(['AR', 'OPTIONALMSG'])
# Output location
global ONEDRIVEDIR
ONEDRIVEDIR = util.get_config(['ONEDRIVEDIR'])
global WORKDIR
WORKDIR = util.get_config(['WORKDIR'])
# Log Conifg
global RECORDS
RECORDS = 1
global STATEMENTDATE
STATEMENTDATE = None

def format_amount(val):
    return f"{val:,.2f}" if val and val > 0 else ""

# Don't delete, there is some glitch in the system.
def _safe_json(response, context):
    """Parse response JSON with actionable error context."""
    try:
        return response.json()
    except (JSONDecodeError, ValueError):
        body_preview = (response.text or "").strip().replace("\n", " ")[:200]
        raise RuntimeError(
            f"{context} returned non-JSON response "
            f"(status={response.status_code}, url={response.url}, body='{body_preview}')"
        )

def ar_init():
    ar_init_with_date(input("Statement Date (format 2025-05-01): "))

def ar_init_with_date(date: str):
    global STATEMENTDATE
    STATEMENTDATE = date
    util.check_folder(os.path.join(ONEDRIVEDIR, WORKDIR, "ar_export"))
    util.clear_folder(os.path.join(ONEDRIVEDIR, WORKDIR, "ar_export"))
    url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/UserSettings"
    payload = {"FW_SEUserOptions":[{"OptionName":"arReviewPaidStatus","OptionValue":"Unpaid","_transType":"U"}]}
    requests.post(url, headers=HEADERS, json=payload)

def init_output():
    global wb
    output_dir = Path(os.path.join(ONEDRIVEDIR, WORKDIR))
    if output_dir.is_dir():
        for file_path in output_dir.glob("Job*.xlsx"):
            file_path.unlink(missing_ok=True)
            print(f"Deleted: {file_path}")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["ClientID","From","To","CC","Subject","AttachmentName","AttachmentContent","Body", "ClientName"])

def ar_download_csv():
    try:
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/AccountsReceivable/AR Statement","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.1,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"Y","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"AR Statement","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":1,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[],"ReportColumns":[{"heading":"Number","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Invoice","username":"","customGridColumnSort":""},{"heading":"Date","width":0.8,"format":"M/d/yyyy","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"InvoiceDate","username":"","customGridColumnSort":""},{"heading":"Due Date","width":0.8,"format":"M/d/yyyy","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"DueDate","username":"","customGridColumnSort":""},{"heading":"Invoiced","width":0.8,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"OriginalAmt","username":"","customGridColumnSort":""},{"heading":"Balance Due","width":0.8,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Balance","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":"","baseShowDetail":"Y","statementType":"C","radioSD":"radioSD3","radioAU":"InvoiceDate","gracePeriod":"30","firmAlign":"Center","showClientMemo":"N","printLongName":"N","printFirmName":"Y","printByLine":"Y","printAddress":"Y","printProjName":"Y","printProjDesc":"Y","printProjNumber":"N","printFooter":"Y","showInvoiceLeadingZeros":"Y","excludeUReceipts":"Y","byLine":"","address1":"#101 6951 72 Street","address2":"Delta, BC","address3":"V4G 0A2","address4":"","HeaderMsg":"","footerMsg":"QCA Systems Ltd.    \nMain Office: #101 6951 72 Street, Delta, BC, V4G 0A2          Phone: (604)-940-0868          Fax: (604)-940-0869\nNorth Shore Office: #201 197 Forester Street, North Vancouver, BC, V7H 0A6\nAll Invoices are due upon receipt","PrintContactFirstName":"Y","PrintContactLastName":"Y","PrintContactMiddleName":"N","PrintContactPreferredName":"N","PrintContactPrefix":"N","PrintContactSuffix":"N","PrintContactTitle":"N","invoiceAddressee":"1","showFooter":"Y","statementSummary":"Y","agingSummary":"Y","clientSelName":"","MarginAndImages":"[{\"ImageID\":\"Our Firm Block\",\"Type\":\"FirmAddress\",\"TopPosition\":0.07,\"LeftPosition\":2.23,\"ColBand\":\"Header\",\"ImageWidth\":0,\"ImageHeight\":0,\"Item\":1,\"Selected\":\"Y\"},{\"ImageID\":\"Client Address\",\"Type\":\"ClientAddress\",\"TopPosition\":1.55,\"LeftPosition\":0,\"ColBand\":\"Header\",\"ImageWidth\":0,\"ImageHeight\":0,\"Item\":2,\"Selected\":\"Y\"},{\"ImageID\":\"Date Block\",\"Type\":\"DateBlock\",\"TopPosition\":1.17,\"LeftPosition\":5.63,\"ColBand\":\"Header\",\"ImageWidth\":0,\"ImageHeight\":0,\"Item\":3,\"Selected\":\"Y\"},{\"ImageID\":\"Statement Label\",\"Type\":\"StatementLabel\",\"TopPosition\":0.03,\"LeftPosition\":0.04,\"ColBand\":\"Header\",\"ImageWidth\":0,\"ImageHeight\":0,\"Item\":4,\"Selected\":\"Y\"}]","ageDays1":"30","ageDays2":"60","ageDays3":"90","ageDays4":"120","ageDays5":"150","statementDate":STATEMENTDATE+"T00:00:00.000","_desc_saveOptionRole":["","","","",""],"saveOptionRole":["ACCOUNTING","ACCOUNTANT","CONTROLLER","CONTROLLER-RO","[CREATOR_USERNAME]"],"baseOriginalFavoriteId":"DA82BA5C8CD94E2CAD6A96814CC49C5C"}}

        response = requests.post(url, headers=HEADERS,json=payload  )
        data = _safe_json(response, "AR statement build API")
        report_path_raw = data["return"]["ReportPath"]
        report_path = report_path_raw.replace(" ", "%20")

        # Step 2: Get Nonce
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Security/Nonce"
        payload = {}

        response = requests.post(url, headers=HEADERS, json=payload  )
        nonce = _safe_json(response, "AR nonce API")

        # Step 3: Get Viewer
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx??&nonce="+nonce+"&reportPath="+report_path+"&allowSchedule=N&reportName=AR%20Statement"

        response = requests.get(url, headers=HEADERS)
        html = response.text

        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id     = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        session_token = re.search(r'_token="([^"]+)"', html)

        # Step 4: SessionKeepALive
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/Reporting/Reserved.ReportViewerWebControl.axd?OpType=SessionKeepAlive&ControlID="+control_id.group(1)+"&RSProxy=https%3a%2f%2fqcadeltek03.qcasystems.com%2fReportServer"
        response = requests.post(url, headers=HEADERS)

        # Step 5: Download PDF
        fileName = "result.csv"

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
                "&ContentDisposition=OnlyHtmlInline"
                f"&Format=CSV" )

        response = requests.get(url_pdf, headers = HEADERS,stream=True  )

        if response.status_code == 200:
            csvName = os.path.join(ONEDRIVEDIR, WORKDIR, fileName)
            with open(csvName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            print("✅ AR statement list downloaded")
        else:
            print(f"❌ Error, status code: {response.status_code}")
            print("Msg:", response.content[:500])

        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/app/base/MakeVisionServiceRequest?method=DeleteAndStopReport"
        payload = {"sessionID": session_token.group(1), "reportPath": report_path_raw}
        response = requests.post(url, headers=HEADERS, json=payload )
        
    except Exception as e:
        print("❌ Error in download the full AR statement list ")


def ar_process():
    clientNames = set()
    with open(os.path.join(ONEDRIVEDIR, WORKDIR, "result.csv"), mode='r', newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        next(reader)
        for row in reader:
            if len(row) >= 3:
                value = row[2]
                if value:
                    clientNames.add(value)
    total = len(clientNames)
    zipClientName = []
    for i, clientName in enumerate(clientNames, start=1):
        due_invoice = ""
        if (not "QCA Systems Ltd." in clientName):
        #if ("GCT Canada Limited Partnership" in clientName):
            try:
                print(f"\n▶ [{i}/{total}] {clientName}")
                clientID = util.get_clientID(clientName)
                data = ar_review(clientID)
                email = util.get_vendor_email(clientID)
                if email != '':
                    print(f"  ✅ Email: {email}")
                else:
                    print(f"  ❌ Email not found")
                fileName = ar_download_statement_pdf(clientID, clientName)
                if fileName != '':
                    print(f"  ✅ Statement downloaded")
                else:
                    print(f"  ❌ Statement download failed")
                pmList = ""
                if ( (data['Age2'] != "" and int(data['Age2']) > 0) or
                    (data['Age3'] != "" and int(data['Age3']) > 0) or
                    (data['Age4'] != "" and int(data['Age4']) > 0) or
                    (data['Age5'] != "" and int(data['Age5']) > 0) ):
                    pmList = ar_get_pm_list(clientID)
                
                print(f"  Balance  0-30: ${data['Age1']:,.2f}  |  31-45: ${data['Age2']:,.2f}  |  46-60: ${data['Age3']:,.2f}  |  61-90: ${data['Age4']:,.2f}  |  90+: ${data['Age5']:,.2f}")

                tableContent, due_invoice = ar_generate_invoices_table(clientID, due_invoice)
                tableContent = '<table border="1" width="500" style="border-collapse: collapse"><thead><tr style="text-align: center;"><th>Invoice</th><th>0-30</th><th>31-45</th><th>46-60</th><th>61-90</th><th>90+</th></tr></thead><tbody>' + tableContent + "</tbody></table>"

                check = ar_details(clientID)
                if check:
                    ar_create_record(clientID, clientName, email, fileName, pmList, tableContent, due_invoice)
                    if ar_need_zip(clientID, clientName):
                        zipClientName.append(clientName)
                else:
                    print("Only has the credit invoices, exclude from the list.")
                
            except Exception as e:
                print(f"  ❌ Error: {e}")
                continue
    return zipClientName

def ar_review(clientID):
    try:
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/ARReview/"+ clientID +"?sumColumns=Total%2CAge1%2CAge2%2CAge3%2CAge4%2CAge5%2CTax%2CInterest%2CRetainage%2CRetainers"
        response = requests.get(url, headers = HEADERS)
        data = _safe_json(response, f"AR review API for client {clientID}")
        return data[0]
    except Exception as e:
        print("❌ Error in get full invoice of vendor " + clientID)

def ar_details(clientID):

    url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/ARReview/"+ clientID
    response = requests.get(url, headers = HEADERS)
    records = _safe_json(response, f"AR details API for client {clientID}")
    has_positive_invoice = False
    for r in records:
        if r['Total'] <= 0:
            continue
        has_positive_invoice = True
        ar_download_proj_invoice_pdf(r['WBS1'], clientID)
        print(f"  ✅ {r['WBS1']} invoice downloaded")
    
    return has_positive_invoice

def ar_download_statement_pdf(clientID, clientName):
    try:
        # Step 2: Build
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/AccountsReceivable/AR Statement","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.1,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"Y","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"AR Statement","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":1,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[],"ReportColumns":[{"heading":"Number","width":0.85,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Invoice","username":"","customGridColumnSort":""},{"heading":"Date","width":0.8,"format":"M/d/yyyy","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"InvoiceDate","username":"","customGridColumnSort":""},{"heading":"Due Date","width":0.8,"format":"M/d/yyyy","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"DueDate","username":"","customGridColumnSort":""},{"heading":"Invoiced","width":0.8,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"OriginalAmt","username":"","customGridColumnSort":""},{"heading":"Balance Due","width":0.8,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Balance","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":"","baseShowDetail":"Y","statementType":"C","radioSD":"radioSD3","radioAU":"InvoiceDate","gracePeriod":"30","firmAlign":"Center","showClientMemo":"N","printLongName":"N","printFirmName":"Y","printByLine":"Y","printAddress":"Y","printProjName":"Y","printProjDesc":"Y","printProjNumber":"N","printFooter":"Y","showInvoiceLeadingZeros":"Y","excludeUReceipts":"Y","byLine":"","address1":"#101 6951 72 Street","address2":"Delta, BC","address3":"V4G 0A2","address4":"","HeaderMsg":"","footerMsg":"QCA Systems Ltd.    \nMain Office: #101 6951 72 Street, Delta, BC, V4G 0A2          Phone: (604)-940-0868          Fax: (604)-940-0869\nNorth Shore Office: #201 197 Forester Street, North Vancouver, BC, V7H 0A6\nAll Invoices are due upon receipt","PrintContactFirstName":"Y","PrintContactLastName":"Y","PrintContactMiddleName":"N","PrintContactPreferredName":"N","PrintContactPrefix":"N","PrintContactSuffix":"N","PrintContactTitle":"N","invoiceAddressee":"1","showFooter":"Y","statementSummary":"Y","agingSummary":"Y","clientSelName":{"pKey":"","name":"","type":"client","whereClauseSearch":"N","isLegacy":"N","searchOptions":[{"name":"selectedResultIds","value":clientID,"type":"client","seq":1,"searchLevel":0,"valueDescription":""}]},"MarginAndImages":"[{\"ImageID\":\"Our Firm Block\",\"Type\":\"FirmAddress\",\"TopPosition\":0.07,\"LeftPosition\":2.23,\"ColBand\":\"Header\",\"ImageWidth\":0,\"ImageHeight\":0,\"Item\":1,\"Selected\":\"Y\"},{\"ImageID\":\"Client Address\",\"Type\":\"ClientAddress\",\"TopPosition\":1.55,\"LeftPosition\":0,\"ColBand\":\"Header\",\"ImageWidth\":0,\"ImageHeight\":0,\"Item\":2,\"Selected\":\"Y\"},{\"ImageID\":\"Date Block\",\"Type\":\"DateBlock\",\"TopPosition\":1.17,\"LeftPosition\":5.63,\"ColBand\":\"Header\",\"ImageWidth\":0,\"ImageHeight\":0,\"Item\":3,\"Selected\":\"Y\"},{\"ImageID\":\"Statement Label\",\"Type\":\"StatementLabel\",\"TopPosition\":0.03,\"LeftPosition\":0.04,\"ColBand\":\"Header\",\"ImageWidth\":0,\"ImageHeight\":0,\"Item\":4,\"Selected\":\"Y\"}]","ageDays1":"30","ageDays2":"60","ageDays3":"90","ageDays4":"120","ageDays5":"150","statementDate": STATEMENTDATE + "T00:00:00.000","_desc_saveOptionRole":["","","","",""],"saveOptionRole":["ACCOUNTING","ACCOUNTANT","CONTROLLER","CONTROLLER-RO","[CREATOR_USERNAME]"],"baseOriginalFavoriteId":"DA82BA5C8CD94E2CAD6A96814CC49C5C"}}

        response = requests.post(url, headers=HEADERS, json=payload  )
        data = _safe_json(response, f"Statement PDF build API for client {clientID}")
        report_path_raw = data["return"]["ReportPath"]
        report_path = report_path_raw.replace(" ", "%20")

        # Step 3: Get Nonce
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Security/Nonce"
        payload = {}
        response = requests.post(url, headers=HEADERS, json=payload  )
        nonce = _safe_json(response, f"Statement PDF nonce API for client {clientID}")

        # Step 4: Get Viewer
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx??&nonce="+nonce+"&reportPath="+report_path+"&allowSchedule=N&reportName=AR%20Statement"

        response = requests.get(url, headers=HEADERS )

        html = response.text

        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id     = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        session_token = re.search(r'_token="([^"]+)"', html)

        fileName = "Statement of account - " + clientName + " as of " + STATEMENTDATE + ".pdf"

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
                "&ContentDisposition=OnlyHtmlInline"
                f"&Format=PDF" )

        response = requests.get(url_pdf, headers = HEADERS,stream=True  )

        if response.status_code == 200:
            pdfName = os.path.join(ONEDRIVEDIR, WORKDIR, "ar_export", fileName)
            with open(pdfName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            return fileName
        
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/app/base/MakeVisionServiceRequest?method=DeleteAndStopReport"
        payload = {"sessionID": session_token.group(1), "reportPath": report_path_raw}
        response = requests.post(url, headers=HEADERS, json=payload )

        return ""
    except Exception as e:
        print("❌ Error in download statement of " + clientName)

def ar_billing_term(projectId):
    projectIdClear = util.cleanup_projectID(projectId)
    if projectIdClear == "":
        return None
    url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/BillingTermsBO/{projectIdClear}?meta=channel%2Caccess"
    response = requests.get(url, headers = HEADERS)
    data = _safe_json(response, f"Billing terms API for project {projectIdClear}")
    return int(data[0].get('DaysBeforeDue') or 30)

def ar_check_due(invoice):
    projectId = invoice['InvoiceMainWBS1']
    if projectId == "":
        return None
    daysBeforeDue = ar_billing_term(projectId)
    if daysBeforeDue == None:
        return None
    return daysBeforeDue < int(invoice['DaysOut'])

def ar_generate_invoices_table(clientID, dueInvoice):
    url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/ARReview/"+ clientID
    response = requests.get(url, headers = HEADERS)
    records = _safe_json(response, f"AR review list API for client {clientID}")
    message = ""

    AGE_FIELDS = ["Age1", "Age2", "Age3", "Age4", "Age5"]

    for r in records:
        if r['Total'] <= 0:
            continue
        projectID = r['WBS1']
        projectID = util.cleanup_projectID(projectID)

        url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/ARReview/"+ projectID +"/"+ clientID +"/ARReviewDetail"
        response = requests.get(url, headers=HEADERS)
        invoices = _safe_json(response, f"AR review detail API for project {projectID} and client {clientID}")
        rows = []
        due_invoices = []
        for invoice in invoices:
            due = ar_check_due(invoice)
            if due == None:
                continue
            aging_amounts = {field: invoice.get(field, 0) for field in AGE_FIELDS}

            if not any(amount > 0 for amount in aging_amounts.values()):
                continue

            if due:
                due_invoices.append(invoice.get("InvoiceNumber", ""))
            
            color = "red" if due else "inherit"

            row = [f'<tr><td align="center">{invoice.get("InvoiceNumber", "")}</td>']

            for amount in aging_amounts.values():
                row.append(f'<td align="right" style="color:{color};">{format_amount(amount)}</td>')
            row.append('</tr>')
            rows.append("".join(row))

        message += "".join(rows)
        if due:
            dueInvoice += invoice.get("InvoiceNumber", "") + ", "

    return message, dueInvoice

def ar_download_proj_invoice_pdf(projectID, clientId):
    projectIDConverted = util.cleanup_projectID(projectID)
    projectName = util.cleanup_projectName(projectID)
    url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/ARReview/"+ projectIDConverted +"/"+ clientId +"/ARReviewDetail"
    response = requests.get(url, headers=HEADERS)
    invoices = _safe_json(response, f"Project invoice detail API for project {projectIDConverted} and client {clientId}")
    for invoice in invoices: 
        if abs(invoice['Total']) < 1e-9:
            continue
        clientName = invoice['ClientName']
        clientName = clientName.replace("/"," ")
        try:
            url = "https://qcadeltek03.qcasystems.com/vantagepoint/app/Invoices/GetInvoiceFileInfo?invoiceMainWBS1="+invoice['InvoiceMainWBS1']+"&wbs1="+ projectName +"&invoiceNumber="+ invoice['InvoiceNumber'] +"&creditMemoRefno=&linkCompany="
            response = requests.get(url, headers=HEADERS)

            url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/InteractiveDetail/" + projectIDConverted + "/InvoiceHistory/"+ invoice['InvoiceNumber'] +"/print?printBackupReport=Y&printSupportDocuments=N&DownloadInvoice=N&creditMemo=&hasDraftInvoice=&applicationId=ARReview"
            response = requests.get(url, headers=HEADERS)

            if response.headers.get('Content-Type') == 'application/pdf' and response.content.startswith(b'%PDF'):
                fileName = invoice['InvoiceNumber'] + "_" + clientName + ".pdf"
                pdfName = os.path.join(ONEDRIVEDIR, WORKDIR, "ar_export", fileName)
                with open(pdfName, "wb") as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
                ar_create_record(clientId, "", "", fileName, "", "")
                continue

            data = _safe_json(response, f"Interactive detail API for invoice {invoice['InvoiceNumber']}")

            report_path_raw = data["ReportPath"]
            report_path = report_path_raw.replace(" ", "%20")

            url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Security/Nonce"
            payload = {}
            response = requests.post(url, headers=HEADERS, json=payload  )
            nonce = _safe_json(response, f"Invoice nonce API for invoice {invoice['InvoiceNumber']}")

             # Step 4: Get Viewer
            url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx??&nonce="+nonce+"&reportPath="+report_path+"&allowSchedule=N&reportName=Invoice&embedded=Y&runtimeParameters%5B0%5D%5BshowBillingBackup%5D=1&runtimeParameters%5B1%5D%5BPreInvoice%5D=N&runtimeParameters%5B2%5D%5BHeaderInvoice%5D="+invoice['InvoiceNumber']+"&runtimeParameters%5B3%5D%5BMainWBS1%5D="+projectName+"&runtimeParameters%5B4%5D%5BmainWBS1Name%5D=&runtimeParameters%5B5%5D%5BInvoice%5D="+invoice['InvoiceNumber']+"&runtimeParameters%5B6%5D%5BActivePeriod%5D="

            response = requests.post(url, headers=HEADERS)

            html = response.text

            report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
            control_id     = re.search(r"ControlID=([A-Za-z0-9]+)", html)
            session_token = re.search(r'_token="([^"]+)"', html)

            fileName = invoice['InvoiceNumber'] + "_" + clientName + ".pdf"

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
                    "&ContentDisposition=OnlyHtmlInline"
                    f"&Format=PDF" )

            response = requests.get(url_pdf, headers = HEADERS, stream=True  )

            if response.status_code == 200:
                pdfName = os.path.join(ONEDRIVEDIR, WORKDIR, "ar_export", fileName)
                with open(pdfName, "wb") as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
                ar_create_record(clientId, "", "", fileName,"", "", "")
            
        except Exception as e:
            print("❌ Error in download " + invoice['InvoiceNumber'] + " of " + clientName)
            continue

def ar_get_pm_list(clientID):
    url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/ARReview/" + clientID
    response = requests.get(url, headers=HEADERS)
    projList = _safe_json(response, f"PM list AR review API for client {clientID}")
    list = set()
    pmList = ""
    for proj in projList:
        if proj['Age2'] > 0 or proj['Age3'] > 0 or proj['Age4'] > 0 or proj['Age5'] > 0:
            projName = proj['WBS1']
            projName = util.cleanup_projectName(projName)
            url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/project/{projName}/plans/liveplan?missingjtd=no"
            response = requests.get(url, headers=HEADERS)
            data = _safe_json(response, f"Project plan API for project {projName}")[0]
            if data['ProjMgrEmail'] != '': 
                print(f"  ✅ PM {proj['WBS1']}: {data['ProjMgrEmail']}")
                if not data['ProjMgrEmail'] in list:
                    list.add(data['ProjMgrEmail'] )
                    pmList += data['ProjMgrEmail'] + ";"
            else:
                print(f"  ❌ PM email not found for {proj['WBS1']}")        
    return pmList

def ar_need_zip(clientID, clientName):
    count = 0
    curPath = os.path.join(ONEDRIVEDIR, WORKDIR, "ar_export")
    for file in os.listdir( curPath ):
        if clientName in file:
            count+=1
    return count > 5

def ar_zipfile(clientID, clientName):
    ws = wb["Sheet1"]
    count = 0
    global RECORDS
    for row in range(ws.max_row, 0, -1):
        if clientID == ws[f'A{row}'].value:
            count += 1
    if count > 5:
        for row in range(ws.max_row, 0, -1):
            if clientID == ws[f'A{row}'].value and ("Statement of account" not in str(ws[f'F{row}'].value or "")):
                ws.delete_rows(row)
                RECORDS = RECORDS - 1
        curPath = os.path.join(ONEDRIVEDIR, WORKDIR, "ar_export")
        zipFileName = "Invoices of " + clientName + ".zip"
        zipFilePath = os.path.join(curPath, zipFileName)
        with zipfile.ZipFile(zipFilePath, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file in os.listdir(curPath):
                if file == zipFileName:
                    continue
                if clientName in file and "Statement of account" not in file:
                    file_path = os.path.join(curPath, file)
                    arcname = os.path.relpath(file_path, curPath)
                    zipf.write(file_path, arcname=file)
        
        ar_create_record(clientID, clientName, "", zipFileName, "", "")
        timestamp = datetime.now().strftime("%Y%m%d")
        excelName = os.path.join(ONEDRIVEDIR, WORKDIR, f"Job_{timestamp}.xlsx")
        wb.save(excelName)

def ar_create_record(clientID, clientName, email, fileName, pmList, tableContent, dueInvoice = ""):
    ws = wb.active
    ws.title = "Sheet1"
    dueInvoiceInfo = ""
    if dueInvoice != "":
        dueInvoiceInfo = (
            "<p>We kindly remind you that the following invoices are due: "
            "<span style='background-color: yellow; color: red; font-weight: bold;'>"
            + dueInvoice +
            "</span> If you already paid this invoice or have any questions, let us know!</p>"
        )
    body = "<html><p>Dear <b>" + clientName + "</b></p><p> Please find the attached file for <b> Statement of account - " + clientName + " as of " + STATEMENTDATE + "</b>.</p>" + tableContent + dueInvoiceInfo + OPTIONALMSG + BODY
    row = [clientID, MAIL_FROM, email, pmList + CC, SUBJECT + clientName + " as of " + STATEMENTDATE, fileName, WORKDIR+"\\ar_export\\"+fileName, body, clientName]
    ws.append(row)
    global RECORDS
    RECORDS += 1

def main():
    ar_init()
    init_output()
    ar_download_csv()
    zipClientName = ar_process()
    
    for clientName in zipClientName:
        clientID = util.get_clientID(clientName)
        ar_zipfile(clientID, clientName)
    util.save_excel(wb, RECORDS, folder_path=os.path.join(ONEDRIVEDIR, WORKDIR))

if __name__ == "__main__":
    main()