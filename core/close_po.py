import csv
import io
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
global folder_path
folder_path = r"C:\\temp\\weekly_open_po"

def retrieve_open_po_report():
    try:
        # Step 1: Build
        url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
        payload = {"reportPath":"/Standard/Purchasing/Open PO","reportOptions":{"baseAlternateRowColor":"","baseBottomMargin":0.5,"baseCulture":"default","baseDefaultCurrencyFormat":"###T###T###D##;(###T###T###D##);#","baseDefaultDateFormat":"M/d/yyyy","baseDefaultHTMLFormatting":"Y","baseDefaultNumberFormat":"###T###T###D##;-###T###T###D##;#","baseFont":"Arial","baseFooterText":"[version] - [options]","baseGridTable":"","baseGroupIndent":0.1,"baseHeadingEndDate":"","baseHeadingRowColor":"","baseHeadingStartDate":"","baseHideDocumentMap":"Y","baseHideSingleLineTotals":"N","baseLeftMargin":0.5,"defaultPage2Top":0,"baseOrientation":"automatic","baseOverrideHeadingDate":"N","basePageHeight":11,"basePageSize":"letter","basePageWidth":8.5,"baseReportName":"Open Purchase Orders","baseRightMargin":0.5,"baseShowBorderLines":"N","baseShowFinalTotals":"N","baseShowTotalsOnHeader":"N","baseStartColumnPosition":0.1,"baseTopMargin":0.5,"baseUnitOfMeasure":"in","baseUseDashpartLayout":"N","baseUseLookupFilterToGrid":"N","ReportGroups":[{"label":"PO Number","sort":"ASC","color":"000080","subTotal":"Y","showHeading":"Y","pageHeading":"N","collapseExpand":"D","line":"None","pageBreak":"N","groupID":"PONumber","customGridColumnSort":""}],"ReportColumns":[{"heading":"Order Date","width":0.7,"format":"M/d/yyyy","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"OrderDate","username":"","customGridColumnSort":""},{"heading":"PO Number","width":0.7,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"PONumber","username":"","customGridColumnSort":""},{"heading":"Changed Order Number","width":0.9,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"CONumber","username":"","customGridColumnSort":""},{"heading":"Seq","width":0.55,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Seq","username":"","customGridColumnSort":""},{"heading":"Unit Price","width":0.5,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"UnitPrice","username":"","customGridColumnSort":""},{"heading":"Qty Ordered","width":0.8,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"OrderQty","username":"","customGridColumnSort":""},{"heading":"Amount Ordered","width":0.75,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"TotalAmount","username":"","customGridColumnSort":""},{"heading":"Qty Open","width":0.8,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Qty","username":"","customGridColumnSort":""},{"heading":"Total Open Amount","width":0.75,"format":"###T###T###D##;(###T###T###D##);#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"TotalOpenAmount","username":"","customGridColumnSort":""},{"heading":"Due Date","width":0.6,"format":"M/d/yyyy","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"DueDate","username":"","customGridColumnSort":""},{"heading":"Status","width":0.7,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Status","username":"","customGridColumnSort":""},{"heading":"Receive Status","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ReceiveStatus","username":"","customGridColumnSort":""},{"heading":"Qty Received","width":0.625,"format":"###T###T###D##;-###T###T###D##;#","align":"right","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"ReceivedQty","username":"","customGridColumnSort":""},{"heading":"Date Accepted","width":0.625,"format":"M/d/yyyy","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"receiveDate","username":"","customGridColumnSort":""},{"heading":"Closed","width":0.7,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"closed","username":"","customGridColumnSort":""},{"heading":"Vendor","width":0.75,"format":"","align":"left","sectionName":"Section 1","sectionRow":0,"sectionColumn":1,"columnID":"Vendor","username":"","customGridColumnSort":""}],"ReportSections":[],"baseRecordSelection":"","baseShowDetail":"Y","OrderGroup":"1","DeliveryGroup":"2","timeFrameOption":"1","chkAging":"Y","SpecificDate":"11/13/2024 12:24:07 PM","saveOptionRole":"[CREATOR_USERNAME]","baseOriginalFavoriteId":"93D53D3654C14B378792E8B88D545099","baseSelectionRows":24703}}

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
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&ResetReportViewerOnPreview=Y&reportPath="+report_path+"&allowSchedule=Y&origReportPath=/Standard/Purchasing/Open%20PO&reportName=Open%20Purchase%20Orders"

        # Step 4: Get report session
        response = requests.get(url, headers=HEADERS )
        html = response.text
        report_session = re.search(r"ReportSession=([A-Za-z0-9]+)", html)
        control_id = re.search(r"ControlID=([A-Za-z0-9]+)", html)
        sqlrsReportViewer = re.search(r'_token="([^"]+)"', html)

        if not (report_session and control_id):
            raise RuntimeError("Error")
        
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
                f"&FileName=Open+Purchase+Orders&ContentDisposition=OnlyHtmlInline&Format=CSV" )
        
        response = requests.get(url, headers = HEADERS,stream=True  )
        if response.status_code == 200 and response.headers.get("Content-Type") == "text/csv; charset=utf-8":
            vendors = set()
            content = response.content.decode("utf-8")
            lines = content.splitlines()

            start = None
            for i, line in enumerate(lines):
                if line.startswith("groupHeader1_GroupColumn"):
                    start = i
                    break
                
            detail_csv_text = "\n".join(lines[start:])
            csv_reader = csv.DictReader(io.StringIO(detail_csv_text))

            for row in csv_reader:
                vendors.add(row["detail_Vendor"])
            
            print(f"Total {len(vendors)} vendors has open po.")

            return list(vendors)

    except Exception as e:
        print("⚠️ Failed to download the Close PO:", e)

def get_2way_result(vendor = ""):

    try:
        url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/DataEntry/OpenPOVoucherDetail/b2b2b1937e7c4e5cae198ae644e0e509/{vendor}/2"
        response = requests.get(url, headers=HEADERS)
        data = response.json()
        return data

    except Exception as e:
        print(f"⚠️ Failed to retrieve {vendor}: {e}")

def get_project_information(po_masterKey = ""):
    try:
        url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/PurchaseReceive/POMaster/{po_masterKey}?meta=access"
        response = requests.get(url, headers=HEADERS)
        data = response.json()
        po_number = data[0]["PONumber"]

        url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/PurchaseOrder/POMaster/DefaultDistribution/{po_masterKey}"
        response = requests.get(url, headers=HEADERS)
        data = response.json()
        project = data[0]["WBS1"]
        client = data[0]["ClientName"]

        url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/project/{project}?meta=channel%2Caccess"
        response = requests.get(url, headers=HEADERS)
        data = response.json()
        biller = data[0]["BillerName"]
        status = data[0]["Status"]

        url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/PurchaseReceive/POMaster/{po_masterKey}/ReceiveMaster"
        response = requests.get(url, headers=HEADERS)
        data = response.json()
        last_received_date = data[0]["ReceiveDate"]

        result = {
            "PO_Number": po_number,
            "project": project,
            "client": client,
            "Status": status,
            "biller": biller,
            "last_received_date": last_received_date
        }

        print(f"✅ Retrieved project information of PO {po_number}.")
        return result
    
    except Exception as e:
        print(f"⚠️ Failed to retrieve project information of PO {po_masterKey}: {e}")
        return None


def get_all_vendors_result(vendors: list):
    all_data = []
    all_pos = set()
    all_po_info = []

    for vendor in vendors:
        data = get_2way_result(vendor)

        if isinstance(data, dict):
            data = data.get("data", [])
            
        if data:
            all_data.extend(data)
            for item in data:
                key = item.get("POMasterPKey")
                if key:
                    all_pos.add(key)
        
        print(f"✅ Retrieved {len(data)} rows for vendor {vendor}.")

    for master_key in all_pos:
        row = get_project_information(master_key)
        # print(row)
        if row:
            all_po_info.append(row)
        else:
            print(f"⚠️ Skipped PO {master_key}: no project information returned.")

    return all_data, all_po_info

def export_to_csv(all_data, all_po_info):

    # 1. collect all keys
    fieldnames = sorted({k for row in all_data for k in row.keys()})

    # 2. write csv
    raw_data_filename = os.path.join(folder_path, "Close_po_raw_data_"+ date.today().strftime("%Y-%m-%d") + ".csv")
    
    with open(raw_data_filename, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)

        writer.writeheader()
        writer.writerows(all_data)

    print(f"✅ Exported {len(all_data)} rows -> {raw_data_filename}")

    # 1. collect all keys
    fieldnames = [
        "PO_Number",
        "project",
        "client",
        "Status",
        "biller",
        "last_received_date"
    ]

    # 2. write csv
    po_info_filename = os.path.join(folder_path, "po_info_"+ date.today().strftime("%Y-%m-%d") + ".csv")
    with open(po_info_filename, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)

        writer.writeheader()
        writer.writerows(all_po_info)

    print(f"✅ Exported {len(all_po_info)} rows -> {po_info_filename}")

    return raw_data_filename, po_info_filename

def download_templete(templete_abspath):
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
    file_name = "Weekly Clear Open PO_" + date.today().strftime("%Y-%m-%d") + ".xlsx"
    save_dir = str(save_path.absolute())
    util.download_from_gdrive(file_id= "1Zzexus_07LF1ZtyGg-j8lg6zLNQARoOE", file_name=file_name, save_dir=save_dir)
    global targetFile
    targetFile = os.path.join(save_path, file_name)
    return targetFile

def generate_report(raw_data_filename, po_info_filename, targetFile):
    # Copy from Template
    util.excel_full_copy(inputFile = raw_data_filename, inputSheet = None, targetFile = targetFile, targetSheet = "raw_data", onlyValue = True)
    util.excel_full_copy(inputFile = po_info_filename, inputSheet = None, targetFile = targetFile, targetSheet = "po_infomation", onlyValue = True)

    source = folder_path
    onedrivedir = util.get_config(["ONEDRIVEDIR"])
    workdir = util.get_config(["WORKDIR"])

    target = os.path.join(onedrivedir, workdir, "Weekly Open Po")
    util.move_and_replace_files(source_dir=source, target_dir=target)

def main():
    util.check_folder("Weekly Open Po")
    
    util.clear_folder("Weekly Open Po")
    targetFile = download_templete(folder_path)

    vendors = retrieve_open_po_report()
 
    all_data, all_po_info = get_all_vendors_result(vendors)
    
    raw_data_filename, po_info_filename, = export_to_csv(all_data, all_po_info)

    generate_report(raw_data_filename, po_info_filename, targetFile)

