import tools.util as util
import requests
import os
from datetime import datetime, date
from openpyxl import Workbook
from zoneinfo import ZoneInfo
import core.packing_slip as packing_slip

def get_master_key(po):
    try:
        url = f"https://qcadeltek03.qcasystems.com/vantagepoint/vision/PurchaseReceive/POMaster/?searchType=ALL&filter={po}&page=1&pagesize=100&order=name"

        response = requests.get(url, headers=HEADERS )
        data = response.json()

        if len(data) == 0:
            return ""
        return data[0]
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
    
def get_shiping_data(masterKey):
    try:
        url = f"https://qcadeltek03.qcasystems.com/vantagepoint/vision/PurchaseReceive/ReceiveDocuments/{masterKey}"

        response = requests.get(url, headers=HEADERS)
        data = response.json()
        return data
    except Exception as e:
        print(f"Get shipping data from masterkey {masterKey} had error {e}")

def download_packing_list(shippingData):
    try:
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Security/Nonce"
        response = requests.get(url, headers=HEADERS)
        nonce = response.json

        url = f"https://qcadeltek03.qcasystems.com/vantagepoint/vision//PurchaseReceive/ReceiveDocumentsDetail/{shippingData["MasterPKey"]}%7C{shippingData["FileID"]}%7C{shippingData["DetailPKey"]}/Documents/{shippingData["FileID"]}?filename=&nonce={nonce}=&active_period="
        response = requests.get(url, headers=HEADERS)

        if response.status_code == 200:
            pdfName = os.path.join("packing_slip", shippingData["FileName"])
            with open(pdfName, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            #print("✅ "+ shippingData["FileName"] + " Downloaded")
        else:
            print(f"❌ Error")
    except Exception as e:
        print(f"Downloading {shippingData["FileName"]} had error {e}")

def main():
    global HEADERS
    HEADERS = util.set_headers()
    date = input("Please input the start date (format: 2025-04-01): ")
    util.check_folder("packing_slip")
    util.clear_folder("packing_slip")
    wb = Workbook()
    ws = wb.active
    ws.append(["PO","Vendor", "PackingSlip", "Shipping Date", "Receiving Date", "Days Difference", "Packing List"])
    po = 24139
    while po > 20000:
        try:
            masterData = get_master_key(po)
            po -= 1
            if (masterData == ""):
                continue
            masterKey = masterData["MasterPKey"]
            vendorName = masterData["VendorName"]
            receivingList = get_receiving_list(masterKey)
            receivingDict = {}
            packingDict = {}

            for receiving in receivingList:
                receivingDate = receiving["ReceiveDate"]
                dt_utc = datetime.fromisoformat(receivingDate).replace(tzinfo=ZoneInfo("UTC"))
                dt_van = dt_utc.astimezone(ZoneInfo("America/Vancouver"))
                receivingDate = datetime.strptime(dt_van.strftime("%Y-%m-%d"), "%Y-%m-%d")
                if receivingDate < datetime.strptime(date, "%Y-%m-%d"):
                    continue
                receivingDict[receiving["PKey"]] = receivingDate
                packingDict[receiving["PKey"]] = receiving["PackingSlip"]

            shippingData = get_shiping_data(masterKey)
                #vendor_pattern  = re.compile(r"\b(" + "|".join(VENDORS) + r")\b", re.IGNORECASE)

            for shipping in shippingData:
                try:
                    if (shipping["DetailPKey"] not in receivingDict):
                        continue
                    receivingDate = receivingDict[shipping["DetailPKey"]]
                    download_packing_list(shipping)
                    fromAI = packing_slip.main(shipping["FileName"], receivingDate)
                    if fromAI == "Not found" :
                        print(f"{po+1} | {vendorName} | {packingDict[shipping["DetailPKey"]]} | | { receivingDate.strftime("%Y-%m-%d")} | ---------- | {shipping["FileName"]} ")
                        ws.append([po+1,vendorName, packingDict[shipping["DetailPKey"]], "" , receivingDate.strftime("%Y-%m-%d"), "" , shipping["FileName"]])
                        continue
                    shippingDate = datetime.strptime(fromAI, "%Y-%m-%d")
                    diff = abs((shippingDate - receivingDate).days)
                    print(f"{po+1} | {vendorName} | {packingDict[shipping["DetailPKey"]]} | {shippingDate.strftime("%Y-%m-%d")} | { receivingDate.strftime("%Y-%m-%d")} | {diff} | {shipping["FileName"]} ")
                    ws.append([po+1,vendorName, packingDict[shipping["DetailPKey"]], shippingDate.strftime("%Y-%m-%d"), receivingDate.strftime("%Y-%m-%d"), diff, shipping["FileName"]])
                    
                except Exception as e:
                    print(f"{po} had error {e}")
                    continue
        except Exception as e:
            print(f"{po} had error {e}")
            continue
   

    wb.save("packing_slip/output.xlsx")

if __name__=="__main__":
    main()
