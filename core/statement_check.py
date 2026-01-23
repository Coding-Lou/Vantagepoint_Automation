from datetime import datetime
from zoneinfo import ZoneInfo
import requests
import tools.util as util

def init():
    global DATE
    global HEADERS
    HEADERS = util.set_headers()

def check_statement(vendor_key, invoice_number):
    year = datetime.now(ZoneInfo("America/Vancouver")).year
    url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/VendorReview/{vendor_key}/Vouchers?startRow=1&pageSize=50&initialLoad=true&sortDef=InvoiceDate_D&filterHash%5B0%5D%5Bname%5D=Invoice&filterHash%5B0%5D%5Bvalue%5D={invoice_number}&filterHash%5B0%5D%5Btype%5D=string&filterHash%5B0%5D%5Bseq%5D=1001&filterHash%5B0%5D%5Bopp%5D=LIKE&filterHash%5B0%5D%5BsearchLevel%5D=0&filterHash%5B0%5D%5BgridColumnFilter%5D=true&PaidStatus=A&Period={year}01&offset=0"
    response = requests.get(url, headers=HEADERS)
    data = response.json()
    if len(data) == 0 :
        print(f"{invoice_number} not been vouched")
    else:
        print(f"{invoice_number} voucher number {data[0]['Voucher']} payment date {data[0]['PaymentDate'][:10]} ")

def main(vendor_key: str = None, invoice_numbers: str = None):
    """
    Main function for statement check.
    
    Args:
        vendor_key: Vendor key to check (if None, will prompt for input)
        invoice_numbers: Comma-separated invoice numbers (if None, will prompt for input)
    """
    init()
    
    if vendor_key is None:
        vendor_key = input("Enter vendor key: ")
    if invoice_numbers is None:
        invoice_numbers = input("Enter invoice numbers separated by commas: ")
    
    invoice_number_list = invoice_numbers.split(',') if invoice_numbers else []
    for invoice_number in invoice_number_list:
        check_statement(vendor_key.strip(), invoice_number.strip())

if __name__ == "__main__":
    main()