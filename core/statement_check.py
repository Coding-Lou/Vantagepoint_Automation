from datetime import datetime
from zoneinfo import ZoneInfo
import requests
import tools.util as util


def check_statement(vendor_key: str, invoice_number: str) -> bool | None:
    """Query the ERP for a single invoice and print its payment status."""
    # Always fetch fresh headers to avoid stale auth credentials
    headers = util.set_headers()
    year = datetime.now(ZoneInfo("America/Vancouver")).year

    url = (
        f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/VendorReview"
        f"/{vendor_key}/Vouchers?startRow=1&pageSize=50&initialLoad=true"
        f"&sortDef=InvoiceDate_D"
        f"&filterHash%5B0%5D%5Bname%5D=Invoice"
        f"&filterHash%5B0%5D%5Bvalue%5D={invoice_number}"
        f"&filterHash%5B0%5D%5Btype%5D=string"
        f"&filterHash%5B0%5D%5Bseq%5D=1001"
        f"&filterHash%5B0%5D%5Bopp%5D=LIKE"
        f"&filterHash%5B0%5D%5BsearchLevel%5D=0"
        f"&filterHash%5B0%5D%5BgridColumnFilter%5D=true"
        f"&PaidStatus=A&Period={year}01&offset=0"
    )

    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        print(f"{invoice_number} - Error: HTTP {response.status_code} (authentication may be expired, please re-login)")
        return None

    try:
        data = response.json()
    except Exception:
        # Non-JSON response usually means the session has expired
        print(f"{invoice_number} - Error: Invalid response (authentication may be expired, please re-login)")
        return None

    if len(data) == 0:
        return False

    cheque_date_str = data[0]["CheckDate"]
    if not cheque_date_str or cheque_date_str.strip() == "":
        cheque_date_str = "will be paid on next payment run"
    else:
        cheque_date_str = f" paid on {cheque_date_str[:10]}"

    print(f"{invoice_number} voucher number {data[0]['Voucher']} {cheque_date_str}")
    return True


def main(vendor_key: str = None, invoice_numbers: str = None):
    """
    Check payment status for one or more invoices against a vendor account.

    Args:
        vendor_key: Vantagepoint vendor key (prompts if omitted)
        invoice_numbers: Comma-separated invoice numbers (prompts if omitted)
    """
    login_check = util.check_login()

    if vendor_key is None:
        vendor_key = input("Enter vendor key: ")
    if invoice_numbers is None:
        invoice_numbers = input("Enter invoice numbers separated by commas: ")

    # Verify the session is still valid before processing the full list
    if login_check:
        headers = util.set_headers()
        test_url = (
            f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/VendorReview"
            f"/{vendor_key}/Vouchers?startRow=1&pageSize=1&initialLoad=true"
        )
        try:
            test_response = requests.get(test_url, headers=headers, timeout=10)
            if test_response.status_code != 200:
                print("Warning: Authentication may be expired. Please re-login if you get incorrect results.")
        except Exception:
            print("Warning: Could not verify authentication. Please re-login if you get incorrect results.")

    not_vouched_list = []
    invoice_number_list = invoice_numbers.split(",") if invoice_numbers else []

    for invoice_number in invoice_number_list:
        result = check_statement(vendor_key.strip(), invoice_number.strip())
        if not result:
            not_vouched_list.append(invoice_number)

    print()
    print(f"Total not vouched invoices: {len(not_vouched_list)}")
    print("--------------------------------")
    for invoice_number in not_vouched_list:
        print(f"{invoice_number} not been vouched")


if __name__ == "__main__":
    main()
