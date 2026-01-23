from datetime import datetime
from zoneinfo import ZoneInfo
import requests
import tools.util as util
from pathlib import Path
import json
import time

#region agent log
DEBUG_LOG_PATH = Path(r"c:\cursor\.cursor\debug.log")

def _agent_log(hypothesis_id: str, location: str, message: str, data: dict = None):
    payload = {
        "sessionId": "debug-session",
        "runId": "pre-fix",
        "hypothesisId": hypothesis_id,
        "location": location,
        "message": message,
        "data": data or {},
        "timestamp": int(time.time() * 1000),
    }
    try:
        DEBUG_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with DEBUG_LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(json.dumps(payload, ensure_ascii=False) + "\n")
    except Exception:
        pass
#endregion

def init():
    global DATE
    # No longer setting global HEADERS - will get fresh headers for each request
    #region agent log
    _agent_log("H1", "statement_check.init", "init() called", {
        "note": "HEADERS will be fetched fresh for each request"
    })
    #endregion

def check_statement(vendor_key, invoice_number):
    # Get fresh headers for each request to ensure they're not stale
    headers = util.set_headers()
    #region agent log
    _agent_log("H4", "statement_check.check_statement", "check_statement() entry", {
        "vendor_key": vendor_key,
        "invoice_number": invoice_number,
        "headers_token": headers.get("Token", "")[:20] + "..." if headers.get("Token") else None,
        "headers_fresh": True,
    })
    #endregion
    year = datetime.now(ZoneInfo("America/Vancouver")).year
    url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/VendorReview/{vendor_key}/Vouchers?startRow=1&pageSize=50&initialLoad=true&sortDef=InvoiceDate_D&filterHash%5B0%5D%5Bname%5D=Invoice&filterHash%5B0%5D%5Bvalue%5D={invoice_number}&filterHash%5B0%5D%5Btype%5D=string&filterHash%5B0%5D%5Bseq%5D=1001&filterHash%5B0%5D%5Bopp%5D=LIKE&filterHash%5B0%5D%5BsearchLevel%5D=0&filterHash%5B0%5D%5BgridColumnFilter%5D=true&PaidStatus=A&Period={year}01&offset=0"
    #region agent log
    _agent_log("H4", "statement_check.check_statement", "before request.get", {
        "url": url[:100] + "..." if len(url) > 100 else url,
    })
    #endregion
    response = requests.get(url, headers=headers)
    #region agent log
    _agent_log("H4", "statement_check.check_statement", "after request.get", {
        "status_code": response.status_code,
        "response_length": len(response.text) if response.text else 0,
        "response_preview": response.text[:200] if response.text else None,
        "using_fresh_headers": True,
    })
    #endregion
    
    # Check if response indicates authentication issue
    # If status is 200 but response is empty or contains error, might be auth issue
    if response.status_code == 200:
        try:
            data = response.json()
        except:
            # If response is not valid JSON, might be an error page
            #region agent log
            _agent_log("H7", "statement_check.check_statement", "Invalid JSON response - possible auth issue", {
                "response_preview": response.text[:200] if response.text else None,
            })
            #endregion
            print(f"{invoice_number} - Error: Invalid response (authentication may be expired, please re-login)")
            return
    else:
        #region agent log
        _agent_log("H7", "statement_check.check_statement", "Non-200 status - possible auth issue", {
            "status_code": response.status_code,
        })
        #endregion
        print(f"{invoice_number} - Error: HTTP {response.status_code} (authentication may be expired, please re-login)")
        return
    
    #region agent log
    _agent_log("H4", "statement_check.check_statement", "after response.json()", {
        "data_length": len(data) if isinstance(data, list) else 0,
        "data_type": type(data).__name__,
    })
    #endregion
    
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
    #region agent log
    _agent_log("H1", "statement_check.main", "main() entry", {
        "has_vendor_key": bool(vendor_key),
        "has_invoice_numbers": bool(invoice_numbers),
    })
    #endregion
    
    #region agent log
    login_check = util.check_login()
    _agent_log("H2", "statement_check.main", "check_login() result", {
        "is_logged_in": bool(login_check),
        "user_email": login_check if login_check else None,
    })
    #endregion
    
    init()
    
    if vendor_key is None:
        vendor_key = input("Enter vendor key: ")
    if invoice_numbers is None:
        invoice_numbers = input("Enter invoice numbers separated by commas: ")
    
    # Verify authentication is actually valid by making a test API request
    # This catches cases where check_login() passes but actual API requests fail
    if login_check:
        #region agent log
        _agent_log("H7", "statement_check.main", "Verifying auth with test API request", {})
        #endregion
        headers = util.set_headers()
        # Make a simple test request to verify authentication works
        test_url = f"https://qcadeltek03.qcasystems.com/Vantagepoint/vision/VendorReview/{vendor_key}/Vouchers?startRow=1&pageSize=1&initialLoad=true"
        try:
            test_response = requests.get(test_url, headers=headers, timeout=10)
            #region agent log
            _agent_log("H7", "statement_check.main", "Test API request result", {
                "status_code": test_response.status_code,
                "response_length": len(test_response.text) if test_response.text else 0,
            })
            #endregion
            # If we get a non-200 status or the response seems invalid, auth might be expired
            if test_response.status_code != 200:
                print("⚠️ Warning: Authentication may be expired. Please re-login if you get incorrect results.")
        except Exception as e:
            #region agent log
            _agent_log("H7", "statement_check.main", "Test API request exception", {
                "error": str(e),
            })
            #endregion
            print("⚠️ Warning: Could not verify authentication. Please re-login if you get incorrect results.")
    
    invoice_number_list = invoice_numbers.split(',') if invoice_numbers else []
    for invoice_number in invoice_number_list:
        check_statement(vendor_key.strip(), invoice_number.strip())

if __name__ == "__main__":
    main()