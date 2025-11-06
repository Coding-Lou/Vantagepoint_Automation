from playwright.sync_api import sync_playwright
import util
import time, json
import os
import requests

USER_BROWSER_DIR = os.path.join(os.getcwd(), "edge_profile")
DOMAIN_FILTER = "qcasystems"
TOKEN = None
WWWBEARER = None
COOKIES = None

def set_wwwbearer():
    global TOKEN
    global WWWBEARER
    try:
        url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/token"
        headers = {
            "content-type": "application/json; charset=UTF-8"
        }
        payload = {
            "sessionID": TOKEN,
            "grant_type": "password",
            "database": "VPProduction (QCADELTEK03)"
        }
        response = requests.post(url, headers = headers, data=payload)
        data = response.json()
        WWWBEARER = data["access_token"]
        util.set_config("WWWBEARER", WWWBEARER)
    except Exception as e:
        print("❌ Login Failed, Please check the cookie and token in the configuration file.")

def set_token_cookies(request):
    global TOKEN
    global COOKIES
    if "/vision/token" not in request.url:
        return
    
    # === retrieve token ===
    for k, v in request.headers.items():
        if "token" in k.lower():
            TOKEN = v
            util.set_config("TOKEN", TOKEN)

    # === retrieve cookie ===
    try:
        context = request.frame.page.context
        cookies = context.cookies()

        # filter domain cookies
        filtered = [c for c in cookies if DOMAIN_FILTER in c["domain"]]

        deduped = {}
        for c in filtered:
            deduped[c["name"]] = c

        COOKIES = "; ".join([f"{c['name']}={c['value']}" for c in deduped.values()])
        util.set_config("COOKIES", COOKIES)

    except Exception as e:
        print("⚠️ Failed to get cookies:", e)

def set_asp_net_cookie():
    global TOKEN
    global COOKIES
    global WWWBEARER
    # Build
    url = "https://qcadeltek03.qcasystems.com/Vantagepoint/vision/Reporting/Build"
    headers = util.set_headers()
    payload = {"reportPath":"/Standard/AccountingGeneral/Remittance Advice","reportOptions":{"BankCode":"RBC-CDN","period":202607,"PostSeq":22,"ShowSSN":"N","checkpayee":"2","vendor":"UPSCAN","Employee":"","CheckNo":"'700000278'"}}
    response = requests.post(url, headers=headers, json=payload)
    data = response.json()
    report_path_raw = data["return"]["ReportPath"]
    report_path = report_path_raw.replace(" ", "%20")

    # Get Nonce
    url = "https://qcadeltek03.qcasystems.com/vantagepoint/vision/Security/Nonce"
    payload = {}
    response = requests.post(url, headers=headers, json=payload)
    nonce = response.json()

    # Get Viewer
    url = "https://qcadeltek03.qcasystems.com/vantagepoint/reporting/viewer.aspx?&nonce="+nonce+"&reportPath="+report_path+"&allowSchedule=N&reportName=Remittance%20Advice"
    response = requests.get(url, headers=headers)

    asp_net_cookie = response.headers.get("Set-Cookie").split(";", 1)[0]
    COOKIES = COOKIES + ";" + asp_net_cookie

    print(f"ASP_NET Cookie: " + asp_net_cookie)
    print(f"Current Cookies: {COOKIES}")
    util.set_config("COOKIES", COOKIES)
 
def sso_login():
    global USER_BROWSER_DIR
    with sync_playwright() as p:
        edge = p.chromium
        context = edge.launch_persistent_context(
            USER_BROWSER_DIR,
            headless=False,
            channel="msedge"
        )
        page = context.new_page()
        page.on("request", set_token_cookies)
        page.goto("https://qcadeltek03.qcasystems.com/Vantagepoint/app")
        time.sleep(5)
        context.close()

    set_wwwbearer()
    set_asp_net_cookie()