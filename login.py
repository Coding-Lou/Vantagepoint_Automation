from playwright.sync_api import sync_playwright
import util
import time, json
import os
import requests

USER_BROWSER_DIR = os.path.join(os.getcwd(), "edge_profile")
DOMAIN_FILTER = "qcasystems"
TOKEN = None

def set_token_cookies(request):
    if "/vision/token" not in request.url:
        return
    
    # === retrieve token ===
    for k, v in request.headers.items():
        if "token" in k.lower():
            TOKEN = v
            util.update_config("TOKEN", TOKEN)

    # === retrieve cookie ===
    try:
        context = request.frame.page.context
        cookies = context.cookies()

        # filter domain cookies
        filtered = [c for c in cookies if DOMAIN_FILTER in c["domain"]]

        deduped = {}
        for c in filtered:
            deduped[c["name"]] = c

        cookie_str = "; ".join([f"{c['name']}={c['value']}" for c in deduped.values()])
        util.update_config("COOKIES", cookie_str)

    except Exception as e:
        print("⚠️ Failed to get cookies:", e)

def set_wwwbearer():
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
        wwwbearer = data["access_token"]
        util.update_config("WWWBEARER", wwwbearer)

    except Exception as e:
        print("❌ Login Failed, Please check the cookie and token in the configuration file.")


def Sso_login():
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

        time.sleep(3)
        
        context.close()
    
    set_wwwbearer()
    
