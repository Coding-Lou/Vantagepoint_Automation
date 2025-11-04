from playwright.sync_api import sync_playwright
import util
import time, json

USER_DATA_DIR = (
    "C://Users//jlou//OneDrive - QCA Systems Ltd//Documents//Automation//edge_profile"
)
DOMAIN_FILTER = "qcadeltek03.qcasystems.com"

def log_request(request):
    if "/vision/token" not in request.url:
        return
    
    # === retrieve token ===
    token = None
    for k, v in request.headers.items():
        if "token" in k.lower():
            token = v
            util.update_config("TOKEN", token)

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
        util.update_config("COOKIE", cookie_str)

    except Exception as e:
        print("⚠️ Failed to get cookies:", e)
        

def Sso_login():
    with sync_playwright() as p:
        edge = p.chromium
        context = edge.launch_persistent_context(
            USER_DATA_DIR,
            headless=False,
            channel="msedge"
        )
        page = context.new_page()
        page.on("request", log_request)
        page.goto("https://qcadeltek03.qcasystems.com/Vantagepoint/app")

        time.sleep(10)
        
        context.close()
    
