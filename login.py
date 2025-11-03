from playwright.sync_api import sync_playwright
import time, json

USER_DATA_DIR = "C://Users//jlou//OneDrive - QCA Systems Ltd//Documents//Automation//edge_profile"

def update_token(token):
    print(token)
    try:
        with open("Vantagepoint_Automation\config.json", "r", encoding="utf-8") as f:
            config = json.load(f)
        
        config["TOKEN"] = token
        
        with open("Vantagepoint_Automation\config.json", "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
            f.flush()

        print("💾 Token has been updated. ")

    except Exception as e:
        print("⚠️ Token updated failed: ", e)


def Sso_login():
    with sync_playwright() as p:
        edge = p.chromium
        context = edge.launch_persistent_context(
            USER_DATA_DIR,
            headless=False,
            channel="msedge"
        )
        page = context.new_page()
        page.goto("https://qcadeltek03.qcasystems.com/Vantagepoint/app")

        def log_request(request):
            if "/vision/token" in request.url: 
                for k, v in request.headers.items():
                    if ("token" in k):
                        token = v
                update_token(token)
                
        page.on("request", log_request)
        time.sleep(10)
        context.close()
    
    
