from playwright.sync_api import sync_playwright
import time

USER_DATA_DIR = "C://Users//jlou//OneDrive - QCA Systems Ltd//Documents//Automation//edge_profile"

with sync_playwright() as p:
    edge = p.chromium
    context = edge.launch_persistent_context(
        USER_DATA_DIR,
        headless=False,
        channel="msedge"
    )
    page = context.new_page()
    page.goto("https://qcadeltek03.qcasystems.com/Vantagepoint/app")
    print(page.title())

    def log_request(request):
       if "/Vantagepoint/vision/token" in request.url: 
           print(f"Request URL: {request.url}")
           print("Request headers:")
           for k, v in request.headers.items():
               print(f"  {k}: {v}")
           print("=" * 60)

    page.on("request", log_request)
    time.sleep(10)
    context.close()
