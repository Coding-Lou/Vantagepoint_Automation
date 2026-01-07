from tools.runtime_logger import RuntimeLogger
import tools.util as util
import tools.login as login
import tools.schedule_task as schedule_task
import core.ap as ap
import core.ar as ar
import core.project_status as project_status
import core.bridge_report as bridge_report
import core.shipping_monitor as shipping_monitor
import core.daily_received as daily_received
import core.on_call as on_call
import time
from pathlib import Path
import argparse

# =========================
# Set up logger
# =========================
BASE_DIR = Path(__file__).resolve().parent
LOGS_DIR = BASE_DIR / "logs"
LOGS_DIR.mkdir(parents=True, exist_ok=True)
log_file = LOGS_DIR / "runtime.log"

def run_task_by_flag(args):
    LOGIN = util.check_login()

    if args.ap:
        while not LOGIN:
            login.sso_login()
            LOGIN = util.check_login()
        ap.main()
    elif args.ar:
        while not LOGIN:
            login.sso_login()
            LOGIN = util.check_login()
        ar.main()
    elif args.project_status:
        while not LOGIN:
            login.sso_login()
            LOGIN = util.check_login()
        project_status.main()
    elif args.bridge:
        while not LOGIN:
            login.sso_login()
            LOGIN = util.check_login()
        bridge_report.main()
    elif args.shipping:
        while not LOGIN:
            login.sso_login()
            LOGIN = util.check_login()
        shipping_monitor.main()
    elif args.daily_received:
        while not LOGIN:
            login.sso_login()
            LOGIN = util.check_login()
        daily_received.main()
    elif args.on_call:
        on_call.run_oncall_task(vendorName = args.name)
    elif args.init_schedule:
        schedule_task.main()
    else:
        return False

    return True

def main():
    parser = argparse.ArgumentParser(description="Vantagepoint Automation Tool")
    parser.add_argument("--ap", action="store_true", help="Run AP task")
    parser.add_argument("--ar", action="store_true", help="Run AR task")
    parser.add_argument("--project_status", action="store_true", help="Run Project Status task")
    parser.add_argument("--bridge", action="store_true", help="Run Bridge Report task")
    parser.add_argument("--shipping", action="store_true", help="Run Shipping Monitor task")
    parser.add_argument("--daily_received", action="store_true", help="Run Daily Receiving task")
    parser.add_argument("--on_call", action="store_true", help="Run On Call task")
    parser.add_argument("--name",type=str,help="Vendor name for on-call task (e.g. Pembina)")

    parser.add_argument("--init_schedule", action="store_true", help="Initialize scheduled tasks")

    args = parser.parse_args()

    # no arguments provided, run interactive menu
    if run_task_by_flag(args):
        return

    # ========== interactive menu ==========
    print("Checking if the program is latest version ...")
    util.check_update()
    util.show_welcome_banner()
    util.init_workdir()

    start_time = time.perf_counter()
    LOGIN = util.check_login()

    while True:
        util.show_menu()
        userInput = input("Enter your choice: ").strip()

        non_login_required_actions = {"7", "0"}
        if userInput not in non_login_required_actions and not LOGIN:
            while not LOGIN:
                login.sso_login()
                LOGIN = util.check_login()

        if userInput == "0":
            print("Exiting program...")
            return
        elif userInput == "1":
            ap.main()
        elif userInput == "2":
            ar.main()
        elif userInput == "3":
            project_status.main()
        elif userInput == "4":
            bridge_report.main()
        elif userInput == "5":
            shipping_monitor.main()
        elif userInput == "6":
            daily_received.main()
        elif userInput == "7":
            option = input("Merge Amazon invoices? (Y/N): ").strip().upper()
            if option.startswith("Y"):
                util.merge_amazon_invoices()
            else:
                util.merge_pdfs()
        else:
            print("Invalid option. Please try again.")

        end_time = time.perf_counter()
        total_seconds = end_time - start_time
        minutes, seconds = divmod(int(total_seconds), 60)
        print()
        print(f"Total execution time {minutes} minutes, {seconds} seconds")
        print("\n" + "-" * 55 + "\n")
        
if __name__=="__main__":
    with RuntimeLogger(log_file) as logger:
        main()

    print("🎉Done, Have a good day.")
