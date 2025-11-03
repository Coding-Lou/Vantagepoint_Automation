import login
import time

def main():
    #print("Input your choice: ")
    #print("1 for AP Remittance")
    #print("2 for AR Noticement")
    #print("3 for quick amazon print")
    #print("4 for export project status report")
    #print("5 for month-end revenue audit (careful use not finished yet)")

    #userInput = input()
    
    #if (userInput == "1"): ap_main()
    #if (userInput == "2"): ar_main()
    #if (userInput == "3"): amazon_print()
    #if (userInput == "4"): projectStatus_main()
    #if (userInput == "5"): labour_process()

    #projectStatus_main()
    #amazon_print()
    login.Sso_login()
    print()
    print("🎉Done, Have a good day. ")

if __name__=="__main__":
    start_time = time.perf_counter()
    main()
    end_time = time.perf_counter()
    total_seconds = end_time - start_time
    minutes, seconds = divmod(int(total_seconds), 60)
    print(f"Total execution time {minutes} minutes, {seconds} seconds")