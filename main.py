import login
import time
import util

def main():
    print("Input your choice: ")
    #print("1 - AP Remittance")
    #print("2 - AR Noticement")
    #print("3 - export project status report")
    #print("4 - Merging Amazon Invoice PDF")
    #print("5 - Merging PDF")
    #print("6 - month-end revenue audit (careful use not finished yet)")
    print("99 - Test New Approach to update cookie")
    print()
    print("0 - Exit the program")
    userInput = input()
    
    #if (userInput == "1"): ap_main()
    #if (userInput == "2"): ar_main()
    #if (userInput == "3"): util.merge_amazon_invoices()
    #if (userInput == "4"): util.merge_pdfs()
    #if (userInput == "4"): projectStatus_main()
    #if (userInput == "5"): labour_process()
    if (userInput == "99"): login.Sso_login()

    if (userInput == "0"): return
    
    #projectStatus_main()
    #amazon_print()
    

if __name__=="__main__":
    util.show_welcome_banner()
    start_time = time.perf_counter()
    main()
    end_time = time.perf_counter()
    total_seconds = end_time - start_time
    minutes, seconds = divmod(int(total_seconds), 60)
    print(f"Total execution time {minutes} minutes, {seconds} seconds")
    print()
    input("🎉Done, Have a good day. Press Enter to exit the script.")
