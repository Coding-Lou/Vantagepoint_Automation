import sys
import time
import os
import shutil
import subprocess

def main():
    if len(sys.argv) < 3:
        print("Invalid args")
        return
    old_exe = sys.argv[1]
    new_exe = sys.argv[2]
    print("Updater running...")
    for _ in range(30):
        try:
            os.remove(old_exe)
            break
        except PermissionError:
            time.sleep(0.5)
    else:
        print("Cannot update. Old EXE still running.")
        return

    shutil.copyfile(new_exe, old_exe)

    subprocess.Popen([old_exe])

    print("Update completed.")
    sys.exit(0)


if __name__ == "__main__":
    main()
