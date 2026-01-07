import json
import subprocess
import sys
from pathlib import Path

# =========================
# basic path（IDE / CMD ）
# =========================

def get_runtime_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

BASE_DIR = get_runtime_dir()

TASKS_FILE = (
    BASE_DIR / "tasks.json"
    if (BASE_DIR / "tasks.json").exists()
    else BASE_DIR.parent / "config" / "tasks.json"
)

if not TASKS_FILE.exists():
    raise FileNotFoundError(f"tasks.json not found at {TASKS_FILE}")

# =========================
# Schedule command build
# =========================
def build_schedule_args(schedule: dict) -> list[str]:
    stype = schedule["type"]

    if stype == "minute":
        return ["/SC", "MINUTE", "/MO", str(schedule.get("interval", 5))]

    if stype == "daily":
        return ["/SC", "DAILY", "/ST", schedule["time"]]

    if stype == "weekly":
        days = ",".join(schedule["days"])
        return ["/SC", "WEEKLY", "/D", days, "/ST", schedule["time"]]

    raise ValueError(f"Unsupported schedule type: {stype}")


# =========================
# create task
# =========================
def create_task(task: dict):
    task_id = task["id"]
    schedule = task["schedule"]
    schedule_args = build_schedule_args(schedule)

    # exe path
    exe_path = Path(task["exe"]).resolve()
    if not exe_path.exists():
        print(f"Executable not found: {exe_path}")
        return

    # args optional
    args = task.get("args", "")
    if args:
        task_cmd = f'"{exe_path}" {args}'
    else:
        task_cmd = f'"{exe_path}"'

    # schtasks command
    cmd = [
        "schtasks",
        "/create",
        "/f",
        "/tn",
        task_id,
        "/tr",
        task_cmd,
        *schedule_args,
    ]

    print("Creating / Updating task:")
    print(" ", " ".join(cmd))
    subprocess.run(cmd, check=True)
    print(f"[OK] Task ready: {task_id}")

# =========================
# main entry
# =========================
def main():
    if not TASKS_FILE.exists():
        print(f"{TASKS_FILE} not found")
        return
    with open(TASKS_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    for task in data.get("tasks", []):
        create_task(task)

if __name__ == "__main__":
    main()