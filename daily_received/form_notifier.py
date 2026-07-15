import json
import requests
from datetime import datetime
from zoneinfo import ZoneInfo
from typing import Optional

"""
    Send task execution status to a form-based HTTP endpoint.

    :param endpoint_url: Target API endpoint URL
    :param task_name: Name of the scheduled task
    :param executor: Person or system executing the task
    :param status: Execution status (e.g. Success, Failed)
    :param timestamp: Task timestamp
    :param end_date: Task end date (YYYY-MM-DD)
    :param message: Optional message
    :return: True if request succeeds, False otherwise
"""

def send_task_execution_status(
    task_name: str,
    executor: str,
    status: str,
    message: Optional[str] = ""
) -> bool:
    
    endpoint_url = "https://qcasystems.webhook.office.com/webhookb2/28fd1276-81aa-4ff8-b185-7a52aa09ae26@4ddfc3d2-92b6-4329-87f8-20567757aab2/IncomingWebhook/6005140319ae4f1aba7f7662eaf169d7/c5ce2c95-5bb8-4678-81cf-fc8200e47db7/V2hpWa31WvgmTP3b0lRKsqOxISDq502das5tK_qGh9e6c1"

    timestamp = datetime.now(ZoneInfo("America/Vancouver")).strftime("%Y-%m-%d %H:%M:%S")

    payload = {
        "title": f"{task_name} Execution Status",
        "text": (
            f"**Executor:** {executor}\n\n"
            f"**Status:** {status}\n\n"
            f"**Message:** {message}\n\n"
            f"**Time:** {timestamp}"
        )
    }

    headers = {
        "Content-Type": "application/json"
    }

    try:
        response = requests.post(
            endpoint_url,
            headers=headers,
            json=payload,
            timeout=10
        )

        response.raise_for_status()
        return True

    except requests.RequestException as exc:
        print(f"Failed to send task execution status: {exc}")
        return False
