import sys
from pathlib import Path
from datetime import datetime
import threading


class RuntimeLogger:
    def __init__(self, log_file: Path, mode: str = "a"):
        self.log_file = Path(log_file)
        self.log_file.parent.mkdir(parents=True, exist_ok=True)

        self._stdout = sys.stdout
        self._stderr = sys.stderr
        self._lock = threading.Lock()

        self._file = open(self.log_file, mode, encoding="utf-8")

    def _timestamp(self) -> str:
        return datetime.now().strftime("[%Y-%m-%d %H:%M:%S] ")

    # ---- core write ----
    def write(self, message: str):
        if not message:
            return

        with self._lock:
            ts = self._timestamp()
            self._stdout.write(message)
            self._file.write(ts + message)
            self._file.flush()

    def flush(self):
        with self._lock:
            self._stdout.flush()
            self._file.flush()

    # ---- stderr support ----
    def error(self, message: str):
        with self._lock:
            ts = self._timestamp()
            self._stderr.write(message)
            self._file.write(ts + "[ERROR] " + message)
            self._file.flush()


    # ---- context manager ----
    def close(self):
        sys.stdout = self._stdout
        sys.stderr = self._stderr
        self._file.close()

    def __enter__(self):
        sys.stdout = self
        sys.stderr = self
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_val:
            self.error(str(exc_val))
        self.close()
