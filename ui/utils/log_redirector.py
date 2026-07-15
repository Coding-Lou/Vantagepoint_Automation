"""
Log redirector for capturing stdout/stderr and redirecting to Signal.
"""
import sys
from typing import Callable, Optional


class LogRedirector:
    """
    Redirects stdout/stderr to a callback function.

    This class is used to capture print() output from scripts and
    redirect them to UI through Signal mechanism.
    """

    def __init__(self, callback: Callable[[str, str], None]):
        """
        Initialize log redirector.

        Args:
            callback: Callback function that receives (level, message)
        """
        self.callback = callback
        self.buffer = ""
        # Once the callback's underlying Qt signal source is gone, stop using it.
        # sys.stdout is process-global, so a redirector installed by one worker can
        # outlive that worker (e.g. overlapping workers restore stdout out of order).
        # A later print() on any thread would then emit through a deleted signal
        # source and raise RuntimeError, corrupting unrelated callers. Guard against it.
        self._callback_alive = True

    def _emit(self, level: str, line: str) -> None:
        """Invoke the callback, degrading gracefully if its signal source is dead."""
        if not self._callback_alive:
            self._fallback_write(level, line)
            return
        try:
            self.callback(level, line)
        except RuntimeError:
            # e.g. "Signal source has been deleted" — the owning QObject was destroyed.
            self._callback_alive = False
            self._fallback_write(level, line)

    def _fallback_write(self, level: str, line: str) -> None:
        """Write to the real stdout so output is not silently lost."""
        try:
            sys.__stdout__.write(f"{line}\n")
        except Exception:
            pass

    def write(self, message: str) -> None:
        """
        Write message to buffer and emit when newline is found.

        Args:
            message: Message string to write
        """
        if not message:
            return

        self.buffer += message

        # Split by newline and emit each line
        while "\n" in self.buffer:
            line, self.buffer = self.buffer.split("\n", 1)
            if line.strip():
                level = self._detect_level(line)
                self._emit(level, line)

    def flush(self) -> None:
        """Flush remaining buffer content."""
        if self.buffer.strip():
            level = self._detect_level(self.buffer)
            self._emit(level, self.buffer)
            self.buffer = ""
    
    def _detect_level(self, message: str) -> str:
        """
        Detect log level from message content.
        
        Args:
            message: Message string
            
        Returns:
            Log level string: "INFO", "SUCCESS", "WARNING", or "ERROR"
        """
        msg_lower = message.lower()
        
        # Error keywords
        error_keywords = ["error", "失败", "❌", "exception", "traceback"]
        if any(keyword in msg_lower for keyword in error_keywords):
            return "ERROR"
        
        # Warning keywords
        warning_keywords = ["warning", "警告", "⚠️", "warn"]
        if any(keyword in msg_lower for keyword in warning_keywords):
            return "WARNING"
        
        # Success keywords
        success_keywords = ["success", "成功", "✅", "🎉", "done", "完成"]
        if any(keyword in msg_lower for keyword in success_keywords):
            return "SUCCESS"
        
        # Default to INFO
        return "INFO"
