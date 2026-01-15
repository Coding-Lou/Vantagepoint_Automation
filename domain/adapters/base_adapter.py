"""
Base adapter for wrapping script execution.
"""
from typing import Callable, Optional, Dict, Any
from abc import ABC, abstractmethod


class BaseAdapter(ABC):
    """
    Base adapter class for wrapping script execution.
    
    Adapters encapsulate original script logic and provide
    a unified interface for UI layer to call.
    """
    
    def __init__(
        self,
        log_callback: Optional[Callable[[str, str], None]] = None,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ):
        """
        Initialize adapter.
        
        Args:
            log_callback: Callback for log messages (level, message)
            progress_callback: Callback for progress updates (current, total, message)
        """
        self.log_callback = log_callback
        self.progress_callback = progress_callback
    
    @abstractmethod
    def execute(self, **kwargs) -> Dict[str, Any]:
        """
        Execute the task.
        
        Args:
            **kwargs: Task parameters
            
        Returns:
            Result dictionary with 'success', 'message', and other fields
        """
        pass
    
    def _log(self, level: str, message: str) -> None:
        """
        Log a message through callback.
        
        Args:
            level: Log level (INFO, SUCCESS, WARNING, ERROR)
            message: Log message
        """
        if self.log_callback:
            self.log_callback(level, message)
    
    def _progress(self, current: int, total: int, message: str = "") -> None:
        """
        Report progress through callback.
        
        Args:
            current: Current progress value
            total: Total progress value
            message: Progress message
        """
        if self.progress_callback:
            self.progress_callback(current, total, message)
