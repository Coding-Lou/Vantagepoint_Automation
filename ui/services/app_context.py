"""
Application context for global state management.

Provides singleton pattern for accessing login state and user information
across the entire application.
"""
from typing import Optional, Dict, Any
from PySide6.QtCore import QObject, Signal


class AppContext(QObject):
    """
    Singleton application context for global state management.
    
    Manages:
    - Login state
    - User information
    - Application-wide settings
    """
    
    # Signals
    login_state_changed = Signal(bool)  # is_logged_in
    user_info_changed = Signal(dict)  # user_info
    
    _instance: Optional['AppContext'] = None
    
    def __new__(cls):
        """Create singleton instance."""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        """Initialize application context."""
        if hasattr(self, '_initialized'):
            return
        
        super().__init__()
        self._initialized = True
        
        # Login state
        self._is_logged_in: bool = False
        self._user_info: Dict[str, Any] = {}
        self._user_email: Optional[str] = None
    
    @property
    def is_logged_in(self) -> bool:
        """Get login state."""
        return self._is_logged_in
    
    @property
    def user_email(self) -> Optional[str]:
        """Get user email."""
        return self._user_email
    
    @property
    def user_info(self) -> Dict[str, Any]:
        """Get user information dictionary."""
        return self._user_info.copy()
    
    def set_login_state(self, is_logged_in: bool, user_info: Optional[Dict[str, Any]] = None) -> None:
        """
        Update login state.
        
        Args:
            is_logged_in: Whether user is logged in
            user_info: Optional user information dictionary
        """
        self._is_logged_in = is_logged_in
        
        if user_info:
            self._user_info = user_info.copy()
            self._user_email = user_info.get("email") or user_info.get("EMail")
        elif not is_logged_in:
            # Clear user info on logout
            self._user_info = {}
            self._user_email = None
        
        # Emit signal
        self.login_state_changed.emit(is_logged_in)
        if user_info:
            self.user_info_changed.emit(self._user_info)
    
    def clear_login_state(self) -> None:
        """Clear login state and user information."""
        self.set_login_state(False, None)


# Global instance getter
def get_app_context() -> AppContext:
    """
    Get global application context instance.
    
    Returns:
        AppContext singleton instance
    """
    return AppContext()
