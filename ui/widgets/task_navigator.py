"""
Task navigation widget for left sidebar.
"""
from typing import Dict, List, Optional

from PySide6.QtWidgets import QWidget
from PySide6.QtCore import Signal

from qfluentwidgets import NavigationInterface, FluentIcon


class TaskNavigator(QWidget):
    """
    Task navigation widget displayed on the left side.
    
    Provides navigation interface for switching between different tasks.
    Clicking a task item switches the content area but does not trigger execution.
    """
    
    # Signal emitted when a task is selected
    task_selected = Signal(str)  # task_id
    
    # Task definitions
    TASKS: List[Dict[str, str]] = [
        {"id": "ap", "name": "AP", "description": "Remittance", "icon": FluentIcon.DOCUMENT},
        {"id": "ar", "name": "AR", "description": "Statements", "icon": FluentIcon.DOCUMENT},
        {"id": "project_status", "name": "Project Status", "description": "Status Report", "icon": FluentIcon.DOCUMENT},
        {"id": "bridge_report", "name": "Bridge Report", "description": "Bridge Analysis", "icon": FluentIcon.DOCUMENT},
        {"id": "shipping_monitor", "name": "Shipping Monitor", "description": "Shipping Tracking", "icon": FluentIcon.DOCUMENT},
        {"id": "daily_received", "name": "Daily Receiving", "description": "Daily Receiving Notice", "icon": FluentIcon.DOCUMENT},
        {"id": "on_call", "name": "On-Call", "description": "On-Call Task", "icon": FluentIcon.DOCUMENT},
        {"id": "pdf_merge", "name": "PDF Merge", "description": "PDF Merge Tool", "icon": FluentIcon.DOCUMENT},
    ]
    
    def __init__(self, parent=None):
        """
        Initialize task navigator.
        
        Args:
            parent: Parent widget
        """
        super().__init__(parent)
        self.setFixedWidth(200)
        self.nav_interface: Optional[NavigationInterface] = None
        self._setup_ui()
    
    def _setup_ui(self) -> None:
        """Setup UI components."""
        # Create navigation interface
        self.nav_interface = NavigationInterface(self, showMenuButton=False)
        
        # Connect signal
        self.nav_interface.currentChanged.connect(self._on_nav_changed)
    
    def add_task_page(self, task_id: str, page: QWidget) -> None:
        """
        Add a task page to navigation.
        
        Args:
            task_id: Task identifier
            page: Task page widget
        """
        task_info = next((t for t in self.TASKS if t["id"] == task_id), None)
        if not task_info:
            return
        
        self.nav_interface.addSubInterface(
            page,
            routeKey=task_id,
            text=task_info["name"],
            icon=task_info.get("icon", FluentIcon.DOCUMENT)
        )
    
    def _on_nav_changed(self, route_key: str) -> None:
        """
        Handle navigation change.
        
        Args:
            route_key: Selected task route key
        """
        self.task_selected.emit(route_key)
