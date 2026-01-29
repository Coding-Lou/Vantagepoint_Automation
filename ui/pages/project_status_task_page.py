"""
Project Status task page with modern Fluent Design.

Task: Generate project status reports.
"""
from typing import Dict, Any, Tuple

from PySide6.QtWidgets import (
    QFormLayout,
    QHBoxLayout,
    QComboBox,
)
from PySide6.QtCore import Qt

from qfluentwidgets import (
    LineEdit,
    BodyLabel,
    PrimaryPushButton,
    PushButton,
    FluentIcon,
)

from ui.pages.base_task_page import BaseTaskPage
from workers.project_status_worker import ProjectStatusWorker
import tools.util as util_module
import core.project_status as ps_module


class ProjectStatusTaskPage(BaseTaskPage):
    """
    Task page for project status report generation.
    """
    
    def __init__(self, parent=None):
        """Initialize Project Status task page."""
        # Initialize period_data_map before calling super().__init__()
        # because BaseTaskPage.__init__ calls _create_config_widgets() which uses this attribute
        self.period_data_map: Dict[str, str] = {}
        
        super().__init__(
            "project_status",
            "Project Status",
            parent
        )
    
    def _get_page_description(self) -> str:
        """Get page description."""
        return "Generate comprehensive project status reports including invoices, earnings, expenses, and labor hours."
    
    def _create_config_widgets(self) -> None:
        """Create configuration widgets with modern Fluent Design."""
        # Tutorial section
        tutorial_label = BodyLabel("How to Use")
        tutorial_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(tutorial_label)
        
        tutorial_text = BodyLabel(
            "1. Select an accounting period from the dropdown menu.\n"
            "2. Enter project name(s) separated by commas (e.g., Project1, Project2, Project3).\n"
            "3. Click 'Execute' to generate the project status report."
        )
        tutorial_text.setWordWrap(True)
        tutorial_text.setStyleSheet("color: #666; font-size: 12px; padding: 8px; background-color: #f5f5f5; border-radius: 4px;")
        self.config_layout.addWidget(tutorial_text)
        
        # Spacer
        self.config_layout.addSpacing(12)
        
        # Period section
        period_label = BodyLabel("Accounting Period")
        period_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(period_label)
        
        self.period_combo = QComboBox()
        self.period_combo.setStyleSheet("""
            QComboBox {
                padding: 6px;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                background-color: white;
            }
            QComboBox::drop-down {
                border: none;
                padding-right: 8px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 4px solid #666;
                margin-right: 4px;
            }
        """)
        
        # Load periods from print_period()
        self._load_periods()
        self.config_layout.addWidget(self.period_combo)
        
        # Spacer
        self.config_layout.addSpacing(12)
        
        # Project names section
        projects_label = BodyLabel("Project Name(s)")
        projects_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(projects_label)
        
        self.projects_edit = LineEdit()
        self.projects_edit.setPlaceholderText("Project1, Project2, Project3 (comma-separated)")
        self.projects_edit.setStyleSheet("""
            LineEdit {
                padding: 6px;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
            }
        """)
        self.config_layout.addWidget(self.projects_edit)
    
    def _load_config(self) -> None:
        """Load configuration from config file (if any)."""
        # Project status doesn't have saved config, but we can load default period if available
        pass
    
    def _load_periods(self) -> None:
        """Load periods from print_period() and populate the dropdown."""
        # Ensure period_combo exists
        if not hasattr(self, 'period_combo'):
            return
            
        try:
            # Update headers to ensure fresh authentication
            ps_module.HEADERS = util_module.set_headers()
            
            # Get period data from print_period()
            period_data = ps_module.print_period()
            
            # Clear existing items and mapping
            self.period_combo.clear()
            self.period_data_map.clear()
            
            # Populate dropdown with formatted display text
            for p in period_data:
                display_text = f"{p['Period']} | From: {p['AccountPdStart'][:10]} To: {p['AccountPdEnd'][:10]}"
                # Convert Period to string to ensure type compatibility
                period_value = str(p['Period'])
                
                # Store mapping: display text -> Period value
                self.period_data_map[display_text] = period_value
                
                # Add to dropdown
                self.period_combo.addItem(display_text)
            
            # Select first item if available
            if self.period_combo.count() > 0:
                self.period_combo.setCurrentIndex(0)
                
        except Exception as e:
            # If loading fails, add an error message
            try:
                self.period_combo.clear()
                self.period_combo.addItem("Error loading periods - please check connection")
            except Exception:
                pass
            print(f"Error loading periods: {e}")
    
    def _validate_params(self) -> Tuple[bool, str]:
        """Validate task parameters."""
        # Check if period is selected
        if self.period_combo.currentIndex() < 0 or self.period_combo.currentText() == "":
            return False, "Please select an accounting period"
        
        # Check if period data is valid (not error message)
        current_text = self.period_combo.currentText()
        if current_text.startswith("Error loading periods"):
            return False, "Please ensure periods are loaded correctly"
        
        # Verify period value exists in mapping
        if current_text not in self.period_data_map:
            return False, "Invalid period selection"
        
        projects = self.projects_edit.text().strip()
        if not projects:
            return False, "Please enter at least one project name"
        
        # Validate that there's at least one non-empty project name
        project_list = [p.strip() for p in projects.split(",") if p.strip()]
        if not project_list:
            return False, "Please enter at least one valid project name"
        
        return True, ""
    
    def _get_params(self) -> Dict[str, Any]:
        """Get task parameters from UI."""
        # Get Period value from selected display text
        current_text = self.period_combo.currentText()
        period = self.period_data_map.get(current_text, "")
        
        # Ensure period is a string (convert if needed)
        if period:
            period = str(period)
        
        projects_text = self.projects_edit.text().strip()
        
        return {
            "period": period,
            "project_names": projects_text,  # Comma-separated string
        }
    
    def _create_worker(self, params: Dict[str, Any]) -> ProjectStatusWorker:
        """Create Project Status worker."""
        return ProjectStatusWorker(params)
