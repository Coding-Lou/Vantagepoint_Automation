"""
Project Status task page with modern Fluent Design.

Task: Generate project status reports.
"""
from typing import Dict, Any, Tuple

from PySide6.QtWidgets import (
    QFormLayout,
    QHBoxLayout,
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


class ProjectStatusTaskPage(BaseTaskPage):
    """
    Task page for project status report generation.
    """
    
    def __init__(self, parent=None):
        """Initialize Project Status task page."""
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
            "1. Enter the accounting period in YYYYMM format (e.g., 202607).\n"
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
        
        self.period_edit = LineEdit()
        self.period_edit.setPlaceholderText("202607 (YYYYMM format)")
        self.period_edit.setStyleSheet("""
            LineEdit {
                padding: 6px;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
            }
        """)
        self.config_layout.addWidget(self.period_edit)
        
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
    
    def _validate_params(self) -> Tuple[bool, str]:
        """Validate task parameters."""
        period = self.period_edit.text().strip()
        if not period:
            return False, "Please enter an accounting period (YYYYMM format)"
        
        # Validate period format (should be 6 digits: YYYYMM)
        if not period.isdigit() or len(period) != 6:
            return False, "Period must be in YYYYMM format (e.g., 202607)"
        
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
        period = self.period_edit.text().strip()
        projects_text = self.projects_edit.text().strip()
        
        return {
            "period": period,
            "project_names": projects_text,  # Comma-separated string
        }
    
    def _create_worker(self, params: Dict[str, Any]) -> ProjectStatusWorker:
        """Create Project Status worker."""
        return ProjectStatusWorker(params)
