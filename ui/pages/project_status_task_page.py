"""
Project Status task page with modern Fluent Design.

Task: Generate project status reports.
"""
from typing import Dict, Any, Tuple

from PySide6.QtWidgets import (
    QFormLayout,
    QHBoxLayout,
    QButtonGroup,    
)
from PySide6.QtCore import Qt
from datetime import date

from qfluentwidgets import (
    ComboBox,
    RadioButton,
    LineEdit,
    BodyLabel,
)

from ui.pages.base_task_page import BaseTaskPage
from ui.utils.theme_colors import ThemeColors
from workers.project_status_worker import ProjectStatusWorker
from ui.services.app_context import get_app_context
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
            "project_status",
            parent
        )
        
        # Connect to login state changes to reload periods when user logs in
        app_context = get_app_context()
        app_context.login_state_changed.connect(self._on_login_state_changed)
    
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
            "2. Choose to filter by charge type and created date, or manually enter project names.\n"
            "3. Click 'Execute' to generate the project status report."
        )
        tutorial_text.setWordWrap(True)
        tutorial_text.setStyleSheet(ThemeColors.get_tutorial_box_style())
        self.config_layout.addWidget(tutorial_text)
        
        # Spacer
        self.config_layout.addSpacing(12)
        
        # Period section
        period_label = BodyLabel("Accounting Period")
        period_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(period_label)
        
        self.period_combo = ComboBox()
        bg = ThemeColors.background_input()
        text = ThemeColors.text_primary()
        border = ThemeColors.border_primary()
        arrow_color = ThemeColors.text_secondary()
        self.period_combo.setStyleSheet(f"""
            ComboBox {{
                padding: 6px;
                border: 1px solid {border};
                border-radius: 4px;
                background-color: {bg};
                color: {text};
            }}
            ComboBox::drop-down {{
                border: none;
                padding-right: 8px;
            }}
            ComboBox::down-arrow {{
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 4px solid {arrow_color};
                margin-right: 4px;
            }}
        """)
        
        # Load periods from print_period()
        self._load_periods()
        self.config_layout.addWidget(self.period_combo)
        
        # Spacer
        self.config_layout.addSpacing(12)
        
        # Filter mode selection section
        filter_mode_label = BodyLabel("Filter Mode")
        filter_mode_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(filter_mode_label)
        
        # Create button group for radio buttons
        self.filter_mode_group = QButtonGroup(self)
        
        # Year selection dropdown (only visible when filter mode is selected)
        self.year_label = BodyLabel("Project Created After Year")
        self.year_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(self.year_label)
        
        self.year_combo = ComboBox()
        current_year = date.today().year
        # Add last 5 years to dropdown
        for i in range(5):
            year = current_year - i
            self.year_combo.addItem(f"{str(year)}-01-01T00:00:00", userData=year)
        
        # Set default to 3 years ago (current behavior)
        default_year = current_year - 3
        for i in range(self.year_combo.count()):
            if self.year_combo.itemData(i) == default_year:
                self.year_combo.setCurrentIndex(i)
                break
        
        bg = ThemeColors.background_input()
        text = ThemeColors.text_primary()
        border = ThemeColors.border_primary()
        arrow_color = ThemeColors.text_secondary()
        self.year_combo.setStyleSheet(f"""
            ComboBox {{
                padding: 6px;
                border: 1px solid {border};
                border-radius: 4px;
                background-color: {bg};
                color: {text};
            }}
            ComboBox::drop-down {{
                border: none;
                padding-right: 8px;
            }}
            ComboBox::down-arrow {{
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 4px solid {arrow_color};
                margin-right: 4px;
            }}
        """)
        self.config_layout.addWidget(self.year_combo)
        
        # Option 1: Filter by charge type and created date
        self.filter_radio = RadioButton(
            "Filter by charge type (Regular) and created date (after selected year)"
        )
        self.filter_radio.setChecked(False)
        self.filter_radio.setStyleSheet("""
            RadioButton {
                padding: 4px;
                font-size: 12px;
            }
            RadioButton::indicator {
                width: 16px;
                height: 16px;
            }
        """)
        self.filter_mode_group.addButton(self.filter_radio, 0)
        self.config_layout.addWidget(self.filter_radio)
        
        # Option 2: Manual project input
        self.manual_radio = RadioButton("Manually enter project names")
        self.manual_radio.setChecked(True)  # Default to manual input
        self.manual_radio.setStyleSheet("""
            RadioButton {
                padding: 4px;
                font-size: 12px;
            }
            RadioButton::indicator {
                width: 16px;
                height: 16px;
            }
        """)
        self.filter_mode_group.addButton(self.manual_radio, 1)
        self.config_layout.addWidget(self.manual_radio)
        
        # Connect radio buttons to show/hide project input
        self.filter_radio.toggled.connect(self._on_filter_mode_changed)
        self.manual_radio.toggled.connect(self._on_filter_mode_changed)
        
        # Spacer
        self.config_layout.addSpacing(12)
        
        # Project names section (initially visible)
        self.projects_label = BodyLabel("Project Name(s)")
        self.projects_label.setStyleSheet("font-weight: 600; font-size: 13px;")
        self.config_layout.addWidget(self.projects_label)
        
        self.projects_edit = LineEdit()
        self.projects_edit.setPlaceholderText("Project1, Project2, Project3 (comma-separated)")
        self.projects_edit.setStyleSheet(
            f"""
            LineEdit {{
                {ThemeColors.get_input_style()}
            }}
            """
        )
        self.config_layout.addWidget(self.projects_edit)
        
        # Set initial visibility state (manual mode is default, so project input is visible, year selector is hidden)
        self._on_filter_mode_changed()
    
    def _load_config(self) -> None:
        """Load configuration from config file (if any)."""
        # Project status doesn't have saved config, but we can load default period if available
        pass
    
    def _on_filter_mode_changed(self) -> None:
        """Handle filter mode radio button change - show/hide project input and year selector."""
        # Show project input only when manual mode is selected
        is_manual_mode = self.manual_radio.isChecked()
        is_filter_mode = self.filter_radio.isChecked()
        self.projects_label.setVisible(is_manual_mode)
        self.projects_edit.setVisible(is_manual_mode)
        # Show year selector only when filter mode is selected
        if hasattr(self, 'year_combo') and hasattr(self, 'year_label'):
            self.year_label.setVisible(is_filter_mode)
            self.year_combo.setVisible(is_filter_mode)
    
    def _on_login_state_changed(self, is_logged_in: bool) -> None:
        """Handle login state change - reload periods when user logs in."""
        if is_logged_in:
            # User just logged in, reload periods
            self._load_periods()
    
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
                # Convert Period to string to ensure type compatibility with change_period() function
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
        
        # If manual mode is selected, validate project names
        if self.manual_radio.isChecked():
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
        
        # Ensure period is a string (convert if needed for type safety)
        if period:
            period = str(period)
        
        # Determine filter mode
        use_filter = self.filter_radio.isChecked()
        projects_text = self.projects_edit.text().strip() if self.manual_radio.isChecked() else ""
        
        # Get selected year from dropdown (only used when filter mode is enabled)
        selected_year = None
        if use_filter and hasattr(self, 'year_combo'):
            selected_year = self.year_combo.currentData()
            if selected_year is None:
                # Fallback to current text if data is not available
                try:
                    selected_year = int(self.year_combo.currentText())
                except ValueError:
                    selected_year = date.today().year - 3  # Default fallback
        
        return {
            "period": period,
            "use_filter": use_filter,  # True if filtering by charge type and created date
            "project_names": projects_text,  # Comma-separated string (empty if use_filter is True)
            "start_year": selected_year,  # Selected year for filter (only used when use_filter is True)
        }
    
    def _create_worker(self, params: Dict[str, Any]) -> ProjectStatusWorker:
        """Create Project Status worker."""
        return ProjectStatusWorker(params)
