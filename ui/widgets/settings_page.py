"""
Settings page for viewing and editing config.json.

Provides a user-friendly interface for editing configuration without exposing raw JSON.
"""
import json
from pathlib import Path
from typing import Dict, Any, Optional, List
from copy import deepcopy

from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QScrollArea,
    QMessageBox,
)
from PySide6.QtCore import Qt, Signal

from qfluentwidgets import (
    CardWidget,
    TitleLabel,
    BodyLabel,
    PrimaryPushButton,
    PushButton,
    FluentIcon,
)

import tools.config_manager as config_manager
from ui.widgets.json_tree_editor import JsonTreeEditor


class SettingsPage(QWidget):
    """
    Settings page for viewing and editing configuration.
    
    Features:
    - View mode (read-only) by default
    - Edit mode with Save/Cancel buttons
    - User-friendly form-based editing (no raw JSON)
    - Data validation
    - Error handling
    """
    
    # Signal emitted when config is saved
    config_saved = Signal()
    
    def __init__(self, parent=None):
        """Initialize settings page."""
        super().__init__(parent)
        self.config_path = config_manager.get_config_path()
        self.original_config: Optional[Dict[str, Any]] = None
        self.current_config: Optional[Dict[str, Any]] = None
        self.edit_mode = False
        self.config_widgets: Dict[str, Any] = {}  # Store widget references by key path
        self._setup_ui()
        self._load_config()
    
    def _setup_ui(self) -> None:
        """Setup UI layout."""
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(12)
        
        # Header
        header_layout = QHBoxLayout()
        title = TitleLabel("Settings")
        header_layout.addWidget(title)
        header_layout.addStretch()
        
        # Edit/Save/Cancel buttons
        self.edit_btn = PushButton("Edit", self, FluentIcon.EDIT)
        self.edit_btn.setMaximumWidth(100)
        self.edit_btn.clicked.connect(self._on_edit_clicked)
        header_layout.addWidget(self.edit_btn)
        
        self.save_btn = PrimaryPushButton("Save", self, FluentIcon.SAVE)
        self.save_btn.setMaximumWidth(100)
        self.save_btn.setEnabled(False)
        self.save_btn.clicked.connect(self._on_save_clicked)
        header_layout.addWidget(self.save_btn)
        
        self.cancel_btn = PushButton("Cancel", self, FluentIcon.CANCEL)
        self.cancel_btn.setMaximumWidth(100)
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self._on_cancel_clicked)
        header_layout.addWidget(self.cancel_btn)
        
        root.addLayout(header_layout)
        
        # Description
        desc = BodyLabel("View and edit application configuration. Click 'Edit' to modify settings.")
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #808080;")
        root.addWidget(desc)
        
        # Scrollable content
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        content = QWidget()
        self.content_layout = QVBoxLayout(content)
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        self.content_layout.setSpacing(16)
        
        scroll.setWidget(content)
        root.addWidget(scroll)
    
    def _load_config(self) -> None:
        """Load configuration from config.json file."""
        try:
            if not self.config_path.exists():
                # Create default config if file doesn't exist
                self._create_default_config()
            
            with open(self.config_path, "r", encoding="utf-8") as f:
                self.original_config = json.load(f)
            
            self.current_config = deepcopy(self.original_config)
            self._render_config()
            
        except json.JSONDecodeError as e:
            QMessageBox.critical(
                self,
                "Configuration Error",
                f"Invalid JSON format in config file:\n{str(e)}\n\nPlease fix the JSON file manually."
            )
        except Exception as e:
            QMessageBox.critical(
                self,
                "Error",
                f"Failed to load configuration:\n{str(e)}"
            )
    
    def _create_default_config(self) -> None:
        """Create a default configuration file."""
        default_config = {
            "VERSION": "1.0.0",
            "TOKEN": "",
            "WWWBEARER": "",
            "COOKIES": "",
            "WORKDIR": "",
            "EXECUTOR": "",
            "WEBHOOK": "",
        }
        
        # Config file is in the main directory, no need to create directory
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(default_config, f, indent=4, ensure_ascii=False)
    
    def _render_config(self) -> None:
        """Render configuration using tree editor."""
        # Clear existing widgets
        while self.content_layout.count():
            child = self.content_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        
        self.config_widgets.clear()
        
        if not self.current_config:
            return
        
        # Create tree editor card
        card = CardWidget()
        card_layout = QVBoxLayout()
        card_layout.setContentsMargins(16, 16, 16, 16)
        card_layout.setSpacing(12)
        
        # Title
        title_label = BodyLabel("Configuration")
        title_label.setStyleSheet("font-weight: 600; font-size: 14px;")
        card_layout.addWidget(title_label)
        
        # Tree editor
        self.tree_editor = JsonTreeEditor(
            self.current_config,
            editable=self.edit_mode,
            parent=card
        )
        self.tree_editor.data_changed.connect(self._on_tree_data_changed)
        card_layout.addWidget(self.tree_editor)
        
        card.setLayout(card_layout)
        self.content_layout.addWidget(card)
    
    
    def _on_edit_clicked(self) -> None:
        """Handle Edit button click - enter edit mode."""
        self.edit_mode = True
        self.edit_btn.setEnabled(False)
        self.save_btn.setEnabled(True)
        self.cancel_btn.setEnabled(True)
        
        # Re-render with editable tree editor
        self._render_config()
    
    def _on_tree_data_changed(self) -> None:
        """Handle tree editor data change."""
        # Update current config from tree editor
        if hasattr(self, 'tree_editor'):
            self.current_config = self.tree_editor.get_data()
    
    def _on_save_clicked(self) -> None:
        """Handle Save button click - save configuration."""
        try:
            # Get data from tree editor
            if hasattr(self, 'tree_editor'):
                updated_config = self.tree_editor.get_data()
            else:
                updated_config = self.current_config
            
            # Validate configuration
            validation_result = self._validate_config(updated_config)
            if not validation_result[0]:
                QMessageBox.warning(
                    self,
                    "Validation Error",
                    f"Configuration validation failed:\n{validation_result[1]}"
                )
                return
            
            # Save to file
            try:
                with open(self.config_path, "w", encoding="utf-8") as f:
                    json.dump(updated_config, f, indent=4, ensure_ascii=False)
                    f.flush()
                
                # Update current and original config
                self.original_config = deepcopy(updated_config)
                self.current_config = deepcopy(updated_config)
                
                # Exit edit mode
                self.edit_mode = False
                self.edit_btn.setEnabled(True)
                self.save_btn.setEnabled(False)
                self.cancel_btn.setEnabled(False)
                
                # Re-render in view mode
                self._render_config()
                
                QMessageBox.information(
                    self,
                    "Success",
                    "Configuration saved successfully."
                )
                
                self.config_saved.emit()
                
            except PermissionError:
                QMessageBox.critical(
                    self,
                    "Permission Error",
                    "Cannot save configuration file. Please check file permissions or close the file if it's open in another program."
                )
            except Exception as e:
                QMessageBox.critical(
                    self,
                    "Save Error",
                    f"Failed to save configuration:\n{str(e)}"
                )
        
        except Exception as e:
            QMessageBox.critical(
                self,
                "Error",
                f"An error occurred while saving:\n{str(e)}"
            )
    
    def _on_cancel_clicked(self) -> None:
        """Handle Cancel button click - discard changes and exit edit mode."""
        # Restore original config
        self.current_config = deepcopy(self.original_config)
        
        # Exit edit mode
        self.edit_mode = False
        self.edit_btn.setEnabled(True)
        self.save_btn.setEnabled(False)
        self.cancel_btn.setEnabled(False)
        
        # Re-render in view mode
        self._render_config()
    
    
    def _validate_config(self, config: Dict[str, Any]) -> tuple[bool, str]:
        """
        Validate configuration data.
        
        Args:
            config: Configuration dictionary to validate
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        # Check for required fields (add as needed)
        required_fields = ["VERSION"]
        
        for field in required_fields:
            if field not in config:
                return False, f"Required field '{field}' is missing."
        
        # Validate data types
        if "VERSION" in config and not isinstance(config["VERSION"], str):
            return False, "VERSION must be a string."
        
        # Add more validation rules as needed
        
        return True, ""
    
    def refresh_config(self) -> None:
        """Refresh configuration from disk (called when config is updated elsewhere)."""
        self._load_config()
