"""
JSON Tree Editor Widget - Tree-based JSON editor similar to jsoneditoronline.org.

Provides a user-friendly tree view for editing JSON configuration with:
- Expandable/collapsible nodes
- Inline editing
- Add/delete nodes
- HTML format preservation
"""
import json
import re
from typing import Dict, Any, Optional, List
from copy import deepcopy

from PySide6.QtWidgets import (
    QTreeWidget,
    QTreeWidgetItem,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLineEdit,
    QTextEdit,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QMessageBox,
    QMenu,
    QInputDialog,
    QAbstractItemView,
)
from PySide6.QtCore import Qt, Signal, QTimer, QObject
from PySide6.QtGui import QAction

from qfluentwidgets import (
    PushButton,
    FluentIcon,
    LineEdit,
    PlainTextEdit,
    CardWidget,
    isDarkTheme,
)

from ui.utils.theme_colors import ThemeColors


class JsonTreeEditor(QWidget):
    """
    Tree-based JSON editor widget.
    
    Features:
    - Tree view of JSON structure
    - Inline editing for values
    - Add/delete nodes
    - HTML format preservation
    - Expand/collapse nodes
    """
    
    # Signal emitted when data changes
    data_changed = Signal()
    
    def __init__(self, data: Dict[str, Any], editable: bool = True, parent=None):
        """
        Initialize JSON tree editor.
        
        Args:
            data: JSON data to display
            editable: Whether the editor is editable
            parent: Parent widget
        """
        super().__init__(parent)
        self.original_data = deepcopy(data)
        self.current_data = deepcopy(data)
        self.editable = editable
        self._setup_ui()
        self._populate_tree()
    
    def _setup_ui(self) -> None:
        """Setup UI layout."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        # Tree widget
        self.tree = QTreeWidget(self)
        self.tree.setHeaderLabels(["Key", "Value", "Type"])
        self.tree.setColumnWidth(0, 250)
        self.tree.setColumnWidth(1, 400)
        self.tree.setColumnWidth(2, 100)
        self.tree.setAlternatingRowColors(True)
        self.tree.setRootIsDecorated(True)
        self.tree.setItemsExpandable(True)
        self.tree.setExpandsOnDoubleClick(False)
        
        # Enable context menu if editable
        if self.editable:
            self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            self.tree.customContextMenuRequested.connect(self._show_context_menu)
            self.tree.itemDoubleClicked.connect(self._on_item_double_clicked)
        
        # Close inline editor when clicking elsewhere
        self.tree.itemClicked.connect(self._on_item_clicked)
        # Also close when clicking empty area
        self.tree.mousePressEvent = self._tree_mouse_press_event
        
        # Track inline editor
        self.current_editor = None
        self.current_editor_item = None
        
        layout.addWidget(self.tree)
        
        # Toolbar for add/delete actions
        if self.editable:
            toolbar = QHBoxLayout()
            toolbar.addStretch()
            
            add_btn = PushButton("Add Item", self, FluentIcon.ADD)
            add_btn.clicked.connect(self._add_root_item)
            toolbar.addWidget(add_btn)
            
            expand_btn = PushButton("Expand All", self)
            expand_btn.clicked.connect(self.tree.expandAll)
            toolbar.addWidget(expand_btn)
            
            collapse_btn = PushButton("Collapse All", self)
            collapse_btn.clicked.connect(self.tree.collapseAll)
            toolbar.addWidget(collapse_btn)
            
            layout.addLayout(toolbar)
    
    def _populate_tree(self) -> None:
        """Populate tree with JSON data."""
        self.tree.clear()
        
        if not self.current_data:
            return
        
        # Add root items
        for key, value in sorted(self.current_data.items()):
            item = self._create_tree_item(key, value, None)
            self.tree.addTopLevelItem(item)
            item.setExpanded(True)
    
    def _create_tree_item(
        self,
        key: str,
        value: Any,
        parent: Optional[QTreeWidgetItem]
    ) -> QTreeWidgetItem:
        """
        Create a tree item for a key-value pair.
        
        Args:
            key: Key name
            value: Value (can be dict, list, str, etc.)
            parent: Parent tree item (None for root items)
            
        Returns:
            Created tree item
        """
        item = QTreeWidgetItem(parent)
        item.setText(0, str(key))
        item.setData(0, Qt.ItemDataRole.UserRole, key)
        item.setData(1, Qt.ItemDataRole.UserRole, value)
        
        # Set value and type columns
        if isinstance(value, dict):
            item.setText(1, f"Object ({len(value)} items)")
            item.setText(2, "object")
            item.setExpanded(True)
            # Add child items
            for sub_key, sub_value in sorted(value.items()):
                self._create_tree_item(sub_key, sub_value, item)
        elif isinstance(value, list):
            item.setText(1, f"Array ({len(value)} items)")
            item.setText(2, "array")
            item.setExpanded(True)
            # Add child items for array elements
            for idx, array_value in enumerate(value):
                array_key = f"[{idx}]"
                self._create_tree_item(array_key, array_value, item)
        elif isinstance(value, str):
            # Truncate long strings for display
            display_value = value[:100] + "..." if len(value) > 100 else value
            item.setText(1, display_value)
            item.setText(2, "string")
        elif isinstance(value, bool):
            item.setText(1, str(value))
            item.setText(2, "boolean")
        elif isinstance(value, (int, float)):
            item.setText(1, str(value))
            item.setText(2, "number")
        elif value is None:
            item.setText(1, "null")
            item.setText(2, "null")
        else:
            item.setText(1, str(value))
            item.setText(2, "unknown")
        
        return item
    
    def _show_context_menu(self, position) -> None:
        """Show context menu for tree item."""
        item = self.tree.itemAt(position)
        if not item:
            return
        
        menu = QMenu(self)
        
        # Apply theme-aware menu style
        menu.setStyleSheet(ThemeColors.get_menu_style())
        
        # Edit action
        edit_action = QAction("Edit", self)
        # Use inline editing for simple values, dialog for complex
        value = item.data(1, Qt.ItemDataRole.UserRole)
        if isinstance(value, (str, int, float, bool)) and not isinstance(value, (dict, list)):
            is_html = isinstance(value, str) and self._is_html_content(value)
            is_long_string = isinstance(value, str) and (len(value) > 100 or "\n" in value)
            if is_html or is_long_string:
                edit_action.triggered.connect(lambda: self._edit_item_dialog(item))
            else:
                edit_action.triggered.connect(lambda: self._edit_item_inline(item))
        else:
            edit_action.triggered.connect(lambda: self._edit_item_dialog(item))
        menu.addAction(edit_action)
        
        # Add child action (for objects and arrays)
        value_type = item.data(1, Qt.ItemDataRole.UserRole)
        if isinstance(value_type, (dict, list)):
            add_action = QAction("Add Child", self)
            add_action.triggered.connect(lambda: self._add_child_item(item))
            menu.addAction(add_action)
        
        # Delete action
        delete_action = QAction("Delete", self)
        delete_action.triggered.connect(lambda: self._delete_item(item))
        menu.addAction(delete_action)
        
        menu.exec(self.tree.mapToGlobal(position))
    
    def _tree_mouse_press_event(self, event) -> None:
        """Handle mouse press on tree - close inline editor if clicking empty area."""
        item = self.tree.itemAt(event.pos())
        if not item and self.current_editor:
            self._close_inline_editor()
        # Call original mouse press event
        QTreeWidget.mousePressEvent(self.tree, event)
    
    def _on_item_clicked(self, item: QTreeWidgetItem, column: int) -> None:
        """Handle item click - close inline editor if clicking different item."""
        if self.current_editor and self.current_editor_item != item:
            self._close_inline_editor()
    
    def _on_item_double_clicked(self, item: QTreeWidgetItem, column: int) -> None:
        """Handle double-click on tree item - use inline editing for simple values."""
        if column == 1:  # Value column
            value = item.data(1, Qt.ItemDataRole.UserRole)
            
            # Use inline editing for simple values
            if isinstance(value, (str, int, float, bool)) and not isinstance(value, (dict, list)):
                # Check if it's HTML or long string - use dialog for these
                is_html = isinstance(value, str) and self._is_html_content(value)
                is_long_string = isinstance(value, str) and (len(value) > 100 or "\n" in value)
                
                if is_html or is_long_string:
                    # Use dialog for HTML and long strings
                    self._edit_item_dialog(item)
                else:
                    # Use inline editing for simple values
                    self._edit_item_inline(item)
            elif isinstance(value, (dict, list)):
                # Use dialog for complex types
                self._edit_item_dialog(item)
            else:
                # Use inline editing for other simple types
                self._edit_item_inline(item)
    
    def _edit_item_inline(self, item: QTreeWidgetItem) -> None:
        """Edit a tree item's value using inline editor."""
        value = item.data(1, Qt.ItemDataRole.UserRole)
        
        # Close any existing inline editor
        self._close_inline_editor()
        
        # Get the rectangle for the item
        rect = self.tree.visualItemRect(item)
        
        # Calculate position for value column (column 1)
        column_0_width = self.tree.columnWidth(0)
        column_1_width = self.tree.columnWidth(1)
        
        # Create inline editor
        editor = LineEdit(self.tree)
        editor.setText(str(value))
        editor.setFrame(True)
        
        # Position editor over the value column
        editor.setGeometry(
            column_0_width + 2,  # Start after column 0 with small padding
            rect.y(),
            column_1_width - 4,  # Width minus padding
            rect.height()
        )
        
        # Select all text for easy editing
        editor.selectAll()
        editor.setFocus()
        
        # Store reference
        self.current_editor = editor
        self.current_editor_item = item
        
        # Handle editing completion
        def finish_editing():
            new_text = editor.text()
            original_value = item.data(1, Qt.ItemDataRole.UserRole)
            
            # Convert to original type
            if isinstance(original_value, bool):
                new_value = new_text.lower() in ("true", "1", "yes")
            elif isinstance(original_value, int):
                try:
                    new_value = int(new_text)
                except ValueError:
                    new_value = new_text
            elif isinstance(original_value, float):
                try:
                    new_value = float(new_text)
                except ValueError:
                    new_value = new_text
            else:
                new_value = new_text
            
            # Update item
            item.setData(1, Qt.ItemDataRole.UserRole, new_value)
            self._update_item_display(item)
            # Update parent nodes recursively to reflect changes
            self._update_parent_nodes(item)
            self._update_data_from_tree()
            self.data_changed.emit()
            
            # Close editor
            self._close_inline_editor()
        
        def cancel_editing():
            self._close_inline_editor()
        
        # Connect signals
        editor.editingFinished.connect(finish_editing)
        
        # Handle escape key by installing event filter
        class EscapeFilter(QObject):
            def __init__(self, cancel_func):
                super().__init__()
                self.cancel_func = cancel_func
            
            def eventFilter(self, obj, event):
                if event.type() == event.Type.KeyPress and event.key() == Qt.Key.Key_Escape:
                    self.cancel_func()
                    return True
                return False
        
        escape_filter = EscapeFilter(cancel_editing)
        editor.installEventFilter(escape_filter)
        
        # Show editor
        editor.show()
        editor.setFocus()
    
    def _close_inline_editor(self) -> None:
        """Close the current inline editor if any."""
        if self.current_editor:
            self.current_editor.deleteLater()
            self.current_editor = None
            self.current_editor_item = None
    
    def _edit_item_dialog(self, item: QTreeWidgetItem) -> None:
        """Edit a tree item's value using dialog (for complex types)."""
        value = item.data(1, Qt.ItemDataRole.UserRole)
        value_type = type(value).__name__
        
        # Check if it's HTML content
        is_html = isinstance(value, str) and self._is_html_content(value)
        
        # Create edit dialog
        dialog = QDialog(self)
        dialog.setWindowTitle("Edit Value")
        dialog.setMinimumWidth(500)
        
        layout = QVBoxLayout(dialog)
        
        # Key label (read-only)
        key_label = QLabel(f"Key: {item.text(0)}")
        key_label.setStyleSheet("font-weight: 600;")
        layout.addWidget(key_label)
        
        # Type label
        type_label = QLabel(f"Type: {value_type}")
        layout.addWidget(type_label)
        
        # Value editor
        if is_html:
            # Use QTextEdit for HTML
            editor = QTextEdit(dialog)
            editor.setAcceptRichText(True)
            editor.setHtml(value)
            editor.setMinimumHeight(200)
        elif isinstance(value, str) and (len(value) > 100 or "\n" in value):
            # Use QTextEdit for long strings
            editor = QTextEdit(dialog)
            editor.setPlainText(value)
            editor.setMinimumHeight(200)
        elif isinstance(value, (dict, list)):
            # Use PlainTextEdit for JSON objects/arrays
            editor = PlainTextEdit(dialog)
            editor.setPlainText(json.dumps(value, indent=2, ensure_ascii=False))
            editor.setMinimumHeight(200)
            font = editor.font()
            font.setFamily("Consolas, 'Courier New', monospace")
            editor.setFont(font)
        else:
            # Use LineEdit for simple values
            editor = LineEdit(dialog)
            editor.setText(str(value))
        
        layout.addWidget(editor)
        
        # Buttons
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        if dialog.exec() == QDialog.DialogCode.Accepted:
            # Get new value
            if isinstance(editor, QTextEdit):
                if is_html:
                    new_value = editor.toHtml()
                    # Extract body content if full HTML document
                    body_match = re.search(r'<body[^>]*>(.*?)</body>', new_value, re.DOTALL | re.IGNORECASE)
                    if body_match:
                        new_value = body_match.group(1).strip()
                else:
                    new_value = editor.toPlainText()
            elif isinstance(editor, PlainTextEdit):
                # Parse JSON
                try:
                    new_value = json.loads(editor.toPlainText())
                except json.JSONDecodeError as e:
                    QMessageBox.warning(
                        self,
                        "Invalid JSON",
                        f"Invalid JSON format:\n{str(e)}"
                    )
                    return
            else:
                new_value = editor.text()
                # Try to convert to original type
                if isinstance(value, bool):
                    new_value = new_value.lower() in ("true", "1", "yes")
                elif isinstance(value, int):
                    try:
                        new_value = int(new_value)
                    except ValueError:
                        pass
                elif isinstance(value, float):
                    try:
                        new_value = float(new_value)
                    except ValueError:
                        pass
            
            # Update item
            item.setData(1, Qt.ItemDataRole.UserRole, new_value)
            self._update_item_display(item)
            # Update parent nodes recursively to reflect changes
            self._update_parent_nodes(item)
            self._update_data_from_tree()
            self.data_changed.emit()
    
    def _add_child_item(self, item: QTreeWidgetItem) -> None:
        """Add a child item to an object or array."""
        parent_value = item.data(1, Qt.ItemDataRole.UserRole)
        
        if isinstance(parent_value, dict):
            # Add to object
            key, ok = QInputDialog.getText(self, "Add Key", "Enter key name:")
            if ok and key:
                # Add new key with empty string value
                new_item = self._create_tree_item(key, "", item)
                item.setExpanded(True)
                # Update parent node to include new child
                self._update_parent_nodes(new_item)
                self._update_data_from_tree()
                self.data_changed.emit()
        elif isinstance(parent_value, list):
            # Add to array
            value, ok = QInputDialog.getText(self, "Add Array Item", "Enter value:")
            if ok:
                try:
                    # Try to parse as JSON
                    parsed_value = json.loads(value)
                except json.JSONDecodeError:
                    # Use as string
                    parsed_value = value
                
                array_key = f"[{len(parent_value)}]"
                new_item = self._create_tree_item(array_key, parsed_value, item)
                item.setExpanded(True)
                # Update parent node to include new child
                self._update_parent_nodes(new_item)
                self._update_data_from_tree()
                self.data_changed.emit()
    
    def _add_root_item(self) -> None:
        """Add a root-level item."""
        key, ok = QInputDialog.getText(self, "Add Key", "Enter key name:")
        if ok and key:
            # Add new key with empty string value
            new_item = self._create_tree_item(key, "", None)
            self.tree.addTopLevelItem(new_item)
            new_item.setExpanded(True)
            self._update_data_from_tree()
            self.data_changed.emit()
    
    def _delete_item(self, item: QTreeWidgetItem) -> None:
        """Delete a tree item."""
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete '{item.text(0)}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            parent = item.parent()
            if parent:
                parent.removeChild(item)
                # Update parent node after removing child
                self._update_parent_nodes(parent)
            else:
                index = self.tree.indexOfTopLevelItem(item)
                self.tree.takeTopLevelItem(index)
            
            self._update_data_from_tree()
            self.data_changed.emit()
    
    def _update_item_display(self, item: QTreeWidgetItem) -> None:
        """Update the display text of a tree item."""
        value = item.data(1, Qt.ItemDataRole.UserRole)
        
        if isinstance(value, dict):
            item.setText(1, f"Object ({len(value)} items)")
            item.setText(2, "object")
            # Update children
            item.takeChildren()
            for sub_key, sub_value in sorted(value.items()):
                self._create_tree_item(sub_key, sub_value, item)
        elif isinstance(value, list):
            item.setText(1, f"Array ({len(value)} items)")
            item.setText(2, "array")
            # Update children
            item.takeChildren()
            for idx, array_value in enumerate(value):
                array_key = f"[{idx}]"
                self._create_tree_item(array_key, array_value, item)
        elif isinstance(value, str):
            display_value = value[:100] + "..." if len(value) > 100 else value
            item.setText(1, display_value)
            item.setText(2, "string")
        elif isinstance(value, bool):
            item.setText(1, str(value))
            item.setText(2, "boolean")
        elif isinstance(value, (int, float)):
            item.setText(1, str(value))
            item.setText(2, "number")
        elif value is None:
            item.setText(1, "null")
            item.setText(2, "null")
        else:
            item.setText(1, str(value))
            item.setText(2, "unknown")
    
    def _update_parent_nodes(self, item: QTreeWidgetItem) -> None:
        """
        Update parent nodes recursively when a child value changes.
        
        This ensures that when editing nested values, all parent containers
        (dict/list) are updated to reflect the changes.
        """
        parent = item.parent()
        if parent:
            # Rebuild parent's value from its children
            parent_value = parent.data(1, Qt.ItemDataRole.UserRole)
            if isinstance(parent_value, dict):
                # Rebuild dict from children
                new_dict = {}
                for i in range(parent.childCount()):
                    child = parent.child(i)
                    child_key = child.data(0, Qt.ItemDataRole.UserRole)
                    child_value = self._get_item_value(child)
                    new_dict[child_key] = child_value
                parent.setData(1, Qt.ItemDataRole.UserRole, new_dict)
                self._update_item_display(parent)
            elif isinstance(parent_value, list):
                # Rebuild list from children (maintain order)
                new_list = []
                for i in range(parent.childCount()):
                    child = parent.child(i)
                    child_value = self._get_item_value(child)
                    new_list.append(child_value)
                parent.setData(1, Qt.ItemDataRole.UserRole, new_list)
                self._update_item_display(parent)
            
            # Recursively update parent's parent to propagate changes up the tree
            self._update_parent_nodes(parent)
    
    def _update_data_from_tree(self) -> None:
        """Update current_data from tree structure."""
        self.current_data = {}
        
        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            key = item.data(0, Qt.ItemDataRole.UserRole)
            value = self._get_item_value(item)
            self.current_data[key] = value
    
    def _get_item_value(self, item: QTreeWidgetItem) -> Any:
        """Get the value from a tree item and its children."""
        value = item.data(1, Qt.ItemDataRole.UserRole)
        
        # If it's a container type, rebuild from children
        if isinstance(value, dict):
            result = {}
            for i in range(item.childCount()):
                child = item.child(i)
                child_key = child.data(0, Qt.ItemDataRole.UserRole)
                child_value = self._get_item_value(child)
                result[child_key] = child_value
            return result
        elif isinstance(value, list):
            result = []
            for i in range(item.childCount()):
                child = item.child(i)
                child_value = self._get_item_value(child)
                result.append(child_value)
            return result
        
        # For leaf nodes, return the stored value
        return value
    
    def _is_html_content(self, text: str) -> bool:
        """Check if a string contains HTML tags."""
        if not text or not isinstance(text, str):
            return False
        html_pattern = re.compile(r'<[a-z][\s\S]*?>', re.IGNORECASE)
        return bool(html_pattern.search(text))
    
    def get_data(self) -> Dict[str, Any]:
        """Get the current JSON data from the tree."""
        self._update_data_from_tree()
        return deepcopy(self.current_data)
    
    def set_data(self, data: Dict[str, Any]) -> None:
        """Set the JSON data to display."""
        self.current_data = deepcopy(data)
        self._populate_tree()
    
    def reset(self) -> None:
        """Reset to original data."""
        self.current_data = deepcopy(self.original_data)
        self._populate_tree()
