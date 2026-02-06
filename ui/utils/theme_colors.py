"""
Centralized theme color utility for consistent UI styling across light and dark modes.

This module provides a unified color scheme that adapts to the current theme,
ensuring all components have consistent backgrounds, text colors, and borders.
"""
from qfluentwidgets import isDarkTheme


class ThemeColors:
    """
    Centralized theme color definitions.
    
    Provides consistent colors for:
    - Backgrounds (primary, secondary, input fields, cards)
    - Text (primary, secondary, muted)
    - Borders
    - Special elements (tutorial boxes, log viewers, etc.)
    """
    
    @staticmethod
    def background_primary() -> str:
        """Primary background color (main window background)."""
        return "#ffffff" if not isDarkTheme() else "#1e1e1e"
    
    @staticmethod
    def background_secondary() -> str:
        """Secondary background color (cards, panels)."""
        return "#ffffff" if not isDarkTheme() else "#252525"
    
    @staticmethod
    def background_input() -> str:
        """Input field background color."""
        return "#ffffff" if not isDarkTheme() else "#202020"
    
    @staticmethod
    def background_tutorial() -> str:
        """Tutorial/info box background color."""
        return "#f5f5f5" if not isDarkTheme() else "#2d2d2d"
    
    @staticmethod
    def background_log() -> str:
        """Log viewer background color."""
        return "#f5f5f5" if not isDarkTheme() else "#1e1e1e"
    
    @staticmethod
    def background_menu() -> str:
        """Menu background color."""
        return "#ffffff" if not isDarkTheme() else "#2d2d2d"
    
    @staticmethod
    def background_menu_hover() -> str:
        """Menu item hover background color."""
        return "#e5e5e5" if not isDarkTheme() else "#3c3c3c"
    
    @staticmethod
    def text_primary() -> str:
        """Primary text color."""
        return "#000000" if not isDarkTheme() else "#f0f0f0"
    
    @staticmethod
    def text_secondary() -> str:
        """Secondary text color (descriptions, labels)."""
        return "#666666" if not isDarkTheme() else "#b0b0b0"
    
    @staticmethod
    def text_muted() -> str:
        """Muted text color (subtle text, placeholders)."""
        return "#808080" if not isDarkTheme() else "#808080"
    
    @staticmethod
    def text_log() -> str:
        """Log viewer text color."""
        return "#333333" if not isDarkTheme() else "#d4d4d4"
    
    @staticmethod
    def border_primary() -> str:
        """Primary border color."""
        return "#d0d0d0" if not isDarkTheme() else "#3c3c3c"
    
    @staticmethod
    def border_secondary() -> str:
        """Secondary border color (subtle borders)."""
        return "#e0e0e0" if not isDarkTheme() else "#2d2d2d"
    
    @staticmethod
    def log_color_info() -> str:
        """Log info level color."""
        return "#333333" if not isDarkTheme() else "#d4d4d4"
    
    @staticmethod
    def log_color_success() -> str:
        """Log success level color."""
        return "#2d8659" if not isDarkTheme() else "#4ec9b0"
    
    @staticmethod
    def log_color_warning() -> str:
        """Log warning level color."""
        return "#b8860b" if not isDarkTheme() else "#dcdcaa"
    
    @staticmethod
    def log_color_error() -> str:
        """Log error level color."""
        return "#cc4125" if not isDarkTheme() else "#f48771"
    
    @staticmethod
    def status_success() -> str:
        """Success status color (green)."""
        return "#4caf50"
    
    @staticmethod
    def status_error() -> str:
        """Error status color (red)."""
        return "#f44336"
    
    @staticmethod
    def get_input_style() -> str:
        """
        Get standard input field style sheet.
        
        Returns:
            Style sheet string for input fields (LineEdit, QTextEdit, QDateEdit, etc.)
        """
        bg = ThemeColors.background_input()
        text = ThemeColors.text_primary()
        border = ThemeColors.border_primary()
        return f"""
            padding: 6px;
            border: 1px solid {border};
            border-radius: 4px;
            background-color: {bg};
            color: {text};
        """
    
    @staticmethod
    def get_tutorial_box_style() -> str:
        """
        Get tutorial/info box style sheet.
        
        Returns:
            Style sheet string for tutorial boxes
        """
        bg = ThemeColors.background_tutorial()
        text = ThemeColors.text_secondary()
        return f"color: {text}; font-size: 12px; padding: 8px; background-color: {bg}; border-radius: 4px;"
    
    @staticmethod
    def get_menu_style() -> str:
        """
        Get menu style sheet.
        
        Returns:
            Style sheet string for context menus
        """
        bg = ThemeColors.background_menu()
        text = ThemeColors.text_primary()
        border = ThemeColors.border_primary()
        hover_bg = ThemeColors.background_menu_hover()
        return f"""
            QMenu {{
                background-color: {bg};
                color: {text};
                border: 1px solid {border};
            }}
            QMenu::item:selected {{
                background-color: {hover_bg};
            }}
        """
    
    @staticmethod
    def get_log_viewer_style() -> str:
        """
        Get log viewer style sheet.
        
        Returns:
            Style sheet string for log viewers
        """
        bg = ThemeColors.background_log()
        text = ThemeColors.text_log()
        border = ThemeColors.border_primary()
        return f"""
            QTextEdit {{
                background-color: {bg};
                color: {text};
                font-family: 'Consolas', 'Courier New', monospace;
                border: 1px solid {border};
                border-radius: 4px;
                padding: 8px;
            }}
        """
