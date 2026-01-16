"""
Auto-update module for QCA Accounting Automation Tool.

Provides functionality to check for updates from GitHub Releases,
download new versions, and safely update the application.
"""

from core.update.version_checker import VersionChecker
from core.update.update_downloader import UpdateDownloader
from core.update.update_executor import UpdateExecutor
from core.update.update_manager import UpdateManager

__all__ = [
    "VersionChecker",
    "UpdateDownloader",
    "UpdateExecutor",
    "UpdateManager",
]
