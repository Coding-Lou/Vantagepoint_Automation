"""
Configuration Manager for updating nested configuration values.

This module provides functions to update nested configuration values in config.json
without modifying other tools/util.py code. It can be used by UI components to
update configuration for different modules (AP, AR, etc.).

Usage:
    from tools.config_manager import update_config_value, get_config_value
    
    # Update AP email configuration
    update_config_value(["AP", "FROM"], "newemail@example.com")
    
    # Get configuration value
    from_email = get_config_value(["AP", "FROM"])
"""
import json
from pathlib import Path
import sys
from typing import List, Any, Optional


def get_runtime_dir() -> Path:
    """Get the runtime directory (where the script is located)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_config_path() -> Path:
    """
    Get the path to config.json file.
    In frozen mode (packaged exe), uses exe directory for persistence.
    
    Returns:
        Path to config.json file
    """
    base_dir = get_runtime_dir()
    
    # In frozen mode (packaged exe), config should be in exe directory
    if getattr(sys, "frozen", False):
        config_path = base_dir / "config" / "config.json"
        # Ensure config directory exists
        config_path.parent.mkdir(exist_ok=True)
        # If config doesn't exist, try to copy from bundled resource (one-time setup)
        if not config_path.exists():
            try:
                if hasattr(sys, '_MEIPASS'):
                    # PyInstaller temporary folder
                    bundled_config = Path(sys._MEIPASS) / "config" / "config.json"
                    if bundled_config.exists():
                        import shutil
                        shutil.copy2(bundled_config, config_path)
            except Exception:
                pass
    else:
        # Development mode: use standard paths
        config_path = (
            base_dir / "config.json"
            if (base_dir / "config.json").exists()
            else base_dir.parent / "config" / "config.json"
        )
    return config_path


def get_config_value(key_path: List[str]) -> Optional[Any]:
    """
    Get a configuration value from config.json using nested key path.
    
    Args:
        key_path: List of keys representing the path to the value
                 e.g., ["AP", "FROM"] for config["AP"]["FROM"]
    
    Returns:
        Configuration value or None if not found
    """
    config_path = get_config_path()
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        
        # Navigate through nested keys
        node = config
        for key in key_path:
            if isinstance(node, dict) and key in node:
                node = node[key]
            else:
                return None
        
        return node
    except Exception as e:
        print(f"⚠️ Error reading config: {e}")
        return None


def update_config_value(key_path: List[str], value: Any) -> bool:
    """
    Update a configuration value in config.json using nested key path.
    
    This function will:
    1. Read the current config.json
    2. Navigate to the nested key path
    3. Update the value
    4. Save the updated config back to file
    
    Args:
        key_path: List of keys representing the path to update
                 e.g., ["AP", "FROM"] for config["AP"]["FROM"]
        value: New value to set
    
    Returns:
        True if update was successful, False otherwise
    """
    config_path = get_config_path()
    try:
        # Read current config
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        
        # Navigate to the parent node, creating missing keys as needed
        node = config
        for key in key_path[:-1]:
            if not isinstance(node, dict):
                node = {}
            if key not in node:
                node[key] = {}
            node = node[key]
        
        # Update the final value
        final_key = key_path[-1]
        node[final_key] = value
        
        # Save updated config
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
            f.flush()
        
        return True
    except Exception as e:
        print(f"⚠️ Error updating config: {e}")
        return False


def update_multiple_config_values(updates: dict) -> bool:
    """
    Update multiple configuration values at once.
    
    Args:
        updates: Dictionary mapping key paths to values
                 e.g., {
                     ("AP", "FROM"): "email@example.com",
                     ("AP", "CC"): "cc@example.com"
                 }
                 or {
                     ["AP", "FROM"]: "email@example.com",
                     ["AP", "CC"]: "cc@example.com"
                 }
    
    Returns:
        True if all updates were successful, False otherwise
    """
    config_path = get_config_path()
    try:
        # Read current config
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        
        # Apply all updates
        for key_path, value in updates.items():
            # Convert key_path to list if needed
            if isinstance(key_path, tuple):
                key_path = list(key_path)
            elif isinstance(key_path, str):
                key_path = [key_path]
            
            # Navigate to the parent node, creating missing keys as needed
            node = config
            for key in key_path[:-1]:
                if not isinstance(node, dict):
                    node = {}
                if key not in node:
                    node[key] = {}
                node = node[key]
            
            # Update the final value
            final_key = key_path[-1]
            node[final_key] = value
        
        # Save updated config
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
            f.flush()
        
        return True
    except Exception as e:
        print(f"⚠️ Error updating multiple config values: {e}")
        return False
