"""
Resource path utilities for PADLE

Handles finding resources (icons, data files) in both development
and PyInstaller-bundled environments.
"""

import os
import sys
from typing import Optional


def is_frozen() -> bool:
    """Check if running as PyInstaller frozen executable."""
    return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')


def get_resource_path(relative_path: str) -> str:
    """
    Get absolute path to a resource file.
    
    Works both in development and when bundled with PyInstaller.
    
    Args:
        relative_path: Path relative to the application root
        
    Returns:
        Absolute path to the resource
    """
    if is_frozen():
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_path, relative_path)


def get_icon_path() -> tuple:
    """
    Get paths to application icons.
    
    Returns:
        Tuple of (png_path, ico_path) - either may be None if not found
    """
    png_path = None
    ico_path = None
    
    if is_frozen():
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    png_candidate = os.path.join(base_path, 'icon.png')
    if os.path.exists(png_candidate):
        png_path = png_candidate
    
    ico_candidate = os.path.join(base_path, 'icon.ico')
    if os.path.exists(ico_candidate):
        ico_path = ico_candidate
    
    return png_path, ico_path


def get_data_dir() -> str:
    """
    Get the application data directory.
    
    This is where user data, settings, and caches should be stored.
    Creates the directory if it doesn't exist.
    
    Returns:
        Path to the application data directory
    """
    if sys.platform == 'win32':
        base = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        data_dir = os.path.join(base, 'padle')
    elif sys.platform == 'darwin':
        data_dir = os.path.expanduser('~/Library/Application Support/padle')
    else:
        data_dir = os.path.expanduser('~/.local/share/padle')
    
    os.makedirs(data_dir, exist_ok=True)
    return data_dir


def get_app_dir() -> str:
    """
    Get the application installation directory.
    
    Returns:
        Path to the directory containing the application
    """
    if is_frozen():
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def get_bundled_executable(name: str) -> Optional[str]:
    """
    Get the path to a bundled executable.
    
    When running as a PyInstaller bundle, executables may be bundled
    alongside the main application. This function finds them.
    
    Args:
        name: Name of the executable (without .exe on Windows)
        
    Returns:
        Full path to the executable, or None if not found
    """
    if sys.platform == 'win32':
        name = name + '.exe' if not name.endswith('.exe') else name
    
    if is_frozen():
        bundled_path = os.path.join(sys._MEIPASS, name)
        if os.path.isfile(bundled_path):
            return bundled_path
        
        exe_dir_path = os.path.join(os.path.dirname(sys.executable), name)
        if os.path.isfile(exe_dir_path):
            return exe_dir_path
    else:
        local_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), name)
        if os.path.isfile(local_path):
            return local_path
    
    return None