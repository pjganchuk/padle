#!/usr/bin/env python3
"""
PADLE - Panopto Audio Descriptions List Editor

A tool for creating audio descriptions for recorded lecture videos.
Designed for export to Panopto's WebVTT audio description format.

Workflow:
1. Load a video file
2. Navigate to moments that need description
3. Generate AI descriptions for each moment
4. Review and edit descriptions
5. Export to WebVTT format

Updated for local model inference - no server required.
"""

import sys

# Windows-specific setup - MUST be done before importing tkinter
if sys.platform == 'win32':
    import ctypes
    
    # Set AppUserModelID so Windows shows our icon instead of Python's
    myappid = 'pitt.ctl.padle.1.1'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    
    # Enable DPI awareness for crisp text on high-DPI displays
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

import tkinter as tk
from app import VideoCaptionerApp


def main():
    root = tk.Tk(className="padle")
    app = VideoCaptionerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()