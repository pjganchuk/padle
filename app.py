"""
Main application window for Video Captioner

Updated for PyInstaller bundling support.
"""

import os
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import cv2
import numpy as np
import pytesseract
from PIL import Image, ImageTk

try:
    from screeninfo import get_monitors
    HAS_SCREENINFO = True
except ImportError:
    HAS_SCREENINFO = False

from config import (
    PREVIEW_WIDTH, PREVIEW_HEIGHT, 
    DEFAULT_VOLUME, AUTOSAVE_INTERVAL,
    PIPER_VOICES_DIR
)
from vision_model import get_model, MoondreamLocal
from prompts import PROMPTS
from models import ProjectState
from audio import AudioController
from video import VideoController
from tts import PiperTTS, get_default_voice
from audio_export import export_audio_description_track
from platform_utils import play_audio_file
from resources import get_resource_path, get_icon_path, is_frozen
from download_voices import (
    ENGLISH_VOICES, download_voice, is_voice_downloaded,
    get_display_name, get_size_estimate, get_available_qualities
)

# =============================================================================
# WINDOWS DARK TITLE BAR SUPPORT
# =============================================================================

def is_windows_10_or_greater():
    """Check if running on Windows 10+"""
    if sys.platform != 'win32':
        return False
    try:
        import platform
        version = platform.version()
        # Windows 10 is version 10.0, Windows 11 is 10.0.22000+
        parts = version.split('.')
        if len(parts) >= 2:
            major = int(parts[0])
            return major >= 10
    except:
        pass
    return False

def enable_dark_title_bar(window):
    """Enable dark title bar on Windows 10/11"""
    if not is_windows_10_or_greater():
        return
    
    try:
        import ctypes
        
        # Get window handle
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        
        # DWMWA_USE_IMMERSIVE_DARK_MODE = 20 (Windows 10 build 19041+)
        # DWMWA_USE_IMMERSIVE_DARK_MODE = 19 (older builds)
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        
        # Set dark mode
        value = ctypes.c_int(1)  # 1 = dark, 0 = light
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd,
            DWMWA_USE_IMMERSIVE_DARK_MODE,
            ctypes.byref(value),
            ctypes.sizeof(value)
        )
    except Exception:
        pass  # Fail silently if not supported

# =============================================================================
# CONFIGURATION
# =============================================================================

PREFERRED_MONITOR = "primary"  # "primary", or monitor index (0, 1, 2...)

# =============================================================================
# COLOR SCHEME
# =============================================================================

# Accent color (copper/bronze)
ACCENT_COLOR = "#B87333"
ACCENT_COLOR_HOVER = "#A06328"
ACCENT_COLOR_PRESSED = "#8B5A2B"

# Dark mode colors
DARK_BG = "#252525"
DARK_BG_SECONDARY = "#383838"
DARK_BG_TERTIARY = "#434343"
DARK_BG_ELEVATED = "#4e4e4e"
DARK_FG = "#e0e0e0"
DARK_FG_SECONDARY = "#a0a0a0"
DARK_BORDER = "#444444"

# Light mode colors
LIGHT_BG = "#E0E0E0"
LIGHT_BG_SECONDARY = "#F5F5F5"
LIGHT_BG_TERTIARY = "#E8E8E8"
LIGHT_BG_ELEVATED = "#DEDEDE"
LIGHT_FG = "#000000"
LIGHT_FG_SECONDARY = "#666666"
LIGHT_BORDER = "#CCCCCC"

# Status colors
SUCCESS_COLOR = "#4CAF50"
ERROR_COLOR = "#f44336"


def load_app_icon(window, for_display=False):
    """
    Load application icon with PyInstaller support.
    
    Args:
        window: Tkinter window to set icon on
        for_display: If True, returns a PhotoImage for display in UI
        
    Returns:
        PhotoImage if for_display=True, None otherwise
    """
    png_path, ico_path = get_icon_path()
    photo_image = None
    
    try:
        if sys.platform == 'win32' and ico_path:
            # Windows: Use .ico for window icon (better taskbar support)
            try:
                window.iconbitmap(ico_path)
            except:
                pass
        
        if png_path:
            # Load PNG for Linux/Mac icon and for display
            photo_image = tk.PhotoImage(file=png_path)
            if sys.platform != 'win32':
                window.iconphoto(True, photo_image)
    except Exception:
        pass
    
    if for_display and photo_image:
        return photo_image
    return None


class ReminderDialog:
    """Simple reminder dialog matching app theme"""
    
    def __init__(self, parent, title, message):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("380x150")
        self.dialog.resizable(False, False)
        self.dialog.configure(bg=DARK_BG)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center on parent
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 380) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 150) // 3
        self.dialog.geometry(f"+{x}+{y}")
        
        # Main frame
        main_frame = tk.Frame(self.dialog, bg=DARK_BG, padx=25, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Message
        tk.Label(
            main_frame,
            text=message,
            font=("Segoe UI", 10),
            fg=DARK_FG,
            bg=DARK_BG,
            wraplength=330,
            justify=tk.CENTER
        ).pack(pady=(0, 20))
        
        # OK button
        tk.Button(
            main_frame,
            text="OK",
            font=("Segoe UI", 10),
            bg=ACCENT_COLOR,
            fg="white",
            activebackground="#9A6229",
            activeforeground="white",
            relief=tk.FLAT,
            padx=25,
            pady=5,
            cursor="hand2",
            command=self.dialog.destroy
        ).pack()
        
        self.dialog.wait_window()


class ErrorDetailsDialog:
    """Dialog to show full error details with copyable text"""
    
    def __init__(self, parent, title, error_message):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("600x400")
        self.dialog.resizable(True, True)
        self.dialog.minsize(400, 250)
        self.dialog.configure(bg=DARK_BG)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center on parent
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 600) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 400) // 3
        self.dialog.geometry(f"+{x}+{y}")
        
        # Main frame
        main_frame = tk.Frame(self.dialog, bg=DARK_BG, padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        tk.Label(
            main_frame,
            text="Error Details",
            font=("Segoe UI", 12, "bold"),
            fg=ERROR_COLOR,
            bg=DARK_BG
        ).pack(anchor=tk.W, pady=(0, 10))
        
        # Instructions
        tk.Label(
            main_frame,
            text="You can select and copy the text below:",
            font=("Segoe UI", 9),
            fg=DARK_FG_SECONDARY,
            bg=DARK_BG
        ).pack(anchor=tk.W, pady=(0, 5))
        
        # Text widget with scrollbar
        text_frame = tk.Frame(main_frame, bg=DARK_BG)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.text_widget = tk.Text(
            text_frame,
            wrap=tk.WORD,
            font=("Consolas", 10),
            bg=DARK_BG_TERTIARY,
            fg=DARK_FG,
            insertbackground=DARK_FG,
            selectbackground=ACCENT_COLOR,
            selectforeground="#E0E0E0",
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=DARK_BORDER,
            highlightcolor=ACCENT_COLOR,
            padx=10,
            pady=10,
            yscrollcommand=scrollbar.set
        )
        self.text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.text_widget.yview)
        
        # Insert error message
        self.text_widget.insert(tk.END, error_message)
        
        # Buttons frame
        btn_frame = tk.Frame(main_frame, bg=DARK_BG)
        btn_frame.pack(fill=tk.X)
        
        tk.Button(
            btn_frame,
            text="Copy to Clipboard",
            font=("Segoe UI", 10),
            bg=DARK_BG_TERTIARY,
            fg=DARK_FG,
            activebackground=DARK_BG_ELEVATED,
            activeforeground=DARK_FG,
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor="hand2",
            command=self.copy_to_clipboard
        ).pack(side=tk.LEFT)
        
        tk.Button(
            btn_frame,
            text="Close",
            font=("Segoe UI", 10),
            bg=ACCENT_COLOR,
            fg="white",
            activebackground="#9A6229",
            activeforeground="white",
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor="hand2",
            command=self.dialog.destroy
        ).pack(side=tk.RIGHT)
        
        self.dialog.wait_window()
    
    def copy_to_clipboard(self):
        """Copy error text to clipboard"""
        self.dialog.clipboard_clear()
        self.dialog.clipboard_append(self.text_widget.get(1.0, tk.END).strip())
        self.dialog.update()


def get_monitor_geometry(window_width, window_height):
    """Get x, y position to center window on preferred monitor"""
    monitor_x, monitor_y, monitor_w, monitor_h = 0, 0, 1920, 1080
    
    if HAS_SCREENINFO:
        try:
            monitors = get_monitors()
            target_monitor = None
            
            if PREFERRED_MONITOR == "primary":
                for m in monitors:
                    if m.is_primary:
                        target_monitor = m
                        break
                if not target_monitor and monitors:
                    target_monitor = monitors[0]
            elif isinstance(PREFERRED_MONITOR, int) and PREFERRED_MONITOR < len(monitors):
                target_monitor = monitors[PREFERRED_MONITOR]
            
            if target_monitor:
                monitor_x = target_monitor.x
                monitor_y = target_monitor.y
                monitor_w = target_monitor.width
                monitor_h = target_monitor.height
        except:
            pass
    
    x = monitor_x + (monitor_w - window_width) // 2
    y = monitor_y + (monitor_h - window_height) // 3
    return x, y


class ModelLoadingDialog:
    """Dialog shown while the AI model is loading"""
    
    def __init__(self, root):
        self.root = root
        self.cancelled = False
        self.load_complete = False
        self.load_error = None
        
        self.dialog = tk.Toplevel(root, class_="padle")
        self.dialog.title("PADLE - Loading")
        self.dialog.resizable(False, False)
        self.dialog.configure(bg=DARK_BG)
        self.dialog.grab_set()
        
        # Center on preferred monitor
        window_width = 400
        window_height = 350
        x, y = get_monitor_geometry(window_width, window_height)
        self.dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Load icons
        self.icon_image = None
        png_path, ico_path = get_icon_path()
        
        try:
            if sys.platform == 'win32' and ico_path:
                try:
                    self.dialog.iconbitmap(ico_path)
                except:
                    pass
            
            if png_path:
                icon = tk.PhotoImage(file=png_path)
                if sys.platform != 'win32':
                    self.dialog.iconphoto(True, icon)
                self.icon_image = icon.subsample(4, 4)
        except:
            pass
        
        self.create_widgets()
        
        # Handle window close
        self.dialog.protocol("WM_DELETE_WINDOW", self.on_cancel)
    
    def create_widgets(self):
        main_frame = tk.Frame(self.dialog, bg=DARK_BG, padx=30, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Icon (small)
        if self.icon_image:
            icon_label = tk.Label(main_frame, image=self.icon_image, bg=DARK_BG)
            icon_label.pack(pady=(5, 10))
        
        # Title
        tk.Label(
            main_frame,
            text="PADLE",
            font=("Segoe UI", 14, "bold"),
            fg=ACCENT_COLOR,
            bg=DARK_BG
        ).pack(pady=(0, 5))
        
        # Status message
        self.status_var = tk.StringVar(value="Initializing...")
        self.status_label = tk.Label(
            main_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 10),
            fg=DARK_FG,
            bg=DARK_BG
        )
        self.status_label.pack(pady=(10, 15))
        
        # Note about first run
        self.note_label = tk.Label(
            main_frame,
            text="First launch downloads ~1.7GB model",
            font=("Segoe UI", 9),
            fg=DARK_FG_SECONDARY,
            bg=DARK_BG
        )
        self.note_label.pack()
    
    def update_status(self, message: str):
        """Update the status message (thread-safe)"""
        self.dialog.after(0, lambda: self.status_var.set(message))
    
    def finish(self, error: str = None):
        """Mark loading as complete"""
        self.load_error = error
        self.load_complete = True
        self.dialog.after(0, self.dialog.destroy)
    
    def on_cancel(self):
        """Handle cancel - exit the app"""
        self.cancelled = True
        self.dialog.destroy()
    
    def wait(self):
        """Wait for the dialog to close"""
        self.dialog.wait_window()


class PromptEditorDialog:
    """Dialog for editing AI prompts"""
    
    def __init__(self, parent, prompts, on_save, colors):
        self.prompts = prompts.copy()
        self.on_save = on_save
        self.colors = colors
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Edit AI Prompts")
        self.dialog.geometry("650x400")
        self.dialog.minsize(500, 350)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.configure(bg=colors["bg"])
        
        # Center dialog on parent window
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.dialog.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.dialog.winfo_height()) // 3
        self.dialog.geometry(f"+{x}+{y}")
        
        # Main frame with padding
        main_frame = ttk.Frame(self.dialog, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Instructions
        ttk.Label(main_frame, text="Customize the prompts sent to Moondream AI:", 
                  style="Header.TLabel").pack(anchor=tk.W)
        ttk.Label(main_frame, text="Use {ocr_text} in Slide + OCR prompt to insert extracted text.",
                  style="Status.TLabel").pack(anchor=tk.W, pady=(0, 10))
        
        # Notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Create a tab for each prompt type
        self.editors = {}
        tab_names = [("general", "General"), ("slide", "Slide"), ("slide_ocr", "Slide + OCR")]
        
        for key, label in tab_names:
            frame = ttk.Frame(self.notebook, padding=10)
            self.notebook.add(frame, text=label)
            
            # Text editor with scrollbar
            editor_frame = ttk.Frame(frame)
            editor_frame.pack(fill=tk.BOTH, expand=True)
            
            scrollbar = ttk.Scrollbar(editor_frame, orient=tk.VERTICAL)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            editor = tk.Text(editor_frame, wrap=tk.WORD, font=("Segoe UI", 10),
                            bg=colors["bg3"], fg=colors["fg"], 
                            insertbackground=colors["fg"],
                            selectbackground=ACCENT_COLOR,
                            selectforeground="#E0E0E0",
                            padx=10, pady=10, height=8,
                            highlightthickness=1,
                            highlightbackground=colors["border"],
                            highlightcolor=ACCENT_COLOR,
                            yscrollcommand=scrollbar.set)
            editor.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=editor.yview)
            
            editor.insert(tk.END, self.prompts.get(key, ""))
            self.editors[key] = editor
        
        # Buttons frame with fixed height
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(btn_frame, text="Reset to Defaults", 
                   command=self.reset_defaults).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="Cancel", 
                   command=self.dialog.destroy).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(btn_frame, text="Save", style="Accent.TButton",
                   command=self.save).pack(side=tk.RIGHT)
    
    def reset_defaults(self):
        """Reset prompts to defaults"""
        from prompts import PROMPTS as DEFAULT_PROMPTS
        for key, editor in self.editors.items():
            editor.delete(1.0, tk.END)
            editor.insert(tk.END, DEFAULT_PROMPTS.get(key, ""))
    
    def save(self):
        """Save the edited prompts"""
        for key, editor in self.editors.items():
            self.prompts[key] = editor.get(1.0, tk.END).strip()
        self.on_save(self.prompts)
        self.dialog.destroy()


class VoiceSelectionDialog:
    """Dialog for selecting a TTS voice with preview capability"""
    
    def __init__(self, parent, current_voice_path=None):
        self.parent = parent
        self.result = None  # Will be voice path if OK clicked
        self.tts = PiperTTS()
        self.voices = self.tts.discover_voices()
        self.preview_process = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Select Voice")
        self.dialog.geometry("500x450")
        self.dialog.resizable(False, False)
        self.dialog.configure(bg=DARK_BG)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center on parent
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 500) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 450) // 3
        self.dialog.geometry(f"+{x}+{y}")
        
        self.current_voice_path = current_voice_path
        self.create_widgets()
        
        # Handle window close
        self.dialog.protocol("WM_DELETE_WINDOW", self.on_cancel)
        
        # Wait for dialog to close
        self.dialog.wait_window()
    
    def create_widgets(self):
        # Main frame
        main_frame = tk.Frame(self.dialog, bg=DARK_BG, padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        tk.Label(
            main_frame,
            text="Select Voice for Audio Descriptions",
            font=("Segoe UI", 12, "bold"),
            fg=DARK_FG,
            bg=DARK_BG
        ).pack(anchor=tk.W)
        
        # Voice count
        tk.Label(
            main_frame,
            text=f"{len(self.voices)} voices available in {PIPER_VOICES_DIR}",
            font=("Segoe UI", 9),
            fg=DARK_FG_SECONDARY,
            bg=DARK_BG
        ).pack(anchor=tk.W, pady=(0, 10))
        
        # Voice list frame
        list_frame = tk.Frame(main_frame, bg=DARK_BG)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Scrollbar
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Voice listbox
        self.voice_listbox = tk.Listbox(
            list_frame,
            bg=DARK_BG_TERTIARY,
            fg=DARK_FG,
            selectmode=tk.SINGLE,
            font=("Segoe UI", 10),
            selectbackground=ACCENT_COLOR,
            selectforeground="#E0E0E0",
            activestyle='none',
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=DARK_BORDER,
            highlightcolor=ACCENT_COLOR,
            yscrollcommand=scrollbar.set
        )
        self.voice_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.voice_listbox.yview)
        
        # Populate voice list
        selected_index = 0
        for i, voice in enumerate(self.voices):
            self.voice_listbox.insert(tk.END, f"  {voice.display_name}")
            if voice.path == self.current_voice_path:
                selected_index = i
        
        # Select current voice
        if self.voices:
            self.voice_listbox.selection_set(selected_index)
            self.voice_listbox.see(selected_index)
        
        # Bind selection event
        self.voice_listbox.bind('<<ListboxSelect>>', self.on_voice_select)
        
        # Voice info frame
        info_frame = tk.Frame(main_frame, bg=DARK_BG_SECONDARY, padx=10, pady=8)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.info_label = tk.Label(
            info_frame,
            text="Select a voice to see details",
            font=("Segoe UI", 9),
            fg=DARK_FG_SECONDARY,
            bg=DARK_BG_SECONDARY,
            justify=tk.LEFT
        )
        self.info_label.pack(anchor=tk.W)
        
        # Preview frame
        preview_frame = tk.Frame(main_frame, bg=DARK_BG)
        preview_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(
            preview_frame,
            text="Preview text:",
            font=("Segoe UI", 9),
            fg=DARK_FG_SECONDARY,
            bg=DARK_BG
        ).pack(side=tk.LEFT)
        
        self.preview_text = tk.Entry(
            preview_frame,
            font=("Segoe UI", 10),
            bg=DARK_BG_TERTIARY,
            fg=DARK_FG,
            insertbackground=DARK_FG,
            width=35,
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=DARK_BORDER,
            highlightcolor=ACCENT_COLOR
        )
        self.preview_text.insert(0, "A woman in a blue jacket approaches the podium.")
        self.preview_text.pack(side=tk.LEFT, padx=(10, 5), fill=tk.X, expand=True)
        
        self.preview_btn = tk.Button(
            preview_frame,
            text="Preview",
            font=("Segoe UI", 9),
            bg=DARK_BG_TERTIARY,
            fg=DARK_FG,
            activebackground=DARK_BG_ELEVATED,
            activeforeground=DARK_FG,
            relief=tk.FLAT,
            padx=10,
            pady=2,
            cursor="hand2",
            command=self.preview_voice
        )
        self.preview_btn.pack(side=tk.LEFT)
        
        # Status label
        self.status_label = tk.Label(
            main_frame,
            text="",
            font=("Segoe UI", 9),
            fg=ACCENT_COLOR,
            bg=DARK_BG
        )
        self.status_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Buttons frame
        btn_frame = tk.Frame(main_frame, bg=DARK_BG)
        btn_frame.pack(fill=tk.X)
        
        tk.Button(
            btn_frame,
            text="Cancel",
            font=("Segoe UI", 10),
            bg=DARK_BG_SECONDARY,
            fg=DARK_FG,
            activebackground="#4E4E4E",
            activeforeground=DARK_FG,
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor="hand2",
            command=self.on_cancel
        ).pack(side=tk.RIGHT, padx=(5, 0))
        
        tk.Button(
            btn_frame,
            text="Select Voice",
            font=("Segoe UI", 10, "bold"),
            bg=ACCENT_COLOR,
            fg="white",
            activebackground="#9A6229",
            activeforeground="white",
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor="hand2",
            command=self.on_ok
        ).pack(side=tk.RIGHT)
        
        # Update info for initial selection
        self.on_voice_select(None)
    
    def on_voice_select(self, event):
        """Handle voice selection change"""
        selection = self.voice_listbox.curselection()
        if not selection:
            return
        
        idx = selection[0]
        if idx < len(self.voices):
            voice = self.voices[idx]
            info_text = f"Locale: {voice.locale}  |  Quality: {voice.quality}  |  Sample rate: {voice.sample_rate} Hz"
            self.info_label.config(text=info_text)
    
    def preview_voice(self):
        """Generate and play a preview of the selected voice"""
        selection = self.voice_listbox.curselection()
        if not selection:
            self.status_label.config(text="Please select a voice first")
            return
        
        idx = selection[0]
        if idx >= len(self.voices):
            return
        
        voice = self.voices[idx]
        text = self.preview_text.get().strip()
        
        if not text:
            self.status_label.config(text="Please enter preview text")
            return
        
        self.status_label.config(text="Generating preview...")
        self.preview_btn.config(state=tk.DISABLED)
        self.dialog.update()
        
        def do_preview():
            try:
                # Generate audio
                self.tts.set_voice(voice.path)
                temp_wav = self.tts.synthesize_to_temp(text)
                
                # Play using cross-platform audio player
                if not play_audio_file(temp_wav, blocking=True):
                    self.dialog.after(0, lambda: self.status_label.config(
                        text="No audio player found"))
                
                # Clean up
                if os.path.exists(temp_wav):
                    os.unlink(temp_wav)
                
                self.dialog.after(0, lambda: self.status_label.config(text=""))
                
            except Exception as e:
                self.dialog.after(0, lambda: self.status_label.config(
                    text=f"Preview failed: {str(e)[:40]}"))
            finally:
                self.dialog.after(0, lambda: self.preview_btn.config(state=tk.NORMAL))
        
        # Run in thread to keep UI responsive
        threading.Thread(target=do_preview, daemon=True).start()
    
    def on_ok(self):
        """Handle OK button"""
        selection = self.voice_listbox.curselection()
        if selection and selection[0] < len(self.voices):
            self.result = self.voices[selection[0]].path
        self.dialog.destroy()
    
    def on_cancel(self):
        """Handle cancel"""
        self.result = None
        self.dialog.destroy()


class VoiceDownloadDialog:
    """Dialog for selecting and downloading TTS voices"""
    
    def __init__(self, parent, voices_dir=None):
        self.parent = parent
        self.voices_dir = voices_dir or PIPER_VOICES_DIR
        self.result = False  # True if any voices were downloaded
        self.cancelled = False
        
        # Track checkbox variables and selection state
        self.voice_vars = {}  # (locale, name, quality) -> BooleanVar
        self.voice_selections = {}  # (locale, name, quality) -> bool (preserved across rebuilds)
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Download TTS Voices")
        self.dialog.geometry("600x550")
        self.dialog.resizable(True, True)
        self.dialog.minsize(500, 400)
        self.dialog.configure(bg=DARK_BG)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center on parent
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 600) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 550) // 3
        self.dialog.geometry(f"+{x}+{y}")
        
        self.create_widgets()
        
        # Bind mousewheel globally for this dialog
        self._bind_mousewheel()
        
        # Handle window close
        self.dialog.protocol("WM_DELETE_WINDOW", self.on_cancel)
        
        # Wait for dialog to close
        self.dialog.wait_window()
    
    def create_widgets(self):
        # Main frame
        main_frame = tk.Frame(self.dialog, bg=DARK_BG, padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        tk.Label(
            main_frame,
            text="Download TTS Voices",
            font=("Segoe UI", 14, "bold"),
            fg=DARK_FG,
            bg=DARK_BG
        ).pack(anchor=tk.W)
        
        tk.Label(
            main_frame,
            text=f"Voices will be saved to: {self.voices_dir}",
            font=("Segoe UI", 9),
            fg=DARK_FG_SECONDARY,
            bg=DARK_BG
        ).pack(anchor=tk.W, pady=(0, 10))
        
        # Filter frame
        filter_frame = tk.Frame(main_frame, bg=DARK_BG)
        filter_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Locale filter
        tk.Label(
            filter_frame,
            text="Locale:",
            font=("Segoe UI", 10),
            fg=DARK_FG,
            bg=DARK_BG
        ).pack(side=tk.LEFT)
        
        self.locale_var = tk.StringVar(value="all")
        locales = [("All", "all"), ("US English", "en_US"), ("UK English", "en_GB")]
        
        for text, value in locales:
            rb = tk.Radiobutton(
                filter_frame,
                text=text,
                variable=self.locale_var,
                value=value,
                font=("Segoe UI", 9),
                fg=DARK_FG,
                bg=DARK_BG,
                selectcolor=DARK_BG_TERTIARY,
                activebackground=DARK_BG,
                activeforeground=DARK_FG,
                command=self.apply_filters
            )
            rb.pack(side=tk.LEFT, padx=(10, 0))
        
        # Quality filter
        tk.Label(
            filter_frame,
            text="   Quality:",
            font=("Segoe UI", 10),
            fg=DARK_FG,
            bg=DARK_BG
        ).pack(side=tk.LEFT, padx=(20, 0))
        
        self.quality_var = tk.StringVar(value="all")
        self.quality_buttons = {}  # Store references to update visibility
        
        for text, value in [("All", "all"), ("Low", "low"), ("Medium", "medium"), ("High", "high")]:
            rb = tk.Radiobutton(
                filter_frame,
                text=text,
                variable=self.quality_var,
                value=value,
                font=("Segoe UI", 9),
                fg=DARK_FG,
                bg=DARK_BG,
                selectcolor=DARK_BG_TERTIARY,
                activebackground=DARK_BG,
                activeforeground=DARK_FG,
                command=self.apply_filters
            )
            rb.pack(side=tk.LEFT, padx=(10, 0))
            self.quality_buttons[value] = rb
        
        # Voice list frame with scrollbar
        list_frame = tk.Frame(main_frame, bg=DARK_BG)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Canvas for scrollable checkboxes
        self.canvas = tk.Canvas(
            list_frame,
            bg=DARK_BG_TERTIARY,
            highlightthickness=1,
            highlightbackground=DARK_BORDER
        )
        self.scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=DARK_BG_TERTIARY)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Make canvas window expand to fill width
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Build voice list
        self.build_voice_list()
        
        # Selection summary
        self.summary_label = tk.Label(
            main_frame,
            text="",
            font=("Segoe UI", 10),
            fg=ACCENT_COLOR,
            bg=DARK_BG
        )
        self.summary_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Select all / Deselect all buttons
        select_frame = tk.Frame(main_frame, bg=DARK_BG)
        select_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Button(
            select_frame,
            text="Select All",
            font=("Segoe UI", 9),
            bg=DARK_BG_TERTIARY,
            fg=DARK_FG,
            activebackground=DARK_BG_ELEVATED,
            activeforeground=DARK_FG,
            relief=tk.FLAT,
            padx=10,
            pady=3,
            cursor="hand2",
            command=self.select_all
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        tk.Button(
            select_frame,
            text="Deselect All",
            font=("Segoe UI", 9),
            bg=DARK_BG_TERTIARY,
            fg=DARK_FG,
            activebackground=DARK_BG_ELEVATED,
            activeforeground=DARK_FG,
            relief=tk.FLAT,
            padx=10,
            pady=3,
            cursor="hand2",
            command=self.deselect_all
        ).pack(side=tk.LEFT)
        
        # Buttons frame
        btn_frame = tk.Frame(main_frame, bg=DARK_BG)
        btn_frame.pack(fill=tk.X)
        
        tk.Button(
            btn_frame,
            text="Cancel",
            font=("Segoe UI", 10),
            bg=DARK_BG_SECONDARY,
            fg=DARK_FG,
            activebackground="#4E4E4E",
            activeforeground=DARK_FG,
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor="hand2",
            command=self.on_cancel
        ).pack(side=tk.RIGHT, padx=(5, 0))
        
        self.download_btn = tk.Button(
            btn_frame,
            text="Download Selected",
            font=("Segoe UI", 10, "bold"),
            bg=ACCENT_COLOR,
            fg="white",
            activebackground="#9A6229",
            activeforeground="white",
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor="hand2",
            command=self.start_download
        )
        self.download_btn.pack(side=tk.RIGHT)
        
        # Initial state
        self.update_quality_filter_visibility()
        self.update_summary()
    
    def _on_canvas_configure(self, event):
        """Make scrollable frame expand to canvas width"""
        self.canvas.itemconfig(self.canvas_window, width=event.width)
    
    def _bind_mousewheel(self):
        """Bind mouse wheel scrolling for the dialog"""
        if sys.platform == 'win32':
            self.dialog.bind_all("<MouseWheel>", self._on_mousewheel)
        else:
            self.dialog.bind_all("<Button-4>", self._on_mousewheel)
            self.dialog.bind_all("<Button-5>", self._on_mousewheel)
    
    def _unbind_mousewheel(self):
        """Unbind mouse wheel from dialog"""
        if sys.platform == 'win32':
            self.dialog.unbind_all("<MouseWheel>")
        else:
            self.dialog.unbind_all("<Button-4>")
            self.dialog.unbind_all("<Button-5>")
    
    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling"""
        if sys.platform == 'win32':
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")
    
    def build_voice_list(self, locale_filter=None, quality_filter=None):
        """Build the list of voice checkboxes with optional filtering"""
        # Save current selections before clearing
        for key, var in self.voice_vars.items():
            self.voice_selections[key] = var.get()
        
        # Clear existing widgets
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.voice_vars.clear()
        
        # Get filters
        if locale_filter is None:
            locale_filter = self.locale_var.get()
        if quality_filter is None:
            quality_filter = self.quality_var.get()
        
        # Track current locale for headers
        current_locale = None
        
        for locale, name, qualities in ENGLISH_VOICES:
            # Skip if locale doesn't match filter
            if locale_filter != "all" and locale != locale_filter:
                continue
            
            for quality in qualities:
                # Skip if quality doesn't match filter
                if quality_filter != "all" and quality != quality_filter:
                    continue
                
                key = (locale, name, quality)
                
                # Add locale header if changed
                if locale != current_locale:
                    current_locale = locale
                    locale_name = "US English" if locale == "en_US" else "UK English"
                    
                    header = tk.Label(
                        self.scrollable_frame,
                        text=locale_name,
                        font=("Segoe UI", 10, "bold"),
                        fg=ACCENT_COLOR,
                        bg=DARK_BG_TERTIARY,
                        anchor="w"
                    )
                    header.pack(fill=tk.X, padx=10, pady=(10, 5))
                
                # Create checkbox row
                row_frame = tk.Frame(self.scrollable_frame, bg=DARK_BG_TERTIARY)
                row_frame.pack(fill=tk.X, padx=10, pady=1)
                
                # Check if already downloaded
                downloaded = is_voice_downloaded(self.voices_dir, locale, name, quality)
                
                # Checkbox variable - restore previous selection if available
                var = tk.BooleanVar(value=False)
                if key in self.voice_selections:
                    var.set(self.voice_selections[key])
                elif locale == "en_US" and name == "amy" and quality == "medium" and not downloaded:
                    # Pre-select Amy US Medium if not downloaded (first time only)
                    var.set(True)
                
                self.voice_vars[key] = var
                
                # Checkbox
                display_name = get_display_name(locale, name, quality)
                size_mb = get_size_estimate(quality)
                
                cb = tk.Checkbutton(
                    row_frame,
                    text=f"  {display_name}",
                    variable=var,
                    font=("Segoe UI", 10),
                    fg=DARK_FG if not downloaded else DARK_FG_SECONDARY,
                    bg=DARK_BG_TERTIARY,
                    selectcolor=DARK_BG_SECONDARY,
                    activebackground=DARK_BG_TERTIARY,
                    activeforeground=DARK_FG,
                    anchor="w",
                    command=self.update_summary
                )
                cb.pack(side=tk.LEFT, fill=tk.X, expand=True)
                
                # Size / status label
                if downloaded:
                    status_text = "Downloaded"
                    status_color = SUCCESS_COLOR
                    cb.config(state=tk.DISABLED)
                    var.set(False)  # Don't select already downloaded
                else:
                    status_text = f"~{size_mb} MB"
                    status_color = DARK_FG_SECONDARY
                
                tk.Label(
                    row_frame,
                    text=status_text,
                    font=("Segoe UI", 9),
                    fg=status_color,
                    bg=DARK_BG_TERTIARY,
                    width=12,
                    anchor="e"
                ).pack(side=tk.RIGHT)
    
    def update_quality_filter_visibility(self):
        """Show/hide quality filter options based on available qualities"""
        locale_filter = self.locale_var.get()
        if locale_filter == "all":
            locale_filter = None
        
        available = get_available_qualities(ENGLISH_VOICES, locale_filter)
        
        # Show/hide quality buttons
        for quality, button in self.quality_buttons.items():
            if quality == "all":
                continue
            if quality in available:
                button.pack(side=tk.LEFT, padx=(10, 0))
            else:
                button.pack_forget()
                # Reset to "all" if current selection is hidden
                if self.quality_var.get() == quality:
                    self.quality_var.set("all")
    
    def apply_filters(self):
        """Apply locale and quality filters by rebuilding voice list"""
        # Update quality filter visibility first
        self.update_quality_filter_visibility()
        
        # Rebuild the voice list with current filters
        self.build_voice_list()
        
        self.update_summary()
    
    def select_all(self):
        """Select all visible (not downloaded) voices"""
        for (locale, name, quality), var in self.voice_vars.items():
            # Check if already downloaded
            downloaded = is_voice_downloaded(self.voices_dir, locale, name, quality)
            
            if not downloaded:
                var.set(True)
        
        self.update_summary()
    
    def deselect_all(self):
        """Deselect all voices"""
        for var in self.voice_vars.values():
            var.set(False)
        self.update_summary()
    
    def update_summary(self):
        """Update the selection summary"""
        selected = []
        total_size = 0
        
        for (locale, name, quality), var in self.voice_vars.items():
            if var.get():
                selected.append((locale, name, quality))
                total_size += get_size_estimate(quality)
        
        if selected:
            self.summary_label.config(
                text=f"{len(selected)} voice(s) selected (~{total_size} MB)"
            )
            self.download_btn.config(state=tk.NORMAL)
        else:
            self.summary_label.config(text="No voices selected")
            self.download_btn.config(state=tk.DISABLED)
    
    def start_download(self):
        """Start downloading selected voices"""
        # Collect selected voices
        selected = []
        for (locale, name, quality), var in self.voice_vars.items():
            if var.get():
                selected.append((locale, name, quality))
        
        if not selected:
            return
        
        # Show progress dialog
        progress = VoiceDownloadProgressDialog(
            self.dialog,
            selected,
            self.voices_dir
        )
        
        if progress.completed_count > 0:
            self.result = True
            # Refresh the voice list to show newly downloaded
            self.build_voice_list()
            self.update_summary()
    
    def on_cancel(self):
        """Handle cancel"""
        self._unbind_mousewheel()
        self.dialog.destroy()


class VoiceDownloadProgressDialog:
    """Progress dialog for voice downloads"""
    
    def __init__(self, parent, voices_to_download, voices_dir):
        self.parent = parent
        self.voices_to_download = voices_to_download
        self.voices_dir = voices_dir
        self.cancelled = False
        self.completed_count = 0
        self.failed_count = 0
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Downloading Voices")
        self.dialog.geometry("450x200")
        self.dialog.resizable(False, False)
        self.dialog.configure(bg=DARK_BG)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center on parent
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 450) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 200) // 3
        self.dialog.geometry(f"+{x}+{y}")
        
        self.create_widgets()
        
        # Handle window close
        self.dialog.protocol("WM_DELETE_WINDOW", self.on_cancel)
        
        # Start download in background
        self.download_thread = threading.Thread(target=self.run_downloads, daemon=True)
        self.download_thread.start()
        
        # Wait for dialog to close
        self.dialog.wait_window()
    
    def create_widgets(self):
        # Main frame
        main_frame = tk.Frame(self.dialog, bg=DARK_BG, padx=25, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Current voice label
        self.voice_label = tk.Label(
            main_frame,
            text="Preparing...",
            font=("Segoe UI", 10),
            fg=DARK_FG,
            bg=DARK_BG,
            wraplength=400
        )
        self.voice_label.pack(pady=(0, 10))
        
        # Progress label
        self.progress_label = tk.Label(
            main_frame,
            text="0 / 0",
            font=("Segoe UI", 9),
            fg=DARK_FG_SECONDARY,
            bg=DARK_BG
        )
        self.progress_label.pack(pady=(0, 5))
        
        # Status label
        self.status_label = tk.Label(
            main_frame,
            text="",
            font=("Segoe UI", 9),
            fg=ACCENT_COLOR,
            bg=DARK_BG
        )
        self.status_label.pack(pady=(0, 15))
        
        # Cancel button
        self.cancel_btn = tk.Button(
            main_frame,
            text="Cancel",
            font=("Segoe UI", 10),
            bg=DARK_BG_SECONDARY,
            fg=DARK_FG,
            activebackground="#4E4E4E",
            activeforeground=DARK_FG,
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor="hand2",
            command=self.on_cancel
        )
        self.cancel_btn.pack()
    
    def run_downloads(self):
        """Run downloads in background thread"""
        total = len(self.voices_to_download)
        
        for i, (locale, name, quality) in enumerate(self.voices_to_download):
            if self.cancelled:
                break
            
            display_name = get_display_name(locale, name, quality)
            
            # Update UI
            self.dialog.after(0, lambda dn=display_name, idx=i, t=total: self._update_ui(
                f"Downloading: {dn}",
                f"{idx + 1} / {t}",
                ""
            ))
            
            # Download the voice
            success = download_voice(self.voices_dir, locale, name, quality)
            
            if success:
                self.completed_count += 1
            else:
                self.failed_count += 1
        
        # Finished
        if self.cancelled:
            status = f"Cancelled. Downloaded {self.completed_count} voice(s)."
        elif self.failed_count > 0:
            status = f"Completed with {self.failed_count} error(s)."
        else:
            status = f"Successfully downloaded {self.completed_count} voice(s)!"
        
        self.dialog.after(0, lambda: self._finish(status))
    
    def _update_ui(self, voice_text, progress_text, status_text):
        """Update UI from main thread"""
        self.voice_label.config(text=voice_text)
        self.progress_label.config(text=progress_text)
        self.status_label.config(text=status_text)
    
    def _finish(self, status):
        """Handle download completion"""
        self.voice_label.config(text="Download Complete")
        self.status_label.config(text=status)
        self.cancel_btn.config(text="Close")
    
    def on_cancel(self):
        """Handle cancel/close"""
        if self.download_thread.is_alive():
            self.cancelled = True
            self.status_label.config(text="Cancelling after current download...")
        else:
            self.dialog.destroy()


class NoVoicesDialog:
    """Dialog shown when no TTS voices are found"""
    
    def __init__(self, parent, voices_dir):
        self.parent = parent
        self.voices_dir = voices_dir
        self.result = None  # "download" or None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("No Voices Found")
        self.dialog.geometry("450x200")
        self.dialog.resizable(False, False)
        self.dialog.configure(bg=DARK_BG)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center on parent
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 450) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 200) // 3
        self.dialog.geometry(f"+{x}+{y}")
        
        self.create_widgets()
        
        # Handle window close
        self.dialog.protocol("WM_DELETE_WINDOW", self.on_cancel)
        
        # Wait for dialog to close
        self.dialog.wait_window()
    
    def create_widgets(self):
        # Main frame
        main_frame = tk.Frame(self.dialog, bg=DARK_BG, padx=25, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Message
        tk.Label(
            main_frame,
            text="No TTS voices found",
            font=("Segoe UI", 12, "bold"),
            fg=DARK_FG,
            bg=DARK_BG
        ).pack(pady=(0, 10))
        
        tk.Label(
            main_frame,
            text=f"Voice models are needed for audio export.\n\nVoices directory: {self.voices_dir}",
            font=("Segoe UI", 10),
            fg=DARK_FG_SECONDARY,
            bg=DARK_BG,
            wraplength=400,
            justify=tk.CENTER
        ).pack(pady=(0, 20))
        
        # Buttons frame
        btn_frame = tk.Frame(main_frame, bg=DARK_BG)
        btn_frame.pack()
        
        tk.Button(
            btn_frame,
            text="Download Voices",
            font=("Segoe UI", 10, "bold"),
            bg=ACCENT_COLOR,
            fg="white",
            activebackground="#9A6229",
            activeforeground="white",
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor="hand2",
            command=self.on_download
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        tk.Button(
            btn_frame,
            text="Cancel",
            font=("Segoe UI", 10),
            bg=DARK_BG_SECONDARY,
            fg=DARK_FG,
            activebackground="#4E4E4E",
            activeforeground=DARK_FG,
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor="hand2",
            command=self.on_cancel
        ).pack(side=tk.LEFT)
    
    def on_download(self):
        """User wants to download voices"""
        self.result = "download"
        self.dialog.destroy()
    
    def on_cancel(self):
        """User cancelled"""
        self.result = None
        self.dialog.destroy()


class AudioExportProgressDialog:
    """Progress dialog for audio export"""
    
    def __init__(self, parent, title="Exporting Audio Track"):
        self.parent = parent
        self.cancelled = False
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("400x150")
        self.dialog.resizable(False, False)
        self.dialog.configure(bg=DARK_BG)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center on parent
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 400) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 150) // 3
        self.dialog.geometry(f"+{x}+{y}")
        
        # Main frame
        main_frame = tk.Frame(self.dialog, bg=DARK_BG, padx=25, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Status message
        self.status_var = tk.StringVar(value="Initializing...")
        self.status_label = tk.Label(
            main_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 10),
            fg=DARK_FG,
            bg=DARK_BG,
            wraplength=350
        )
        self.status_label.pack(pady=(0, 10))
        
        # Progress info
        self.progress_var = tk.StringVar(value="0 / 0")
        self.progress_label = tk.Label(
            main_frame,
            textvariable=self.progress_var,
            font=("Segoe UI", 9),
            fg=DARK_FG_SECONDARY,
            bg=DARK_BG
        )
        self.progress_label.pack(pady=(0, 15))
        
        # Cancel button
        self.cancel_btn = tk.Button(
            main_frame,
            text="Cancel",
            font=("Segoe UI", 10),
            bg=DARK_BG_SECONDARY,
            fg=DARK_FG,
            activebackground="#4E4E4E",
            activeforeground=DARK_FG,
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor="hand2",
            command=self.on_cancel
        )
        self.cancel_btn.pack()
        
        # Handle window close
        self.dialog.protocol("WM_DELETE_WINDOW", self.on_cancel)
    
    def update_progress(self, current: int, total: int, message: str):
        """Update the progress display"""
        self.status_var.set(message)
        self.progress_var.set(f"{current} / {total}")
        self.dialog.update()
    
    def on_cancel(self):
        """Handle cancel"""
        self.cancelled = True
        self.dialog.destroy()
    
    def close(self):
        """Close the dialog"""
        if self.dialog.winfo_exists():
            self.dialog.destroy()


class VideoCaptionerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Panopto Audio Descriptions List Editor")
        
        # Set window size and center on preferred monitor
        window_width = 1200
        window_height = 800
        x, y = get_monitor_geometry(window_width, window_height)
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Set window icon using resource path helper
        load_app_icon(self.root)
        
        # Apply dark title bar on Windows
        self.root.after(100, lambda: enable_dark_title_bar(self.root))
        
        # State
        self.project = ProjectState()
        self.model = None  # Moondream model
        
        # Custom prompts (start with defaults)
        self.custom_prompts = PROMPTS.copy()
        
        # Controllers
        self.video = VideoController()
        self.audio = AudioController()
        
        # UI state
        self.is_running = True
        self._updating_timeline = False
        self.selected_caption_id = None
        self.is_processing = False
        
        # Region selection state
        self.selection_start = None  # (x, y) when mouse pressed
        self.selection_end = None    # (x, y) when mouse released
        self.selection_rect_id = None  # Canvas rectangle ID
        self.is_selecting = False
        
        # Dynamic canvas sizing
        self.current_preview_width = PREVIEW_WIDTH
        self.current_preview_height = PREVIEW_HEIGHT
        self._resize_after_id = None  # For debouncing resize events
        
        # Audio export state
        self.selected_voice_path = get_default_voice()
        
        # Colors for dark mode
        self.colors = {"bg": DARK_BG, "bg2": DARK_BG_SECONDARY, "bg3": DARK_BG_TERTIARY,
                       "fg": DARK_FG, "fg2": DARK_FG_SECONDARY, "border": DARK_BORDER}
        
        # Set up video callbacks
        self.video.set_callbacks(
            on_frame=lambda f: self.root.after(0, lambda: self.update_preview(f)),
            on_position=lambda: self.root.after(0, self.update_timeline_display),
            on_end=lambda: self.root.after(0, self.stop_playback)
        )
        
        # Apply theme and build UI
        self.setup_styles()
        self.create_menu()
        self.create_widgets()
        
        # Start background threads
        self.video.start_playback_thread()
        self.start_autosave_thread()
        
        # Initialize Moondream connection
        self.root.after(100, self.initialize_moondream)
        
        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure root window
        self.root.configure(bg=DARK_BG)
        
        # Frame
        style.configure("TFrame", background=DARK_BG)
        
        # Labels
        style.configure("TLabel", background=DARK_BG, foreground=DARK_FG, font=("Segoe UI", 10))
        style.configure("Header.TLabel", background=DARK_BG, foreground=DARK_FG, font=("Segoe UI", 11, "bold"))
        style.configure("Status.TLabel", background=DARK_BG, foreground=DARK_FG_SECONDARY, font=("Segoe UI", 9))
        style.configure("Success.TLabel", background=DARK_BG, foreground=SUCCESS_COLOR, font=("Segoe UI", 9))
        style.configure("Error.TLabel", background=DARK_BG, foreground=ERROR_COLOR, font=("Segoe UI", 9))
        
        # Regular buttons
        style.configure("TButton",
                        background=DARK_BG_TERTIARY,
                        foreground=DARK_FG,
                        borderwidth=1,
                        focuscolor=ACCENT_COLOR,
                        font=("Segoe UI", 10),
                        padding=(10, 5))
        style.map("TButton",
                  background=[("active", DARK_BG_ELEVATED), ("pressed", DARK_BG_ELEVATED)],
                  foreground=[("disabled", DARK_FG_SECONDARY)])
        
        # Accent buttons (copper color)
        style.configure("Accent.TButton",
                        background=ACCENT_COLOR,
                        foreground="#E0E0E0",
                        borderwidth=0,
                        font=("Segoe UI", 10, "bold"),
                        padding=(10, 5))
        style.map("Accent.TButton",
                  background=[("active", ACCENT_COLOR_HOVER), ("pressed", ACCENT_COLOR_PRESSED)])
        
        # Scale/Slider
        style.configure("Horizontal.TScale",
                        background=DARK_BG,
                        troughcolor=DARK_BG_TERTIARY,
                        sliderthickness=15)
        style.map("Horizontal.TScale",
                  background=[("!disabled", ACCENT_COLOR), ("active", ACCENT_COLOR_HOVER)])
        
        # Radiobutton
        style.configure("TRadiobutton",
                        background=DARK_BG,
                        foreground=DARK_FG,
                        font=("Segoe UI", 10),
                        indicatorbackground=DARK_BG_TERTIARY)
        style.map("TRadiobutton",
                  background=[("active", DARK_BG)],
                  indicatorbackground=[("selected", ACCENT_COLOR), ("!selected", DARK_BG_TERTIARY)])
        
        # Scrollbar
        style.configure("TScrollbar",
                        background=DARK_BG_TERTIARY,
                        troughcolor=DARK_BG_SECONDARY,
                        borderwidth=0,
                        arrowcolor=DARK_FG)
        style.map("TScrollbar",
                  background=[("active", DARK_BG_ELEVATED)])
        
        # Separator
        style.configure("TSeparator", background=DARK_BORDER)
        
        # Spinbox
        style.configure("TSpinbox",
                        fieldbackground=DARK_BG_TERTIARY,
                        foreground=DARK_FG,
                        borderwidth=1,
                        arrowcolor=DARK_FG)
        
        # Notebook (tabs)
        style.configure("TNotebook", background=DARK_BG, borderwidth=0)
        style.configure("TNotebook.Tab",
                        background=DARK_BG_TERTIARY,
                        foreground=DARK_FG,
                        padding=(10, 5))
        style.map("TNotebook.Tab",
                  background=[("selected", ACCENT_COLOR)],
                  foreground=[("selected", "#E0E0E0")])
        
        # PanedWindow - use tk.PanedWindow for better sash control
        # Note: ttk.PanedWindow sash styling is very limited
        style.configure("TPanedwindow", background=DARK_BG)
    
    def create_menu(self):
        """Create the menu bar"""
        menubar = tk.Menu(self.root, bg=DARK_BG, fg=DARK_FG,
                         activebackground=ACCENT_COLOR, activeforeground="#E0E0E0")
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0, bg=DARK_BG, fg=DARK_FG,
                           activebackground=ACCENT_COLOR, activeforeground="#E0E0E0")
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Load Video...", command=self.load_video)
        file_menu.add_command(label="Load Project...", command=self.load_project)
        file_menu.add_separator()
        file_menu.add_command(label="Save Project...", command=self.save_project)
        file_menu.add_command(label="Export WebVTT...", command=self.export_webvtt)
        file_menu.add_command(label="Export Audio Track...", command=self.export_audio_track)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_close)
        
        # Settings menu
        settings_menu = tk.Menu(menubar, tearoff=0, bg=DARK_BG, fg=DARK_FG,
                               activebackground=ACCENT_COLOR, activeforeground="#E0E0E0")
        menubar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Edit AI Prompts...", command=self.open_prompt_editor)
        settings_menu.add_separator()
        settings_menu.add_command(label="Download TTS Voices...", command=self.open_voice_download)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0, bg=DARK_BG, fg=DARK_FG,
                           activebackground=ACCENT_COLOR, activeforeground="#E0E0E0")
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Audio Description Guidelines...", command=self.show_guidelines)
        help_menu.add_separator()
        help_menu.add_command(label="About...", command=self.show_about)
    
    def show_guidelines(self):
        """Show audio description guidelines dialog"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Audio Description Guidelines")
        dialog.geometry("500x600")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg=DARK_BG)
        
        # Center dialog on parent window
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 3
        dialog.geometry(f"+{x}+{y}")
        
        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        ttk.Label(frame, text="Audio Description Guidelines", 
                  style="Header.TLabel", font=("Segoe UI", 14, "bold")).pack(anchor=tk.W)
        
        ttk.Label(frame, text="Quick reference for writing quality descriptions",
                  style="Status.TLabel").pack(anchor=tk.W, pady=(0, 15))
        
        # Guidelines text
        guidelines = """What to Describe:
- Actions and movements essential to understanding
- People by appearance ("a woman in a red jacket")
- On-screen text, signs, name tags, slide titles
- Scene changes and settings
- Charts and graphs (type, colors, trends)

Style Rules:
- Use present tense ("walks" not "walked")
- Be objective ("frowns" not "looks angry")
- Be concise - prioritize essential information
- Start with context, then add details

For Slides:
- State the title first
- Summarize key points (don't read everything)
- Describe any images, charts, or diagrams"""
        
        text_widget = tk.Text(frame, wrap=tk.WORD, font=("Segoe UI", 10),
                             bg=DARK_BG_TERTIARY, fg=DARK_FG,
                             padx=15, pady=15, height=14,
                             borderwidth=0, highlightthickness=1,
                             highlightbackground=DARK_BORDER)
        text_widget.insert(tk.END, guidelines)
        text_widget.config(state=tk.DISABLED)
        text_widget.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Link to DCMP
        link_frame = ttk.Frame(frame)
        link_frame.pack(fill=tk.X)
        
        ttk.Label(link_frame, text="For comprehensive guidelines, visit:",
                  style="Status.TLabel").pack(side=tk.LEFT)
        
        link = ttk.Label(link_frame, text="DCMP Description Key",
                        style="Status.TLabel", foreground=ACCENT_COLOR,
                        cursor="hand2", font=("Segoe UI", 9, "underline"))
        link.pack(side=tk.LEFT, padx=(5, 0))
        link.bind("<Button-1>", lambda e: self.open_url("https://dcmp.org/learn/descriptionkey"))
        
        # Close button
        ttk.Button(frame, text="Close", command=dialog.destroy).pack(pady=(15, 0))
    
    def show_about(self):
        """Show about dialog"""
        dialog = tk.Toplevel(self.root)
        dialog.title("About PADLE")
        dialog.geometry("450x350")
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 3
        dialog.geometry(f"+{x}+{y}")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg=DARK_BG)
        
        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Panopto Audio Descriptions List Editor", 
                  style="Header.TLabel", font=("Segoe UI", 14, "bold")).pack()
        ttk.Label(frame, text="Version 1.0",
                  style="Status.TLabel").pack(pady=(5, 15))
        ttk.Label(frame, text="Audio Description Generator for Recorded Videos\nCreates WebVTT descriptions for Panopto\nand other video platforms.",
                  style="TLabel", justify=tk.CENTER).pack()
        ttk.Label(frame, text="Developed by Perry J. Ganchuk\nBuilt with Moondream AI and Claude\nUniversity of Pittsburgh Center for Teaching and Learning, 2026",
                  style="Status.TLabel", justify=tk.CENTER).pack(pady=(15, 0))
        
        ttk.Button(frame, text="Close", command=dialog.destroy).pack(pady=(20, 0))
    
    def open_url(self, url):
        """Open a URL in the default browser"""
        import webbrowser
        webbrowser.open(url)
    
    def open_prompt_editor(self):
        """Open the prompt editor dialog"""
        PromptEditorDialog(self.root, self.custom_prompts, self.save_custom_prompts, 
                          colors=self.colors)
    
    def save_custom_prompts(self, prompts):
        """Save custom prompts from the editor"""
        self.custom_prompts = prompts
        self.project.custom_prompts = prompts.copy()
    
    def open_voice_download(self):
        """Open the voice download dialog"""
        dialog = VoiceDownloadDialog(self.root)
        if dialog.result:
            # Voices were downloaded - refresh the default voice if none selected
            if not self.selected_voice_path:
                self.selected_voice_path = get_default_voice()
    
    def create_widgets(self):
        # Main container with padding
        main_container = ttk.Frame(self.root, padding="10")
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Create vertical PanedWindow (top/bottom split) - using tk.PanedWindow for sash control
        self.main_paned = tk.PanedWindow(
            main_container, 
            orient=tk.VERTICAL,
            bg=DARK_BORDER,
            sashwidth=6,
            sashrelief=tk.FLAT,
            sashpad=0,
            showhandle=False,
            opaqueresize=True,
            borderwidth=0
        )
        self.main_paned.pack(fill=tk.BOTH, expand=True)
        
        # =================================================================
        # TOP PANE: Video player and controls
        # =================================================================
        top_container = ttk.Frame(self.main_paned)
        self.main_paned.add(top_container, stretch="always")
        
        # Top pane content frame
        top_frame = ttk.Frame(top_container)
        top_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left: Video player - now expandable
        video_frame = ttk.Frame(top_frame)
        video_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Video file selection
        file_frame = ttk.Frame(video_frame)
        file_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(file_frame, text="Video File:", style="Header.TLabel").pack(side=tk.LEFT)
        
        self.file_path_var = tk.StringVar(value="No video loaded")
        self.file_path_label = ttk.Label(file_frame, textvariable=self.file_path_var, 
                                         style="Status.TLabel", width=50)
        self.file_path_label.pack(side=tk.LEFT, padx=(10, 10))
        
        self.load_btn = ttk.Button(file_frame, text="Load Video", 
                                   style="Accent.TButton",
                                   command=self.load_video)
        self.load_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.load_project_btn = ttk.Button(file_frame, text="Load Project", 
                                           command=self.load_project)
        self.load_project_btn.pack(side=tk.LEFT)
        
        # Video canvas container - maintains aspect ratio
        self.video_canvas_frame = ttk.Frame(video_frame)
        self.video_canvas_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 5))
        
        # Video canvas - now expands with container
        self.video_canvas = tk.Canvas(self.video_canvas_frame, 
                                      bg=DARK_BG,
                                      highlightthickness=1, 
                                      highlightbackground=DARK_BORDER)
        self.video_canvas.pack(fill=tk.BOTH, expand=True)
        
        # Bind resize event with debouncing
        self.video_canvas.bind("<Configure>", self._on_canvas_configure)
        
        # Bind mouse events for region selection
        self.video_canvas.bind("<Button-1>", self.on_canvas_click)
        self.video_canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.video_canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        
        # Video timeline
        timeline_frame = ttk.Frame(video_frame)
        timeline_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.time_label = ttk.Label(timeline_frame, text="00:00:00")
        self.time_label.pack(side=tk.LEFT)
        
        self.timeline = ttk.Scale(timeline_frame, from_=0, to=100, orient=tk.HORIZONTAL,
                                  command=self.on_timeline_change)
        self.timeline.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        
        self.timeline.bind("<Button-1>", self.on_timeline_click)
        
        self.duration_label = ttk.Label(timeline_frame, text="00:00:00")
        self.duration_label.pack(side=tk.RIGHT)
        
        # Playback controls
        controls_frame = ttk.Frame(video_frame)
        controls_frame.pack(fill=tk.X)
        
        self.skip_back_btn = ttk.Button(controls_frame, text="<< -5s", width=8,
                                        command=lambda: self.skip(-5))
        self.skip_back_btn.pack(side=tk.LEFT, padx=2)
        
        self.play_btn = ttk.Button(controls_frame, text="Play", width=10,
                                   style="Accent.TButton",
                                   command=self.toggle_playback)
        self.play_btn.pack(side=tk.LEFT, padx=5)
        
        self.skip_fwd_btn = ttk.Button(controls_frame, text="+5s >>", width=8,
                                       command=lambda: self.skip(5))
        self.skip_fwd_btn.pack(side=tk.LEFT, padx=2)
        
        # Speed controls
        ttk.Label(controls_frame, text="Speed:", style="Status.TLabel").pack(side=tk.LEFT, padx=(20, 5))
        
        speeds = [("0.5x", 0.5), ("1x", 1.0), ("1.5x", 1.5), ("2x", 2.0)]
        self.speed_var = tk.DoubleVar(value=1.0)
        
        for label, speed in speeds:
            rb = ttk.Radiobutton(controls_frame, text=label, variable=self.speed_var,
                                value=speed, command=self.on_speed_change)
            rb.pack(side=tk.LEFT, padx=3)
        
        # Volume controls
        self.mute_btn = ttk.Button(controls_frame, text="Vol", width=3,
                                   command=self.toggle_mute)
        self.mute_btn.pack(side=tk.RIGHT, padx=(5, 2))
        
        self.volume_var = tk.IntVar(value=DEFAULT_VOLUME)
        self.volume_slider = ttk.Scale(controls_frame, from_=0, to=100, 
                                       orient=tk.HORIZONTAL, length=100,
                                       variable=self.volume_var,
                                       command=self.on_volume_change)
        self.volume_slider.pack(side=tk.RIGHT, padx=2)
        
        ttk.Label(controls_frame, text="Vol:", style="Status.TLabel").pack(side=tk.RIGHT, padx=(10, 2))
        
        # Right: Status and actions
        right_frame = ttk.Frame(top_frame, width=300, padding=(15, 10))
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        right_frame.pack_propagate(False)
        
        # Connection status
        ttk.Label(right_frame, text="Status", style="Header.TLabel").pack(anchor=tk.W)
        
        self.moondream_status = ttk.Label(right_frame, text="* Moondream: Connecting...",
                                          style="Status.TLabel")
        self.moondream_status.pack(anchor=tk.W, pady=(2, 0))
        
        self.video_status = ttk.Label(right_frame, text="* Video: Not loaded",
                                      style="Status.TLabel")
        self.video_status.pack(anchor=tk.W, pady=(2, 0))
        
        ttk.Separator(right_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # Description controls
        ttk.Label(right_frame, text="Describe Current Frame", style="Header.TLabel").pack(anchor=tk.W)
        
        self.describe_btn = ttk.Button(right_frame, text="Describe Now",
                                       style="Accent.TButton",
                                       command=self.describe_current_frame,
                                       state=tk.DISABLED)
        self.describe_btn.pack(fill=tk.X, pady=(5, 2))
        
        self.clear_selection_btn = ttk.Button(right_frame, text="Clear Selection",
                                              command=self.clear_selection,
                                              state=tk.DISABLED)
        self.clear_selection_btn.pack(fill=tk.X, pady=(0, 2))
        
        self.selection_label = ttk.Label(right_frame, text="", style="Status.TLabel")
        self.selection_label.pack(anchor=tk.W)
        
        # Processing status label
        self.processing_label = ttk.Label(right_frame, text="", style="Status.TLabel")
        self.processing_label.pack(anchor=tk.W)
        
        # Error details button (on separate line, starts hidden)
        self.error_details_btn = tk.Button(
            right_frame,
            text="Show Details",
            font=("Segoe UI", 8),
            bg=DARK_BG_TERTIARY,
            fg=DARK_FG,
            activebackground=DARK_BG_ELEVATED,
            activeforeground=DARK_FG,
            relief=tk.FLAT,
            padx=8,
            pady=2,
            cursor="hand2",
            command=self.show_error_details
        )
        # Button starts hidden - will be packed when error occurs
        self.last_error_message = None
        
        # Mode selection
        ttk.Label(right_frame, text="Mode:", style="Status.TLabel").pack(anchor=tk.W, pady=(10, 5))
        
        self.mode_var = tk.StringVar(value="general")
        modes = [("General", "general"), ("Slide", "slide"), ("Slide + OCR", "slide_ocr")]
        
        for text, value in modes:
            rb = ttk.Radiobutton(right_frame, text=text, variable=self.mode_var, value=value)
            rb.pack(anchor=tk.W, padx=(10, 0), pady=1)
        
        ttk.Separator(right_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # Export
        ttk.Label(right_frame, text="Export", style="Header.TLabel").pack(anchor=tk.W)
        
        self.save_project_btn = ttk.Button(right_frame, text="Save Project",
                                           command=self.save_project)
        self.save_project_btn.pack(fill=tk.X, pady=2)
        
        self.export_btn = ttk.Button(right_frame, text="Export WebVTT",
                                     command=self.export_webvtt)
        self.export_btn.pack(fill=tk.X, pady=2)
        
        self.export_audio_btn = ttk.Button(right_frame, text="Export Audio",
                                           command=self.export_audio_track)
        self.export_audio_btn.pack(fill=tk.X, pady=2)
        
        # =================================================================
        # BOTTOM PANE: Captions list and editor (with horizontal split)
        # =================================================================
        bottom_container = ttk.Frame(self.main_paned)
        self.main_paned.add(bottom_container, stretch="always")
        
        # Create horizontal PanedWindow (captions/editor split) - using tk.PanedWindow for sash control
        self.bottom_paned = tk.PanedWindow(
            bottom_container, 
            orient=tk.HORIZONTAL,
            bg=DARK_BORDER,
            sashwidth=6,
            sashrelief=tk.FLAT,
            sashpad=0,
            showhandle=False,
            opaqueresize=True,
            borderwidth=0
        )
        self.bottom_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left: Captions list
        captions_container = ttk.Frame(self.bottom_paned)
        self.bottom_paned.add(captions_container, stretch="always")
        
        captions_frame = ttk.Frame(captions_container)
        captions_frame.pack(fill=tk.BOTH, expand=True)
        
        captions_header = ttk.Frame(captions_frame)
        captions_header.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(captions_header, text="Captions", style="Header.TLabel").pack(side=tk.LEFT)
        
        self.captions_count = ttk.Label(captions_header, text="(0 captions)",
                                        style="Status.TLabel")
        self.captions_count.pack(side=tk.LEFT, padx=10)
        
        self.add_caption_btn = ttk.Button(captions_header, text="+ Add Manual",
                                          command=self.add_manual_caption)
        self.add_caption_btn.pack(side=tk.RIGHT)
        
        # Captions listbox
        captions_list_frame = ttk.Frame(captions_frame)
        captions_list_frame.pack(fill=tk.BOTH, expand=True)
        
        self.captions_listbox = tk.Listbox(captions_list_frame, 
                                           bg=DARK_BG_TERTIARY, fg=DARK_FG,
                                           selectmode=tk.SINGLE, font=("Segoe UI", 12),
                                           selectbackground=ACCENT_COLOR,
                                           selectforeground="#E0E0E0",
                                           activestyle='none',
                                           borderwidth=0, highlightthickness=1,
                                           highlightbackground=DARK_BORDER, 
                                           highlightcolor=ACCENT_COLOR)
        captions_scrollbar = ttk.Scrollbar(captions_list_frame, orient=tk.VERTICAL,
                                           command=self.captions_listbox.yview)
        self.captions_listbox.configure(yscrollcommand=captions_scrollbar.set)
        
        captions_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.captions_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.captions_listbox.bind('<<ListboxSelect>>', self.on_caption_select)
        self.captions_listbox.bind('<Double-1>', self.on_caption_double_click)
        
        # Right: Caption editor
        editor_container = ttk.Frame(self.bottom_paned)
        self.bottom_paned.add(editor_container, stretch="always")
        
        editor_frame = ttk.Frame(editor_container)
        editor_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(editor_frame, text="Edit Caption:", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 5))
        
        # Timecode editor - larger font for readability
        timecode_frame = ttk.Frame(editor_frame)
        timecode_frame.pack(fill=tk.X, pady=(0, 8))
        
        ttk.Label(timecode_frame, text="Timestamp:", font=("Segoe UI", 11)).pack(side=tk.LEFT)
        
        self.tc_hours = ttk.Spinbox(timecode_frame, from_=0, to=99, width=3, 
                                     font=("Segoe UI", 15), state=tk.DISABLED)
        self.tc_hours.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Label(timecode_frame, text=":", font=("Segoe UI", 15)).pack(side=tk.LEFT)
        
        self.tc_minutes = ttk.Spinbox(timecode_frame, from_=0, to=59, width=3,
                                       font=("Segoe UI", 15), state=tk.DISABLED)
        self.tc_minutes.pack(side=tk.LEFT)
        ttk.Label(timecode_frame, text=":", font=("Segoe UI", 15)).pack(side=tk.LEFT)
        
        self.tc_seconds = ttk.Spinbox(timecode_frame, from_=0, to=59, width=3,
                                       font=("Segoe UI", 15), state=tk.DISABLED)
        self.tc_seconds.pack(side=tk.LEFT)
        
        self.use_current_time_btn = ttk.Button(timecode_frame, text="Use Current",
                                               command=self.use_current_timestamp,
                                               state=tk.DISABLED)
        self.use_current_time_btn.pack(side=tk.LEFT, padx=(15, 0))
        
        # Text size controls frame
        size_frame = ttk.Frame(editor_frame)
        size_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(size_frame, text="Text Size:", style="Status.TLabel").pack(side=tk.LEFT)
        
        self.editor_font_size = 14  # Default font size
        
        self.font_decrease_btn = tk.Button(
            size_frame,
            text=" - ",
            font=("Segoe UI", 12, "bold"),
            bg=DARK_BG_TERTIARY,
            fg=DARK_FG,
            activebackground=DARK_BG_ELEVATED,
            activeforeground=DARK_FG,
            relief=tk.FLAT,
            padx=8,
            pady=0,
            cursor="hand2",
            command=self.decrease_editor_font
        )
        self.font_decrease_btn.pack(side=tk.LEFT, padx=(10, 2))
        
        self.font_size_label = ttk.Label(size_frame, text="14", width=3, 
                                          font=("Segoe UI", 10), anchor="center")
        self.font_size_label.pack(side=tk.LEFT)
        
        self.font_increase_btn = tk.Button(
            size_frame,
            text=" + ",
            font=("Segoe UI", 12, "bold"),
            bg=DARK_BG_TERTIARY,
            fg=DARK_FG,
            activebackground=DARK_BG_ELEVATED,
            activeforeground=DARK_FG,
            relief=tk.FLAT,
            padx=8,
            pady=0,
            cursor="hand2",
            command=self.increase_editor_font
        )
        self.font_increase_btn.pack(side=tk.LEFT, padx=(2, 0))
        
        # Caption text editor
        self.caption_editor = tk.Text(editor_frame, height=8, wrap=tk.WORD,
                                      font=("Segoe UI", self.editor_font_size), 
                                      bg=DARK_BG_TERTIARY, fg=DARK_FG,
                                      insertbackground=DARK_FG,
                                      selectbackground=ACCENT_COLOR,
                                      selectforeground="#E0E0E0",
                                      borderwidth=0, highlightthickness=1,
                                      highlightbackground=DARK_BORDER, 
                                      highlightcolor=ACCENT_COLOR,
                                      padx=10, pady=10)
        self.caption_editor.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        editor_buttons = ttk.Frame(editor_frame)
        editor_buttons.pack(fill=tk.X)
        
        self.save_caption_btn = ttk.Button(editor_buttons, text="Save Changes",
                                           style="Accent.TButton",
                                           command=self.save_caption_changes,
                                           state=tk.DISABLED)
        self.save_caption_btn.pack(side=tk.LEFT, padx=2)
        
        self.delete_caption_btn = ttk.Button(editor_buttons, text="Delete Caption",
                                             command=self.delete_selected_caption,
                                             state=tk.DISABLED)
        self.delete_caption_btn.pack(side=tk.LEFT, padx=2)
        
        self.goto_caption_btn = ttk.Button(editor_buttons, text="Go to Timestamp",
                                           command=self.goto_selected_caption,
                                           state=tk.DISABLED)
        self.goto_caption_btn.pack(side=tk.LEFT, padx=2)
    
    # =========================================================================
    # INITIALIZATION
    # =========================================================================
    
    def initialize_moondream(self):
        """Initialize local Moondream model"""
        self.moondream_status.config(text="* Moondream: Loading...", foreground=DARK_FG_SECONDARY)
        self.root.update()
        
        # Show loading dialog
        loading_dialog = ModelLoadingDialog(self.root)
        
        def load_model():
            try:
                model = get_model()
                
                def progress_callback(message):
                    loading_dialog.update_status(message)
                
                success = model.load(progress_callback=progress_callback)
                
                if success:
                    self.model = model
                    loading_dialog.finish()
                else:
                    loading_dialog.finish(error=model.load_error or "Unknown error")
                    
            except Exception as e:
                loading_dialog.finish(error=str(e))
        
        # Start loading in background thread
        load_thread = threading.Thread(target=load_model, daemon=True)
        load_thread.start()
        
        # Wait for dialog to close
        loading_dialog.wait()
        
        # Check results
        if loading_dialog.cancelled:
            self.root.destroy()
            return
        
        if loading_dialog.load_error:
            self.model = None
            self.moondream_status.config(
                text=f"* Moondream: Error", 
                foreground=ERROR_COLOR
            )
            messagebox.showerror(
                "Model Load Error",
                f"Failed to load AI model:\n\n{loading_dialog.load_error}\n\n"
                "The app will continue but AI descriptions will be unavailable."
            )
        else:
            model = get_model()
            device_name = {
                "cuda": "GPU",
                "mps": "Apple Silicon",
                "cpu": "CPU"
            }.get(model.device, model.device)
            self.moondream_status.config(
                text=f"* Moondream: Ready ({device_name})", 
                foreground=SUCCESS_COLOR
            )
        
        self.update_button_states()
    
    # =========================================================================
    # VIDEO LOADING
    # =========================================================================
    
    def load_video(self):
        """Load a video file"""
        filetypes = [
            ("Video files", "*.mp4 *.avi *.mkv *.mov *.webm *.m4v"),
            ("All files", "*.*")
        ]
        
        filepath = filedialog.askopenfilename(title="Select Video File", filetypes=filetypes)
        if not filepath:
            return
        
        if not self.video.load(filepath):
            messagebox.showerror("Error", f"Could not open video: {filepath}")
            return
        
        self.project.video_path = filepath
        self.audio.load(filepath)
        
        filename = os.path.basename(filepath)
        self.file_path_var.set(filename)
        self.video_status.config(text=f"* Video: {filename[:20]}...", foreground=SUCCESS_COLOR)
        self.duration_label.config(text=self.format_time(self.video.duration))
        self.timeline.config(to=self.video.duration)
        
        self.video.seek(0)
        self.update_button_states()
    
    def load_project(self):
        """Load a saved project"""
        filepath = filedialog.askopenfilename(
            title="Load Project",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not filepath:
            return
        
        try:
            self.project.load(filepath)
            
            if self.project.custom_prompts:
                self.custom_prompts = self.project.custom_prompts.copy()
            else:
                self.custom_prompts = PROMPTS.copy()
            
            if self.project.video_path and os.path.exists(self.project.video_path):
                if self.video.load(self.project.video_path):
                    filename = os.path.basename(self.project.video_path)
                    self.file_path_var.set(filename)
                    self.video_status.config(text=f"* Video: {filename[:20]}...", foreground=SUCCESS_COLOR)
                    self.duration_label.config(text=self.format_time(self.video.duration))
                    self.timeline.config(to=self.video.duration)
                    self.video.seek(0)
                    self.audio.load(self.project.video_path)
            
            self.refresh_captions_list()
            self.update_button_states()
            messagebox.showinfo("Success", f"Loaded project with {len(self.project.captions)} captions")
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not load project: {e}")
    
    # =========================================================================
    # REGION SELECTION
    # =========================================================================
    
    def on_canvas_click(self, event):
        """Handle mouse click on video canvas - start selection"""
        self.is_selecting = True
        self.selection_start = (event.x, event.y)
        self.selection_end = (event.x, event.y)
        
        # Remove existing selection rectangle
        if self.selection_rect_id:
            self.video_canvas.delete(self.selection_rect_id)
            self.selection_rect_id = None
    
    def on_canvas_drag(self, event):
        """Handle mouse drag on video canvas - update selection"""
        if not self.is_selecting:
            return
        
        self.selection_end = (event.x, event.y)
        
        # Remove old rectangle and draw new one
        if self.selection_rect_id:
            self.video_canvas.delete(self.selection_rect_id)
        
        x1, y1 = self.selection_start
        x2, y2 = self.selection_end
        
        self.selection_rect_id = self.video_canvas.create_rectangle(
            x1, y1, x2, y2,
            outline=ACCENT_COLOR,
            width=2,
            dash=(5, 3)
        )
    
    def on_canvas_release(self, event):
        """Handle mouse release on video canvas - finalize selection"""
        if not self.is_selecting:
            return
        
        self.is_selecting = False
        self.selection_end = (event.x, event.y)
        
        # Check if selection is too small (just a click)
        x1, y1 = self.selection_start
        x2, y2 = self.selection_end
        
        width = abs(x2 - x1)
        height = abs(y2 - y1)
        
        if width < 10 or height < 10:
            # Too small, clear selection
            self.clear_selection()
            return
        
        # Normalize coordinates (ensure x1 < x2 and y1 < y2)
        self.selection_start = (min(x1, x2), min(y1, y2))
        self.selection_end = (max(x1, x2), max(y1, y2))
        
        # Redraw rectangle with final coordinates
        if self.selection_rect_id:
            self.video_canvas.delete(self.selection_rect_id)
        
        x1, y1 = self.selection_start
        x2, y2 = self.selection_end
        
        self.selection_rect_id = self.video_canvas.create_rectangle(
            x1, y1, x2, y2,
            outline=ACCENT_COLOR,
            width=2
        )
        
        # Update UI
        self.clear_selection_btn.config(state=tk.NORMAL)
        self.selection_label.config(text=f"Selection: {width}x{height}px")
    
    def clear_selection(self):
        """Clear the region selection"""
        if self.selection_rect_id:
            self.video_canvas.delete(self.selection_rect_id)
            self.selection_rect_id = None
        
        self.selection_start = None
        self.selection_end = None
        self.clear_selection_btn.config(state=tk.DISABLED)
        self.selection_label.config(text="")
    
    def get_selected_region(self, frame):
        """Crop frame to selected region, or return full frame if no selection"""
        if not self.selection_start or not self.selection_end:
            return frame
        
        # Get selection coordinates (in canvas/preview coordinates)
        x1, y1 = self.selection_start
        x2, y2 = self.selection_end
        
        # Get original frame dimensions
        orig_h, orig_w = frame.shape[:2]
        
        # Scale selection coordinates to original frame size using current preview dimensions
        scale_x = orig_w / self.current_preview_width
        scale_y = orig_h / self.current_preview_height
        
        orig_x1 = int(x1 * scale_x)
        orig_y1 = int(y1 * scale_y)
        orig_x2 = int(x2 * scale_x)
        orig_y2 = int(y2 * scale_y)
        
        # Clamp to frame bounds
        orig_x1 = max(0, min(orig_x1, orig_w))
        orig_y1 = max(0, min(orig_y1, orig_h))
        orig_x2 = max(0, min(orig_x2, orig_w))
        orig_y2 = max(0, min(orig_y2, orig_h))
        
        # Crop the frame
        cropped = frame[orig_y1:orig_y2, orig_x1:orig_x2]
        
        return cropped
    
    def _on_canvas_configure(self, event):
        """Handle canvas resize with debouncing for smoother UI"""
        # Cancel any pending resize
        if self._resize_after_id:
            self.root.after_cancel(self._resize_after_id)
        
        # Schedule the actual resize handling with a small delay
        self._resize_after_id = self.root.after(50, lambda: self._handle_canvas_resize(event.width, event.height))
    
    def _handle_canvas_resize(self, width, height):
        """Actually handle the canvas resize after debounce"""
        self._resize_after_id = None
        
        # Maintain 16:9 aspect ratio
        target_ratio = 16 / 9
        current_ratio = width / max(height, 1)
        
        if current_ratio > target_ratio:
            # Too wide - constrain by height
            new_height = height
            new_width = int(height * target_ratio)
        else:
            # Too tall - constrain by width
            new_width = width
            new_height = int(width / target_ratio)
        
        # Minimum size
        new_width = max(new_width, 320)
        new_height = max(new_height, 180)
        
        # Update current dimensions
        self.current_preview_width = new_width
        self.current_preview_height = new_height
        
        # Refresh the current frame if we have one
        if self.video.last_frame is not None:
            self.update_preview(self.video.last_frame.copy())
    
    # =========================================================================
    # PLAYBACK CONTROLS
    # =========================================================================
    
    def update_preview(self, frame):
        """Update the video preview canvas"""
        try:
            # Use current dynamic dimensions
            width = self.current_preview_width
            height = self.current_preview_height
            
            frame = cv2.resize(frame, (width, height))
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            image = Image.fromarray(frame)
            photo = ImageTk.PhotoImage(image)
            self.video_canvas.photo = photo
            
            # Center the image in the canvas
            canvas_width = self.video_canvas.winfo_width()
            canvas_height = self.video_canvas.winfo_height()
            x_offset = (canvas_width - width) // 2
            y_offset = (canvas_height - height) // 2
            
            if not hasattr(self, 'video_image_id'):
                self.video_image_id = self.video_canvas.create_image(
                    x_offset, y_offset, anchor=tk.NW, image=photo)
            else:
                self.video_canvas.coords(self.video_image_id, x_offset, y_offset)
                self.video_canvas.itemconfig(self.video_image_id, image=photo)
            
            # Redraw selection rectangle if it exists (scaled to new size)
            if self.selection_start and self.selection_end and not self.is_selecting:
                if self.selection_rect_id:
                    self.video_canvas.delete(self.selection_rect_id)
                
                x1, y1 = self.selection_start
                x2, y2 = self.selection_end
                
                self.selection_rect_id = self.video_canvas.create_rectangle(
                    x1 + x_offset, y1 + y_offset, 
                    x2 + x_offset, y2 + y_offset,
                    outline=ACCENT_COLOR,
                    width=2
                )
        except Exception:
            pass
    
    def update_timeline_display(self):
        """Update timeline and time label"""
        self.time_label.config(text=self.format_time(self.video.current_position))
        self._updating_timeline = True
        self.timeline.set(self.video.current_position)
        self._updating_timeline = False
    
    def toggle_playback(self):
        if self.video.is_playing:
            self.stop_playback()
        else:
            self.start_playback()
    
    def start_playback(self):
        if not self.video.cap:
            return
        self.audio.seek(self.video.current_position)
        self.audio.play()
        self.video.play()
        self.play_btn.config(text="Pause")
    
    def stop_playback(self):
        self.video.pause()
        self.audio.pause()
        self.play_btn.config(text="Play")
    
    def seek_to(self, position: float):
        was_playing = self.video.is_playing
        if was_playing:
            self.stop_playback()
        self.video.seek(position)
        self.audio.seek(position)
        if was_playing:
            self.start_playback()
    
    def skip(self, seconds: float):
        self.seek_to(self.video.current_position + seconds)
    
    def on_timeline_change(self, value):
        if self._updating_timeline:
            return
        if not self.video.is_playing:
            self.seek_to(float(value))
    
    def on_timeline_click(self, event):
        if not self.video.cap:
            return
        widget_width = self.timeline.winfo_width()
        click_x = event.x
        percentage = max(0, min(1, click_x / widget_width))
        new_position = percentage * self.video.duration
        self.seek_to(new_position)
    
    def on_speed_change(self):
        speed = self.speed_var.get()
        self.video.set_speed(speed)
        self.audio.set_rate(speed)
    
    def on_volume_change(self, value):
        volume = int(float(value))
        self.audio.set_volume(volume)
        self.mute_btn.config(text="Mute" if volume == 0 else "Vol")
    
    def toggle_mute(self):
        is_muted = self.audio.toggle_mute()
        if is_muted:
            self.mute_btn.config(text="Mute")
            self.volume_var.set(0)
        else:
            self.mute_btn.config(text="Vol")
            self.volume_var.set(self.audio.volume_before_mute)
    
    def format_time(self, seconds: float) -> str:
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = int(seconds % 60)
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"
    
    # =========================================================================
    # DESCRIPTION GENERATION
    # =========================================================================
    
    def describe_current_frame(self):
        if not self.model or not self.video.cap:
            return
        
        frame = self.video.get_frame_at_position()
        if frame is None:
            return
        
        # Crop to selected region if there is one
        frame = self.get_selected_region(frame)
        
        self.is_processing = True
        self.describe_btn.config(state=tk.DISABLED)
        self.processing_label.config(text="Generating description...")
        
        timestamp = self.video.current_position
        mode = self.mode_var.get()
        
        def describe_thread():
            import time as t
            try:
                start = t.time()
                rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                image = Image.fromarray(rgb_frame)
                
                if mode == "slide_ocr":
                    self.root.after(0, lambda: self.processing_label.config(text="Extracting text..."))
                    ocr_text = pytesseract.image_to_string(image).strip()
                    prompt = self.custom_prompts[mode].format(ocr_text=ocr_text or "[No text detected]")
                else:
                    prompt = self.custom_prompts[mode]
                
                self.root.after(0, lambda: self.processing_label.config(text="Generating description..."))
                result = self.model.query(image, prompt)
                description = result["answer"]
                
                caption = self.project.add_caption(
                    timestamp=timestamp,
                    text=description,
                    mode=mode,
                    is_generated=True
                )
                
                self.root.after(0, self.refresh_captions_list)
                self.root.after(0, lambda cid=caption.id: self.select_caption(cid))
                self.root.after(0, lambda: self.processing_label.config(text="Done - Description added"))
                self.root.after(0, self.hide_error_details_btn)
                
            except Exception as e:
                import traceback
                full_error = f"{str(e)}\n\n--- Full Traceback ---\n{traceback.format_exc()}"
                short_msg = str(e)[:40]
                self.root.after(0, lambda msg=short_msg, full=full_error: self.show_error_with_details(msg, full))
            finally:
                self.is_processing = False
                self.root.after(0, self.update_button_states)
        
        threading.Thread(target=describe_thread, daemon=True).start()
    
    def show_error_with_details(self, short_msg, full_error):
        """Display error with option to show full details"""
        self.last_error_message = full_error
        self.processing_label.config(text=f"Error: {short_msg}...")
        # Pack right after the processing label
        self.error_details_btn.pack(after=self.processing_label, anchor=tk.W, pady=(2, 0))
    
    def hide_error_details_btn(self):
        """Hide the error details button"""
        self.error_details_btn.pack_forget()
        self.last_error_message = None
    
    def show_error_details(self):
        """Open dialog showing full error details"""
        if self.last_error_message:
            ErrorDetailsDialog(self.root, "Error Details", self.last_error_message)
    
    # =========================================================================
    # AUTO-SAVE
    # =========================================================================
    
    def start_autosave_thread(self):
        def autosave_loop():
            while self.is_running:
                time.sleep(AUTOSAVE_INTERVAL)
                if self.project.video_path and self.project.captions:
                    self.root.after(0, self.autosave)
        threading.Thread(target=autosave_loop, daemon=True).start()
    
    def autosave(self):
        if not self.project.video_path:
            return
        try:
            autosave_path = self.project.video_path + ".captioner_autosave.json"
            self.project.save(autosave_path)
        except Exception:
            pass
    
    # =========================================================================
    # CAPTIONS LIST
    # =========================================================================
    
    def refresh_captions_list(self):
        self.captions_listbox.delete(0, tk.END)
        
        for caption in self.project.captions:
            time_str = self.format_time(caption.timestamp)
            status = "[*]" if caption.is_reviewed else "[ ]"
            max_chars = 120
            text_preview = caption.text[:max_chars] + "..." if len(caption.text) > max_chars else caption.text
            text_preview = text_preview.replace('\n', ' ')
            
            line1 = f"{status} [{time_str}]"
            line2 = f"      {text_preview}"
            
            self.captions_listbox.insert(tk.END, line1)
            self.captions_listbox.insert(tk.END, line2)
        
        self.captions_count.config(text=f"({len(self.project.captions)} captions)")

    def on_caption_select(self, event):
        selection = self.captions_listbox.curselection()
        if not selection:
            return
        
        idx = selection[0] // 2
        if idx < len(self.project.captions):
            caption = self.project.captions[idx]
            self.selected_caption_id = caption.id
            
            line1_idx = idx * 2
            line2_idx = idx * 2 + 1
            self.captions_listbox.selection_clear(0, tk.END)
            self.captions_listbox.selection_set(line1_idx)
            self.captions_listbox.selection_set(line2_idx)
            
            self.caption_editor.delete(1.0, tk.END)
            self.caption_editor.insert(tk.END, caption.text)
            
            hours = int(caption.timestamp // 3600)
            minutes = int((caption.timestamp % 3600) // 60)
            seconds = int(caption.timestamp % 60)
            
            self.tc_hours.config(state=tk.NORMAL)
            self.tc_minutes.config(state=tk.NORMAL)
            self.tc_seconds.config(state=tk.NORMAL)
            
            self.tc_hours.delete(0, tk.END)
            self.tc_hours.insert(0, str(hours))
            self.tc_minutes.delete(0, tk.END)
            self.tc_minutes.insert(0, str(minutes))
            self.tc_seconds.delete(0, tk.END)
            self.tc_seconds.insert(0, str(seconds))
            
            self.save_caption_btn.config(state=tk.NORMAL)
            self.delete_caption_btn.config(state=tk.NORMAL)
            self.goto_caption_btn.config(state=tk.NORMAL)
            self.use_current_time_btn.config(state=tk.NORMAL)

    def on_caption_double_click(self, event):
        self.goto_selected_caption()
    
    def select_caption(self, caption_id: int):
        for i, caption in enumerate(self.project.captions):
            if caption.id == caption_id:
                line1_idx = i * 2
                self.captions_listbox.selection_clear(0, tk.END)
                self.captions_listbox.selection_set(line1_idx)
                self.captions_listbox.selection_set(line1_idx + 1)
                self.captions_listbox.see(line1_idx)
                self.on_caption_select(None)
                break

    def save_caption_changes(self):
        if not self.selected_caption_id:
            return
        
        new_text = self.caption_editor.get(1.0, tk.END).strip()
        
        try:
            hours = int(self.tc_hours.get())
            minutes = int(self.tc_minutes.get())
            seconds = int(self.tc_seconds.get())
            new_timestamp = hours * 3600 + minutes * 60 + seconds
        except ValueError:
            messagebox.showerror("Error", "Invalid timestamp values")
            return
        
        self.project.update_caption(self.selected_caption_id, new_text, new_timestamp)
        self.refresh_captions_list()
        self.select_caption(self.selected_caption_id)
    
    def increase_editor_font(self):
        """Increase caption editor font size"""
        if self.editor_font_size < 24:  # Max size
            self.editor_font_size += 2
            self.caption_editor.config(font=("Segoe UI", self.editor_font_size))
            self.font_size_label.config(text=str(self.editor_font_size))
    
    def decrease_editor_font(self):
        """Decrease caption editor font size"""
        if self.editor_font_size > 10:  # Min size
            self.editor_font_size -= 2
            self.caption_editor.config(font=("Segoe UI", self.editor_font_size))
            self.font_size_label.config(text=str(self.editor_font_size))
    
    def use_current_timestamp(self):
        if not self.selected_caption_id:
            return
        
        position = self.video.current_position
        hours = int(position // 3600)
        minutes = int((position % 3600) // 60)
        seconds = int(position % 60)
        
        self.tc_hours.delete(0, tk.END)
        self.tc_hours.insert(0, str(hours))
        self.tc_minutes.delete(0, tk.END)
        self.tc_minutes.insert(0, str(minutes))
        self.tc_seconds.delete(0, tk.END)
        self.tc_seconds.insert(0, str(seconds))
    
    def delete_selected_caption(self):
        if not self.selected_caption_id:
            return
        
        if messagebox.askyesno("Confirm Delete", "Delete this caption?"):
            self.project.delete_caption(self.selected_caption_id)
            self.selected_caption_id = None
            self.caption_editor.delete(1.0, tk.END)
            self.save_caption_btn.config(state=tk.DISABLED)
            self.delete_caption_btn.config(state=tk.DISABLED)
            self.goto_caption_btn.config(state=tk.DISABLED)
            self.use_current_time_btn.config(state=tk.DISABLED)
            
            self.tc_hours.delete(0, tk.END)
            self.tc_minutes.delete(0, tk.END)
            self.tc_seconds.delete(0, tk.END)
            self.tc_hours.config(state=tk.DISABLED)
            self.tc_minutes.config(state=tk.DISABLED)
            self.tc_seconds.config(state=tk.DISABLED)
            
            self.refresh_captions_list()
    
    def goto_selected_caption(self):
        if not self.selected_caption_id:
            return
        caption = self.project.get_caption_by_id(self.selected_caption_id)
        if caption:
            self.seek_to(caption.timestamp)
    
    def add_manual_caption(self):
        caption = self.project.add_caption(
            timestamp=self.video.current_position,
            text="[Enter description here]",
            mode="general",
            is_generated=False
        )
        self.refresh_captions_list()
        self.select_caption(caption.id)
    
    # =========================================================================
    # EXPORT
    # =========================================================================
    
    def save_project(self):
        filepath = filedialog.asksaveasfilename(
            title="Save Project",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not filepath:
            return
        
        try:
            self.project.save(filepath)
            messagebox.showinfo("Success", "Project saved successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save project: {e}")
    
    def export_webvtt(self):
        if not self.project.captions:
            messagebox.showwarning("Warning", "No captions to export")
            return
        
        if self.project.video_path:
            base = os.path.splitext(self.project.video_path)[0]
            suggested = base + "_Audio_Descriptions.txt"
        else:
            suggested = "Audio_Descriptions.txt"
        
        filepath = filedialog.asksaveasfilename(
            title="Export WebVTT",
            initialfile=os.path.basename(suggested),
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("VTT files", "*.vtt"), ("All files", "*.*")]
        )
        if not filepath:
            return
        
        try:
            self.project.export_webvtt(filepath)
            messagebox.showinfo("Success", f"Exported {len(self.project.captions)} captions to WebVTT")
        except Exception as e:
            messagebox.showerror("Error", f"Could not export: {e}")
    
    def export_audio_track(self):
        """Export captions as an audio description track (MP3)"""
        
        # Check prerequisites
        if not self.project.captions:
            messagebox.showwarning("Warning", "No captions to export")
            return
        
        if not self.video.cap or self.video.duration <= 0:
            messagebox.showwarning("Warning", "Please load a video first")
            return
        
        # Check TTS availability
        tts = PiperTTS()
        voices = tts.discover_voices()
        
        if not voices:
            # Show dialog offering to download voices
            no_voices = NoVoicesDialog(self.root, PIPER_VOICES_DIR)
            if no_voices.result == "download":
                # Open voice download dialog
                download_dialog = VoiceDownloadDialog(self.root)
                if download_dialog.result:
                    # Voices were downloaded - refresh
                    voices = tts.discover_voices()
                    self.selected_voice_path = get_default_voice()
            
            if not voices:
                return  # Still no voices
        
        if not tts._find_piper():
            messagebox.showerror(
                "Piper Not Found",
                "Piper TTS executable not found.\n\n"
                "Install with:\npip install piper-tts"
            )
            return
        
        # Show voice selection dialog
        voice_dialog = VoiceSelectionDialog(
            self.root, 
            current_voice_path=self.selected_voice_path
        )
        
        if not voice_dialog.result:
            return  # User cancelled
        
        self.selected_voice_path = voice_dialog.result
        
        # Get output path
        if self.project.video_path:
            base = os.path.splitext(self.project.video_path)[0]
            suggested = base + "_Audio_Descriptions.mp3"
        else:
            suggested = "Audio_Descriptions.mp3"
        
        filepath = filedialog.asksaveasfilename(
            title="Export Audio Description Track",
            initialfile=os.path.basename(suggested),
            defaultextension=".mp3",
            filetypes=[
                ("MP3 files", "*.mp3"),
                ("WAV files", "*.wav"),
                ("All files", "*.*")
            ]
        )
        
        if not filepath:
            return
        
        # Show progress dialog
        progress = AudioExportProgressDialog(self.root)
        
        def run_export():
            try:
                def progress_callback(current, total, message):
                    if not progress.cancelled:
                        self.root.after(0, lambda: progress.update_progress(current, total, message))
                
                success = export_audio_description_track(
                    captions=self.project.captions,
                    video_duration=self.video.duration,
                    output_path=filepath,
                    voice_path=self.selected_voice_path,
                    progress_callback=progress_callback
                )
                
                if not progress.cancelled:
                    self.root.after(0, progress.close)
                    if success:
                        self.root.after(0, lambda: messagebox.showinfo(
                            "Success",
                            f"Audio description track exported!\n\n"
                            f"{len(self.project.captions)} descriptions\n"
                            f"Duration: {self.format_time(self.video.duration)}\n\n"
                            f"Saved to:\n{filepath}"
                        ))
            
            except Exception as e:
                self.root.after(0, progress.close)
                self.root.after(0, lambda: messagebox.showerror(
                    "Export Failed",
                    f"Could not export audio track:\n\n{str(e)}"
                ))
        
        # Run in thread to keep UI responsive
        threading.Thread(target=run_export, daemon=True).start()
    
    # =========================================================================
    # UTILITIES
    # =========================================================================
    
    def update_button_states(self):
        video_loaded = self.video.cap is not None
        moondream_connected = self.model is not None
        
        if video_loaded and moondream_connected and not self.is_processing:
            self.describe_btn.config(state=tk.NORMAL)
        else:
            self.describe_btn.config(state=tk.DISABLED)
    
    def on_close(self):
        self.is_running = False
        self.video.stop()
        self.audio.stop()
        
        if self.project.captions:
            if messagebox.askyesno("Save Project?", "Do you want to save the project before closing?"):
                self.save_project()
        
        # Unload the model to free memory
        if self.model is not None:
            try:
                self.model.unload()
            except:
                pass
        
        self.video.release()
        self.audio.release()
        self.root.destroy()