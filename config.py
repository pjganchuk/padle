"""
Configuration constants for PADLE (Panopto Audio Descriptions List Editor)

Cross-platform compatible (Windows, macOS, Linux)
Updated for PyInstaller bundling support and local model inference.
"""

import os
import sys

# =============================================================================
# PLATFORM DETECTION
# =============================================================================

IS_WINDOWS = sys.platform == 'win32'
IS_MACOS = sys.platform == 'darwin'
IS_LINUX = sys.platform.startswith('linux')

# =============================================================================
# FROZEN/BUNDLED DETECTION
# =============================================================================

def is_frozen() -> bool:
    """Check if running as PyInstaller frozen executable"""
    return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')

IS_FROZEN = is_frozen()

# =============================================================================
# MOONDREAM AI MODEL CONFIGURATION
# =============================================================================

# Model identifier on Hugging Face Hub
MOONDREAM_MODEL_ID = "vikhyatk/moondream2"

# Model revision/version - pin to specific release for stability
MOONDREAM_MODEL_REVISION = "2025-01-09"

# Device selection: "auto", "cuda", "mps", or "cpu"
# "auto" will detect the best available device
MOONDREAM_DEVICE = "auto"

# Model cache directory (None = use default HuggingFace cache)
# Default: ~/.cache/huggingface on Linux/Mac, %USERPROFILE%\.cache\huggingface on Windows
MOONDREAM_CACHE_DIR = None

# =============================================================================
# VIDEO PREVIEW SETTINGS
# =============================================================================

PREVIEW_WIDTH = 800
PREVIEW_HEIGHT = 450

# =============================================================================
# AUDIO SETTINGS
# =============================================================================

DEFAULT_VOLUME = 75  # 0-100

# =============================================================================
# AUTO-SAVE SETTINGS
# =============================================================================

AUTOSAVE_INTERVAL = 60  # seconds

# =============================================================================
# PIPER TTS SETTINGS
# =============================================================================

def _get_default_voices_dir() -> str:
    """Get platform-appropriate default voices directory"""
    if IS_FROZEN:
        exe_dir = os.path.dirname(sys.executable)
        portable_voices = os.path.join(exe_dir, 'voices')
        if os.path.isdir(portable_voices):
            return portable_voices
    
    if IS_WINDOWS:
        local_app_data = os.environ.get('LOCALAPPDATA')
        if local_app_data:
            return os.path.join(local_app_data, 'piper-voices')
        return os.path.expanduser('~/piper-voices')
    else:
        return os.path.expanduser('~/piper-voices')

PIPER_VOICES_DIR = _get_default_voices_dir()

# Speech speed multiplier (1.0 = normal, 1.2 = faster, 0.8 = slower)
PIPER_SPEED = 1.0

# =============================================================================
# AUDIO EXPORT SETTINGS
# =============================================================================

AUDIO_EXPORT_SAMPLE_RATE = 22050  # Hz
AUDIO_EXPORT_BITRATE = "192k"    # MP3 bitrate

# Default voice (used if no preference saved)
# Set to None to auto-detect first available voice
PIPER_DEFAULT_VOICE = None

# =============================================================================
# APPLICATION DATA
# =============================================================================

def _get_app_data_dir() -> str:
    """Get platform-appropriate app data directory"""
    app_name = "padle"
    if IS_WINDOWS:
        base = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        return os.path.join(base, app_name)
    elif IS_MACOS:
        return os.path.expanduser(f'~/Library/Application Support/{app_name}')
    else:
        return os.path.expanduser(f'~/.local/share/{app_name}')

APP_DATA_DIR = _get_app_data_dir()

# =============================================================================
# APPLICATION INFO
# =============================================================================

APP_NAME = "PADLE"
APP_VERSION = "1.1.0"
APP_AUTHOR = "Perry J. Ganchuk"
APP_ORG = "University of Pittsburgh Center for Teaching and Learning"