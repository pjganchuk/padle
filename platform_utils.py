"""
Cross-platform utilities for PADLE

Provides platform-aware functionality for audio playback, 
file handling, and subprocess management.
"""

import os
import sys
import subprocess
import shutil
from typing import Optional


def is_windows() -> bool:
    """Check if running on Windows"""
    return sys.platform == 'win32'


def is_macos() -> bool:
    """Check if running on macOS"""
    return sys.platform == 'darwin'


def is_linux() -> bool:
    """Check if running on Linux"""
    return sys.platform.startswith('linux')


def get_exe_extension() -> str:
    """Get the executable file extension for the current platform"""
    if is_windows():
        return '.exe'
    return ''


def get_exe_name(base_name: str) -> str:
    """Get executable name with platform-appropriate extension"""
    return base_name + get_exe_extension()


def get_venv_bin_dir() -> str:
    """Get the platform-appropriate venv binary directory name"""
    if is_windows():
        return "Scripts"
    return "bin"


def get_subprocess_flags() -> dict:
    """
    Get platform-appropriate subprocess flags.
    
    Returns:
        Dict of kwargs to pass to subprocess.run() or Popen()
    """
    if is_windows():
        return {'creationflags': subprocess.CREATE_NO_WINDOW}
    return {}


def play_audio_file(filepath: str, blocking: bool = True) -> bool:
    """
    Play an audio file using the platform's native audio player.
    
    Args:
        filepath: Path to the audio file (WAV, MP3, etc.)
        blocking: If True, wait for playback to complete
        
    Returns:
        True if playback started successfully
    """
    if not os.path.exists(filepath):
        return False
    
    try:
        if is_windows():
            # Windows: Use the built-in Windows Media Player command line
            # or PowerShell's audio playback
            try:
                # Try using PowerShell's audio playback (works on all Windows)
                ps_cmd = f'(New-Object Media.SoundPlayer "{filepath}").PlaySync()'
                subprocess.run(
                    ['powershell', '-Command', ps_cmd],
                    capture_output=True,
                    timeout=60,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                return True
            except (FileNotFoundError, subprocess.TimeoutExpired):
                pass
            
            # Fallback: Try winsound for WAV files
            if filepath.lower().endswith('.wav'):
                try:
                    import winsound
                    flags = winsound.SND_FILENAME
                    if not blocking:
                        flags |= winsound.SND_ASYNC
                    winsound.PlaySound(filepath, flags)
                    return True
                except Exception:
                    pass
            
            # Last resort: open with default player
            os.startfile(filepath)
            return True
            
        elif is_macos():
            # macOS: Use afplay
            cmd = ['afplay', filepath]
            if blocking:
                subprocess.run(cmd, capture_output=True, timeout=60)
            else:
                subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
            
        else:
            # Linux: Try various players in order of preference
            players = [
                ['paplay', filepath],     # PulseAudio
                ['aplay', filepath],      # ALSA
                ['play', filepath],       # SoX
                ['mpv', '--no-video', filepath],  # mpv
                ['ffplay', '-nodisp', '-autoexit', filepath],  # ffmpeg
            ]
            
            for cmd in players:
                if shutil.which(cmd[0]):
                    try:
                        if blocking:
                            subprocess.run(cmd, capture_output=True, timeout=60)
                        else:
                            subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                        return True
                    except (subprocess.TimeoutExpired, OSError):
                        continue
            
            return False
            
    except Exception:
        return False


def get_default_voices_dir() -> str:
    """
    Get the default directory for Piper voice models.
    
    Returns platform-appropriate default path:
    - Windows: %LOCALAPPDATA%\\piper-voices
    - macOS/Linux: ~/piper-voices
    """
    if is_windows():
        local_app_data = os.environ.get('LOCALAPPDATA')
        if local_app_data:
            return os.path.join(local_app_data, 'piper-voices')
    
    # Default for macOS/Linux or Windows fallback
    return os.path.expanduser('~/piper-voices')


def get_app_data_dir(app_name: str = "padle") -> str:
    """
    Get the application data directory.
    
    Returns platform-appropriate path:
    - Windows: %LOCALAPPDATA%\\{app_name}
    - macOS: ~/Library/Application Support/{app_name}
    - Linux: ~/.local/share/{app_name}
    """
    if is_windows():
        base = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        return os.path.join(base, app_name)
    elif is_macos():
        return os.path.expanduser(f'~/Library/Application Support/{app_name}')
    else:
        return os.path.expanduser(f'~/.local/share/{app_name}')


def find_ffmpeg() -> Optional[str]:
    """
    Find the ffmpeg executable.
    
    Returns:
        Path to ffmpeg or None if not found
    """
    # Check if in PATH
    ffmpeg = shutil.which('ffmpeg')
    if ffmpeg:
        return ffmpeg
    
    # Check common locations
    if is_windows():
        common_paths = [
            os.path.join(os.environ.get('LOCALAPPDATA', ''), 'ffmpeg', 'bin', 'ffmpeg.exe'),
            os.path.join(os.environ.get('PROGRAMFILES', ''), 'ffmpeg', 'bin', 'ffmpeg.exe'),
            'C:\\ffmpeg\\bin\\ffmpeg.exe',
        ]
    elif is_macos():
        common_paths = [
            '/usr/local/bin/ffmpeg',
            '/opt/homebrew/bin/ffmpeg',
        ]
    else:
        common_paths = [
            '/usr/bin/ffmpeg',
            '/usr/local/bin/ffmpeg',
        ]
    
    for path in common_paths:
        if os.path.isfile(path):
            return path
    
    return None


def find_piper() -> Optional[str]:
    """
    Find the piper executable.
    
    Returns:
        Path to piper or None if not found
    """
    exe_name = get_exe_name('piper')
    
    # Check if in PATH
    piper = shutil.which(exe_name)
    if piper:
        return piper
    
    # Check current virtual environment
    venv_path = os.environ.get("VIRTUAL_ENV")
    if venv_path:
        venv_piper = os.path.join(venv_path, get_venv_bin_dir(), exe_name)
        if os.path.isfile(venv_piper):
            return venv_piper
    
    # Check common project locations
    home = os.path.expanduser("~")
    project_dirs = ["Projects", "projects", "Dev", "dev"]
    project_names = ["leadr", "padle", "video-captioner"]
    
    for proj_dir in project_dirs:
        for proj_name in project_names:
            venv_piper = os.path.join(home, proj_dir, proj_name, "venv", get_venv_bin_dir(), exe_name)
            if os.path.isfile(venv_piper):
                return venv_piper
    
    # User local bin (Linux/macOS)
    if not is_windows():
        local_piper = os.path.join(home, ".local", "bin", "piper")
        if os.path.isfile(local_piper):
            return local_piper
    
    return None


def check_dependencies() -> dict:
    """
    Check for required external dependencies.
    
    Returns:
        Dict with dependency names as keys and (available: bool, path: str or None) as values
    """
    deps = {}
    
    # FFmpeg (required for pydub MP3 export)
    ffmpeg = find_ffmpeg()
    deps['ffmpeg'] = (ffmpeg is not None, ffmpeg)
    
    # Piper TTS
    piper = find_piper()
    deps['piper'] = (piper is not None, piper)
    
    # VLC (for video audio playback)
    try:
        import vlc
        deps['vlc'] = (True, 'python-vlc')
    except ImportError:
        deps['vlc'] = (False, None)
    
    # Tesseract OCR
    tesseract = shutil.which('tesseract')
    deps['tesseract'] = (tesseract is not None, tesseract)
    
    return deps