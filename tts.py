"""
Text-to-Speech module using Piper TTS with voice selection support

Cross-platform compatible (Windows, macOS, Linux)
"""

import os
import subprocess
import tempfile
import json
from pathlib import Path
from typing import Optional, List
from dataclasses import dataclass
from config import PIPER_VOICES_DIR, PIPER_SPEED
from platform_utils import is_windows, get_exe_name, get_venv_bin_dir, get_subprocess_flags


@dataclass
class VoiceInfo:
    """Information about a Piper voice"""
    name: str           # Display name (e.g., "Amy (US, Medium)")
    path: str           # Full path to .onnx file
    locale: str         # e.g., "en_US"
    voice_name: str     # e.g., "amy"
    quality: str        # e.g., "medium"
    language: str       # e.g., "English"
    sample_rate: int    # e.g., 22050
    
    @property
    def display_name(self) -> str:
        """Formatted name for UI display"""
        locale_map = {
            "en_US": "US",
            "en_GB": "UK",
        }
        locale_short = locale_map.get(self.locale, self.locale)
        return f"{self.voice_name.title()} ({locale_short}, {self.quality})"


class PiperTTS:
    """Wrapper for Piper TTS engine with voice selection"""
    
    def __init__(self, voice_path: str = None, speed: float = None):
        """
        Initialize Piper TTS.
        
        Args:
            voice_path: Path to .onnx voice model file
            speed: Speech speed multiplier (1.0 = normal)
        """
        self.voice_path = voice_path
        self.speed = speed or PIPER_SPEED
        self._piper_cmd = None
        self._voices_cache = None
        
    def _find_piper(self) -> Optional[str]:
        """Find piper executable (cross-platform)"""
        if self._piper_cmd:
            return self._piper_cmd
        
        exe_name = get_exe_name("piper")
        bin_dir = get_venv_bin_dir()
        
        # Build list of candidate paths
        candidates = [exe_name]  # In PATH
        
        # Check current virtual environment first
        venv_path = os.environ.get("VIRTUAL_ENV")
        if venv_path:
            candidates.insert(0, os.path.join(venv_path, bin_dir, exe_name))
        
        # Common project locations (cross-platform)
        home = os.path.expanduser("~")
        project_dirs = ["Projects", "projects", "Dev", "dev"]
        project_names = ["leadr", "padle", "video-captioner"]
        
        for proj_dir in project_dirs:
            for proj_name in project_names:
                venv_piper = os.path.join(home, proj_dir, proj_name, "venv", bin_dir, exe_name)
                candidates.append(venv_piper)
        
        # User local bin (Linux/macOS)
        if not is_windows():
            candidates.append(os.path.join(home, ".local", "bin", "piper"))
        
        # Check each candidate
        subprocess_flags = get_subprocess_flags()
        for cmd in candidates:
            try:
                # On Windows, check if file exists before running
                if os.path.isfile(cmd) or cmd == exe_name:
                    kwargs = {
                        'capture_output': True,
                        'timeout': 5
                    }
                    kwargs.update(subprocess_flags)
                    result = subprocess.run([cmd, "--help"], **kwargs)
                    if result.returncode == 0:
                        self._piper_cmd = cmd
                        return cmd
            except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
                continue
                
        return None
    
    def discover_voices(self, voices_dir: str = None) -> List[VoiceInfo]:
        """
        Discover available Piper voices in a directory.
        
        Args:
            voices_dir: Directory to scan (default: PIPER_VOICES_DIR)
            
        Returns:
            List of VoiceInfo objects
        """
        voices_dir = voices_dir or PIPER_VOICES_DIR
        voices = []
        
        if not os.path.isdir(voices_dir):
            return voices
        
        # Find all .onnx files
        for filename in os.listdir(voices_dir):
            if not filename.endswith('.onnx'):
                continue
            
            onnx_path = os.path.join(voices_dir, filename)
            json_path = onnx_path + '.json'
            
            # Parse voice info from filename (e.g., en_US-amy-medium.onnx)
            base_name = filename[:-5]  # Remove .onnx
            parts = base_name.split('-')
            
            if len(parts) >= 3:
                locale = parts[0]
                voice_name = parts[1]
                quality = parts[2]
            else:
                # Fallback for non-standard names
                locale = "unknown"
                voice_name = base_name
                quality = "unknown"
            
            # Try to read additional info from JSON config
            sample_rate = 22050
            language = "English"
            
            if os.path.exists(json_path):
                try:
                    with open(json_path, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                    sample_rate = config.get('audio', {}).get('sample_rate', 22050)
                    language = config.get('language', {}).get('name_english', 'English')
                except Exception:
                    pass
            
            voice = VoiceInfo(
                name=base_name,
                path=onnx_path,
                locale=locale,
                voice_name=voice_name,
                quality=quality,
                language=language,
                sample_rate=sample_rate
            )
            voices.append(voice)
        
        # Sort by locale, then name, then quality
        quality_order = {'low': 0, 'medium': 1, 'high': 2}
        voices.sort(key=lambda v: (
            v.locale,
            v.voice_name,
            quality_order.get(v.quality, 99)
        ))
        
        return voices
    
    def get_voices(self, voices_dir: str = None) -> List[VoiceInfo]:
        """Get available voices (cached)"""
        if self._voices_cache is None:
            self._voices_cache = self.discover_voices(voices_dir)
        return self._voices_cache
    
    def refresh_voices(self, voices_dir: str = None) -> List[VoiceInfo]:
        """Refresh the voice cache"""
        self._voices_cache = None
        return self.get_voices(voices_dir)
    
    def set_voice(self, voice_path: str):
        """Set the voice to use"""
        self.voice_path = voice_path
    
    def is_available(self) -> bool:
        """Check if Piper TTS is available"""
        piper = self._find_piper()
        if not piper:
            return False
        if self.voice_path and not os.path.exists(self.voice_path):
            return False
        return True
    
    def get_status(self) -> str:
        """Get status message for UI"""
        piper = self._find_piper()
        if not piper:
            return "Piper TTS not found. Install with: pip install piper-tts"
        if self.voice_path and not os.path.exists(self.voice_path):
            return f"Voice model not found: {self.voice_path}"
        if not self.voice_path:
            return "No voice selected"
        return "Piper TTS: Ready"
    
    def synthesize(self, text: str, output_path: str) -> bool:
        """
        Generate speech from text.
        
        Args:
            text: Text to synthesize
            output_path: Path to save WAV file
            
        Returns:
            True on success, False on failure
        """
        piper = self._find_piper()
        if not piper:
            raise RuntimeError("Piper TTS not found")
        
        if not self.voice_path:
            raise RuntimeError("No voice selected")
        
        if not os.path.exists(self.voice_path):
            raise RuntimeError(f"Voice model not found: {self.voice_path}")
        
        try:
            # Build piper command
            cmd = [
                piper,
                "--model", self.voice_path,
                "--output_file", output_path,
            ]
            
            # Add speed if not default
            if self.speed != 1.0:
                cmd.extend(["--length_scale", str(1.0 / self.speed)])
            
            # Run piper with text as stdin
            kwargs = {
                'input': text.encode('utf-8'),
                'capture_output': True,
                'timeout': 60  # 1 minute timeout per synthesis
            }
            kwargs.update(get_subprocess_flags())
            
            result = subprocess.run(cmd, **kwargs)
            
            if result.returncode != 0:
                error = result.stderr.decode('utf-8', errors='replace')
                raise RuntimeError(f"Piper failed: {error}")
            
            return os.path.exists(output_path)
            
        except subprocess.TimeoutExpired:
            raise RuntimeError("TTS synthesis timed out")
    
    def synthesize_to_temp(self, text: str) -> str:
        """
        Generate speech to a temporary WAV file.
        
        Args:
            text: Text to synthesize
            
        Returns:
            Path to temporary WAV file (caller must delete)
        """
        fd, temp_path = tempfile.mkstemp(suffix='.wav')
        os.close(fd)
        
        try:
            self.synthesize(text, temp_path)
            return temp_path
        except Exception:
            # Clean up on failure
            if os.path.exists(temp_path):
                os.unlink(temp_path)
            raise


def get_default_voice(voices_dir: str = None) -> Optional[str]:
    """Get the default voice path (first US medium quality voice found)"""
    tts = PiperTTS()
    voices = tts.discover_voices(voices_dir)
    
    # Prefer US English, medium quality
    for voice in voices:
        if voice.locale == "en_US" and voice.quality == "medium":
            return voice.path
    
    # Fall back to any US voice
    for voice in voices:
        if voice.locale == "en_US":
            return voice.path
    
    # Fall back to any voice
    if voices:
        return voices[0].path
    
    return None