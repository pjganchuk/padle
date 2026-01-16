"""
Audio playback controller using VLC
"""

import logging

try:
    import vlc
    HAS_VLC = True
except ImportError:
    HAS_VLC = False
    vlc = None

from config import DEFAULT_VOLUME

logger = logging.getLogger(__name__)


class AudioController:
    """Handles audio playback using VLC"""
    
    def __init__(self):
        self.is_muted = False
        self.volume_before_mute = DEFAULT_VOLUME
        self.vlc_instance = None
        self.player = None
        
        if not HAS_VLC:
            logger.warning("VLC not installed - audio playback disabled")
            return
        
        try:
            self.vlc_instance = vlc.Instance('--no-video')  # Audio only
            self.player = self.vlc_instance.media_player_new()
            self.player.audio_set_volume(DEFAULT_VOLUME)
        except Exception as e:
            logger.error(f"Failed to initialize VLC: {e}")
            self.vlc_instance = None
            self.player = None
    
    def load(self, filepath: str):
        """Load audio from video file"""
        if not self.player:
            return
        try:
            media = self.vlc_instance.media_new(filepath)
            self.player.set_media(media)
        except Exception as e:
            logger.error(f"Error loading audio: {e}")
    
    def play(self):
        """Start audio playback"""
        if self.player:
            self.player.play()
    
    def pause(self):
        """Pause audio playback"""
        if self.player:
            self.player.pause()
    
    def stop(self):
        """Stop audio playback"""
        if self.player:
            self.player.stop()
    
    def seek(self, position: float):
        """Seek to position in seconds"""
        if not self.player:
            return
        try:
            # VLC set_time takes milliseconds
            self.player.set_time(int(position * 1000))
        except Exception:
            pass
    
    def set_volume(self, volume: int):
        """Set volume (0-100)"""
        if not self.player:
            return
        self.player.audio_set_volume(volume)
        if volume > 0:
            self.is_muted = False
            self.volume_before_mute = volume
    
    def get_volume(self) -> int:
        """Get current volume"""
        if not self.player:
            return self.volume_before_mute
        return self.player.audio_get_volume()
    
    def set_rate(self, rate: float):
        """Set playback rate (1.0 = normal speed)"""
        if self.player:
            self.player.set_rate(rate)
    
    def mute(self):
        """Mute audio"""
        if not self.player:
            return
        if not self.is_muted:
            self.volume_before_mute = self.get_volume()
            self.player.audio_set_volume(0)
            self.is_muted = True
    
    def unmute(self):
        """Unmute audio and restore previous volume"""
        if not self.player:
            return
        if self.is_muted:
            self.player.audio_set_volume(self.volume_before_mute)
            self.is_muted = False
    
    def toggle_mute(self) -> bool:
        """Toggle mute state, returns new mute state"""
        if self.is_muted:
            self.unmute()
        else:
            self.mute()
        return self.is_muted
    
    def release(self):
        """Release VLC resources"""
        if self.player:
            self.player.release()
        if self.vlc_instance:
            self.vlc_instance.release()
