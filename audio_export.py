"""
Audio Description Track Export

Generates an MP3 audio track from captions using Piper TTS.
The track matches the video duration with descriptions at their timestamps.
"""

import os
import tempfile
from typing import List, Callable, Optional
from pydub import AudioSegment
from models import Caption
from tts import PiperTTS


class AudioTrackExporter:
    """Exports captions as an audio description track"""
    
    def __init__(self, tts: PiperTTS = None):
        """
        Initialize the exporter.
        
        Args:
            tts: PiperTTS instance (creates default if None)
        """
        self.tts = tts or PiperTTS()
        self._temp_files: List[str] = []
    
    def _cleanup_temp_files(self):
        """Remove any temporary files"""
        for path in self._temp_files:
            try:
                if os.path.exists(path):
                    os.unlink(path)
            except Exception:
                pass
        self._temp_files = []
    
    def export(
        self,
        captions: List[Caption],
        video_duration: float,
        output_path: str,
        progress_callback: Callable[[int, int, str], None] = None,
        sample_rate: int = 22050
    ) -> bool:
        """
        Export captions as an audio description track.
        
        Args:
            captions: List of Caption objects with timestamps and text
            video_duration: Total video duration in seconds
            output_path: Path to save the MP3 file
            progress_callback: Optional callback(current, total, message)
            sample_rate: Audio sample rate (default 22050 Hz)
            
        Returns:
            True on success
            
        Raises:
            RuntimeError: If TTS or audio processing fails
        """
        if not captions:
            raise ValueError("No captions to export")
        
        if video_duration <= 0:
            raise ValueError("Invalid video duration")
        
        try:
            total_steps = len(captions) + 2  # TTS for each + create base + export
            current_step = 0
            
            def update_progress(message: str):
                nonlocal current_step
                current_step += 1
                if progress_callback:
                    progress_callback(current_step, total_steps, message)
            
            # Step 1: Create silent base track matching video duration
            update_progress("Creating base audio track...")
            duration_ms = int(video_duration * 1000)
            base_track = AudioSegment.silent(duration=duration_ms, frame_rate=sample_rate)
            
            # Step 2: Generate TTS for each caption and overlay
            sorted_captions = sorted(captions, key=lambda c: c.timestamp)
            
            for i, caption in enumerate(sorted_captions):
                update_progress(f"Generating audio {i+1}/{len(captions)}...")
                
                # Skip empty captions
                if not caption.text.strip():
                    continue
                
                # Generate TTS audio
                temp_wav = self.tts.synthesize_to_temp(caption.text)
                self._temp_files.append(temp_wav)
                
                # Load the generated audio
                description_audio = AudioSegment.from_wav(temp_wav)
                
                # Calculate position in milliseconds
                position_ms = int(caption.timestamp * 1000)
                
                # Ensure we don't go past the end
                if position_ms >= duration_ms:
                    continue
                
                # Overlay at the correct position
                base_track = base_track.overlay(description_audio, position=position_ms)
            
            # Step 3: Export as MP3
            update_progress("Exporting MP3...")
            
            # Determine output format from extension
            output_ext = os.path.splitext(output_path)[1].lower()
            
            if output_ext == '.mp3':
                base_track.export(output_path, format='mp3', bitrate='192k')
            elif output_ext == '.wav':
                base_track.export(output_path, format='wav')
            else:
                # Default to MP3
                base_track.export(output_path, format='mp3', bitrate='192k')
            
            return True
            
        finally:
            self._cleanup_temp_files()
    
    def estimate_duration(self, captions: List[Caption]) -> float:
        """
        Estimate how long export will take.
        
        Args:
            captions: List of captions
            
        Returns:
            Estimated seconds to complete
        """
        # Rough estimate: ~2 seconds per caption for TTS
        return len(captions) * 2 + 5  # Plus overhead


def export_audio_description_track(
    captions: List[Caption],
    video_duration: float,
    output_path: str,
    voice_path: str = None,
    speech_speed: float = None,
    progress_callback: Callable[[int, int, str], None] = None
) -> bool:
    """
    Convenience function to export audio description track.
    
    Args:
        captions: List of Caption objects
        video_duration: Video duration in seconds
        output_path: Where to save the MP3
        voice_path: Path to Piper voice model (optional)
        speech_speed: Speech speed multiplier (optional)
        progress_callback: Progress callback function (optional)
        
    Returns:
        True on success
    """
    tts = PiperTTS(voice_path=voice_path, speed=speech_speed)
    
    if not tts.is_available():
        raise RuntimeError(tts.get_status())
    
    exporter = AudioTrackExporter(tts)
    return exporter.export(
        captions=captions,
        video_duration=video_duration,
        output_path=output_path,
        progress_callback=progress_callback
    )