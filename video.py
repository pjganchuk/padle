"""
Video playback controller using OpenCV
"""

import cv2
import threading
import time
from typing import Optional, Callable
import numpy as np


class VideoController:
    """Handles video playback using OpenCV"""
    
    def __init__(self):
        self.cap: Optional[cv2.VideoCapture] = None
        self.video_path: Optional[str] = None
        self.duration = 0
        self.fps = 30
        self.current_position = 0
        self.is_playing = False
        self.playback_speed = 1.0
        self.last_frame: Optional[np.ndarray] = None
        
        self._lock = threading.Lock()
        self._is_running = True
        self._on_frame_callback: Optional[Callable] = None
        self._on_position_callback: Optional[Callable] = None
        self._on_end_callback: Optional[Callable] = None
    
    def load(self, filepath: str) -> bool:
        """Load a video file. Returns True on success."""
        with self._lock:
            if self.cap:
                self.cap.release()
            
            self.cap = cv2.VideoCapture(filepath)
            
            if not self.cap.isOpened():
                return False
            
            self.video_path = filepath
            
            # Detect actual FPS
            detected_fps = self.cap.get(cv2.CAP_PROP_FPS)
            if detected_fps and detected_fps > 0:
                self.fps = detected_fps
            else:
                self.fps = 30
            
            total_frames = self.cap.get(cv2.CAP_PROP_FRAME_COUNT)
            self.duration = total_frames / self.fps
            self.current_position = 0
            
            return True
    
    def set_callbacks(self, on_frame: Callable = None, on_position: Callable = None, 
                      on_end: Callable = None):
        """Set callback functions for frame updates, position changes, and playback end"""
        self._on_frame_callback = on_frame
        self._on_position_callback = on_position
        self._on_end_callback = on_end
    
    def start_playback_thread(self):
        """Start the video playback thread"""
        def video_loop():
            last_position_update = 0
            position_update_interval = 0.1  # Update timeline 10x per second, not every frame
            
            while self._is_running:
                frame_start = time.time()
                
                if self.is_playing and self.cap:
                    frame = None
                    position = 0
                    end_of_video = False
                    
                    with self._lock:
                        if self.cap.isOpened():
                            ret, frame = self.cap.read()
                            if ret:
                                self.last_frame = frame
                                frame_num = self.cap.get(cv2.CAP_PROP_POS_FRAMES)
                                position = frame_num / self.fps
                                self.current_position = position
                            else:
                                end_of_video = True
                    
                    if frame is not None:
                        if position >= self.duration:
                            if self._on_end_callback:
                                self._on_end_callback()
                            self.current_position = self.duration
                        else:
                            # Always update frame for smooth video
                            if self._on_frame_callback:
                                self._on_frame_callback(frame.copy())
                            # Throttle timeline/position updates
                            if frame_start - last_position_update >= position_update_interval:
                                if self._on_position_callback:
                                    self._on_position_callback()
                                last_position_update = frame_start
                    elif end_of_video:
                        if self._on_end_callback:
                            self._on_end_callback()
                else:
                    time.sleep(0.01)
                    continue
                
                # Maintain source fps for smooth playback, adjusted for speed
                elapsed = time.time() - frame_start
                target_frame_time = (1.0 / self.fps) / self.playback_speed
                sleep_time = max(target_frame_time - elapsed, 0.001)
                time.sleep(sleep_time)
        
        thread = threading.Thread(target=video_loop, daemon=True)
        thread.start()
    
    def play(self):
        """Start video playback"""
        if not self.cap:
            return
        
        # Seek to current position before starting playback
        acquired = self._lock.acquire(timeout=0.1)
        if acquired:
            try:
                frame_num = int(self.current_position * self.fps)
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, frame_num)
            finally:
                self._lock.release()
        
        self.is_playing = True
    
    def pause(self):
        """Pause video playback"""
        self.is_playing = False
    
    def seek(self, position: float):
        """Seek to a specific position in seconds"""
        if not self.cap:
            return
        
        self.current_position = max(0, min(position, self.duration))
        self._update_frame_from_position()
    
    def _update_frame_from_position(self):
        """Update the current frame based on position"""
        if not self.cap:
            return
        
        acquired = self._lock.acquire(timeout=0.1)
        if not acquired:
            return
        
        try:
            frame_num = int(self.current_position * self.fps)
            self.cap.set(cv2.CAP_PROP_POS_FRAMES, frame_num)
            
            ret, frame = self.cap.read()
            if ret:
                self.last_frame = frame
                if self._on_frame_callback:
                    self._on_frame_callback(frame.copy())
        finally:
            self._lock.release()
        
        if self._on_position_callback:
            self._on_position_callback()
    
    def skip(self, seconds: float):
        """Skip forward or backward by seconds"""
        self.seek(self.current_position + seconds)
    
    def set_speed(self, speed: float):
        """Set playback speed (1.0 = normal)"""
        self.playback_speed = speed
    
    def get_frame_at_position(self) -> Optional[np.ndarray]:
        """Get the current frame"""
        return self.last_frame.copy() if self.last_frame is not None else None
    
    def stop(self):
        """Stop the video controller"""
        self._is_running = False
        self.is_playing = False
    
    def release(self):
        """Release video resources"""
        self._is_running = False
        with self._lock:
            if self.cap:
                self.cap.release()
                self.cap = None
