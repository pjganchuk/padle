"""
Data models for Video Captioner
"""

import json
from dataclasses import dataclass, asdict
from typing import List, Optional, Dict


@dataclass
class Caption:
    """Represents a single audio description caption"""
    id: int
    timestamp: float  # Seconds from start
    text: str
    mode: str  # "general", "slide", or "slide_ocr"
    is_generated: bool = False  # True if AI-generated, False if manual
    is_reviewed: bool = False  # True if human has reviewed
    
    def to_dict(self) -> dict:
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: dict) -> 'Caption':
        return cls(**data)


class ProjectState:
    """Manages the project state for saving/loading"""
    
    def __init__(self):
        self.video_path: Optional[str] = None
        self.captions: List[Caption] = []
        self.next_caption_id: int = 1
        self.last_save_path: Optional[str] = None
        self.custom_prompts: Optional[Dict[str, str]] = None  # None means use defaults
    
    def add_caption(self, timestamp: float, text: str, mode: str, 
                    is_generated: bool = False) -> Caption:
        caption = Caption(
            id=self.next_caption_id,
            timestamp=timestamp,
            text=text,
            mode=mode,
            is_generated=is_generated,
            is_reviewed=not is_generated
        )
        self.captions.append(caption)
        self.next_caption_id += 1
        # Keep captions sorted by timestamp
        self.captions.sort(key=lambda c: c.timestamp)
        return caption
    
    def update_caption(self, caption_id: int, text: str, timestamp: float = None) -> Optional[Caption]:
        for caption in self.captions:
            if caption.id == caption_id:
                caption.text = text
                caption.is_reviewed = True
                if timestamp is not None:
                    caption.timestamp = timestamp
                    # Re-sort captions by timestamp
                    self.captions.sort(key=lambda c: c.timestamp)
                return caption
        return None
    
    def delete_caption(self, caption_id: int) -> bool:
        for i, caption in enumerate(self.captions):
            if caption.id == caption_id:
                self.captions.pop(i)
                return True
        return False
    
    def get_caption_by_id(self, caption_id: int) -> Optional[Caption]:
        for caption in self.captions:
            if caption.id == caption_id:
                return caption
        return None
    
    def to_dict(self) -> dict:
        data = {
            "video_path": self.video_path,
            "captions": [c.to_dict() for c in self.captions],
            "next_caption_id": self.next_caption_id,
        }
        # Only save custom prompts if they've been modified
        if self.custom_prompts is not None:
            data["custom_prompts"] = self.custom_prompts
        return data
    
    def save(self, filepath: str):
        with open(filepath, 'w') as f:
            json.dump(self.to_dict(), f, indent=2)
        self.last_save_path = filepath
    
    def load(self, filepath: str):
        with open(filepath, 'r') as f:
            data = json.load(f)
        self.video_path = data.get("video_path")
        self.captions = [Caption.from_dict(c) for c in data.get("captions", [])]
        self.next_caption_id = data.get("next_caption_id", 1)
        self.custom_prompts = data.get("custom_prompts")  # None if not in file
        self.last_save_path = filepath
    
    def export_webvtt(self, filepath: str, default_duration: float = 10.0):
        """Export captions to WebVTT format for Panopto"""
        lines = ["WEBVTT", ""]
        
        for i, caption in enumerate(self.captions):
            # Calculate end time
            if i + 1 < len(self.captions):
                # End at next caption or after default duration, whichever is sooner
                next_start = self.captions[i + 1].timestamp
                end_time = min(caption.timestamp + default_duration, next_start)
            else:
                # Last caption - use default duration based on text length
                words = len(caption.text.split())
                duration = max(default_duration, words * 0.4)  # ~150 words/min
                end_time = caption.timestamp + duration
            
            lines.append(str(i + 1))
            lines.append(f"{self._format_timestamp(caption.timestamp)} --> {self._format_timestamp(end_time)}")
            lines.append(caption.text)
            lines.append("")
        
        with open(filepath, 'w') as f:
            f.write('\n'.join(lines))
    
    def _format_timestamp(self, seconds: float) -> str:
        """Format seconds as HH:MM:SS.mmm"""
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = seconds % 60
        return f"{hours:02d}:{minutes:02d}:{secs:06.3f}"