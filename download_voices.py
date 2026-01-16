#!/usr/bin/env python3
"""
Download English Piper TTS voices from Hugging Face

Can be run as a script or imported for GUI use.

Usage (command line):
    python download_voices.py              # Download all English voices
    python download_voices.py --us-only    # Download only US English voices
    python download_voices.py --medium     # Download only medium quality
"""

import os
import sys
import argparse
import urllib.request
import urllib.error
from typing import Callable, Optional, Tuple

# Base URL for Piper voices
BASE_URL = "https://huggingface.co/rhasspy/piper-voices/resolve/main"

# Approximate file sizes in MB (onnx + json combined)
# These are estimates based on typical model sizes
SIZE_ESTIMATES = {
    "low": 16,
    "medium": 17,
    "high": 46,
}

# English voices available (as of 2025)
# Format: (locale, name, qualities_available)
ENGLISH_VOICES = [
    # US English
    ("en_US", "amy", ["low", "medium"]),
    ("en_US", "arctic", ["medium"]),
    ("en_US", "bryce", ["medium"]),
    ("en_US", "danny", ["low"]),
    ("en_US", "hfc_female", ["medium"]),
    ("en_US", "hfc_male", ["medium"]),
    ("en_US", "joe", ["medium"]),
    ("en_US", "john", ["medium"]),
    ("en_US", "kathleen", ["low"]),
    ("en_US", "kusal", ["medium"]),
    ("en_US", "l2arctic", ["medium"]),
    ("en_US", "lessac", ["low", "medium", "high"]),
    ("en_US", "libritts", ["high"]),
    ("en_US", "libritts_r", ["medium"]),
    ("en_US", "ljspeech", ["high", "medium"]),
    ("en_US", "norman", ["medium"]),
    ("en_US", "ryan", ["low", "medium", "high"]),
    
    # UK English
    ("en_GB", "alan", ["low", "medium"]),
    ("en_GB", "alba", ["medium"]),
    ("en_GB", "aru", ["medium"]),
    ("en_GB", "cori", ["medium", "high"]),
    ("en_GB", "jenny_dioco", ["medium"]),
    ("en_GB", "northern_english_male", ["medium"]),
    ("en_GB", "semaine", ["medium"]),
    ("en_GB", "southern_english_female", ["low"]),
    ("en_GB", "vctk", ["medium"]),
]


def get_voice_filename(locale: str, name: str, quality: str) -> str:
    """Get the base filename for a voice (without extension)"""
    return f"{locale}-{name}-{quality}"


def get_voice_urls(locale: str, name: str, quality: str) -> Tuple[str, str]:
    """Get the download URLs for a voice's onnx and json files"""
    voice_filename = get_voice_filename(locale, name, quality)
    onnx_url = f"{BASE_URL}/en/{locale}/{name}/{quality}/{voice_filename}.onnx"
    json_url = f"{BASE_URL}/en/{locale}/{name}/{quality}/{voice_filename}.onnx.json"
    return onnx_url, json_url


def get_voice_paths(voice_dir: str, locale: str, name: str, quality: str) -> Tuple[str, str]:
    """Get the local file paths for a voice's onnx and json files"""
    voice_filename = get_voice_filename(locale, name, quality)
    onnx_path = os.path.join(voice_dir, f"{voice_filename}.onnx")
    json_path = os.path.join(voice_dir, f"{voice_filename}.onnx.json")
    return onnx_path, json_path


def is_voice_downloaded(voice_dir: str, locale: str, name: str, quality: str) -> bool:
    """Check if a voice is already downloaded"""
    onnx_path, json_path = get_voice_paths(voice_dir, locale, name, quality)
    return os.path.exists(onnx_path) and os.path.exists(json_path)


def get_size_estimate(quality: str) -> int:
    """Get estimated download size in MB for a quality level"""
    return SIZE_ESTIMATES.get(quality, 17)


def download_file(url: str, dest_path: str, 
                  progress_callback: Optional[Callable[[int, int], None]] = None) -> bool:
    """
    Download a file with optional progress callback.
    
    Args:
        url: URL to download
        dest_path: Local path to save file
        progress_callback: Optional callback(bytes_downloaded, total_bytes)
        
    Returns:
        True on success, False on failure
    """
    try:
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        
        request = urllib.request.Request(url)
        request.add_header('User-Agent', 'PADLE Voice Downloader/1.0')
        
        with urllib.request.urlopen(request, timeout=60) as response:
            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            block_size = 8192
            
            with open(dest_path, 'wb') as f:
                while True:
                    chunk = response.read(block_size)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback:
                        progress_callback(downloaded, total_size)
        
        return True
        
    except (urllib.error.URLError, urllib.error.HTTPError, OSError, TimeoutError) as e:
        # Clean up partial download
        if os.path.exists(dest_path):
            try:
                os.unlink(dest_path)
            except OSError:
                pass
        return False


def download_voice(voice_dir: str, locale: str, name: str, quality: str,
                   progress_callback: Optional[Callable[[str, int, int], None]] = None) -> bool:
    """
    Download a single voice (onnx + json files).
    
    Args:
        voice_dir: Directory to save voice files
        locale: Voice locale (e.g., "en_US")
        name: Voice name (e.g., "amy")
        quality: Voice quality (e.g., "medium")
        progress_callback: Optional callback(filename, bytes_downloaded, total_bytes)
        
    Returns:
        True on success, False on failure
    """
    onnx_url, json_url = get_voice_urls(locale, name, quality)
    onnx_path, json_path = get_voice_paths(voice_dir, locale, name, quality)
    
    # Skip if already downloaded
    if os.path.exists(onnx_path) and os.path.exists(json_path):
        return True
    
    # Download ONNX file (the large one)
    onnx_filename = os.path.basename(onnx_path)
    
    def onnx_progress(downloaded, total):
        if progress_callback:
            progress_callback(onnx_filename, downloaded, total)
    
    if not os.path.exists(onnx_path):
        if not download_file(onnx_url, onnx_path, onnx_progress):
            return False
    
    # Download JSON config (small)
    if not os.path.exists(json_path):
        if not download_file(json_url, json_path):
            # Clean up onnx if json fails
            if os.path.exists(onnx_path):
                try:
                    os.unlink(onnx_path)
                except OSError:
                    pass
            return False
    
    return True


def get_available_qualities(voices: list = None, locale_filter: str = None) -> set:
    """
    Get the set of available quality levels.
    
    Args:
        voices: List of voice tuples (default: ENGLISH_VOICES)
        locale_filter: Optional locale to filter by (e.g., "en_US")
        
    Returns:
        Set of quality strings (e.g., {"low", "medium", "high"})
    """
    voices = voices or ENGLISH_VOICES
    qualities = set()
    
    for locale, name, available_qualities in voices:
        if locale_filter and locale != locale_filter:
            continue
        qualities.update(available_qualities)
    
    return qualities


def get_display_name(locale: str, name: str, quality: str) -> str:
    """Get a human-readable display name for a voice"""
    locale_map = {
        "en_US": "US",
        "en_GB": "UK",
    }
    locale_short = locale_map.get(locale, locale)
    name_title = name.replace("_", " ").title()
    return f"{name_title} ({locale_short}, {quality.title()})"


# =============================================================================
# COMMAND LINE INTERFACE
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description="Download Piper TTS English voices")
    parser.add_argument("--dir", default=os.path.expanduser("~/piper-voices"),
                        help="Directory to save voices (default: ~/piper-voices)")
    parser.add_argument("--us-only", action="store_true",
                        help="Download only US English voices")
    parser.add_argument("--gb-only", action="store_true",
                        help="Download only UK English voices")
    parser.add_argument("--quality", choices=["low", "medium", "high"],
                        help="Download only specific quality (default: all)")
    parser.add_argument("--list", action="store_true",
                        help="List available voices without downloading")
    
    args = parser.parse_args()
    
    # Filter voices based on arguments
    voices_to_download = []
    
    for locale, name, qualities in ENGLISH_VOICES:
        # Filter by locale
        if args.us_only and not locale.startswith("en_US"):
            continue
        if args.gb_only and not locale.startswith("en_GB"):
            continue
        
        # Filter by quality
        for quality in qualities:
            if args.quality and quality != args.quality:
                continue
            voices_to_download.append((locale, name, quality))
    
    # List mode
    if args.list:
        print("Available English Piper voices:\n")
        current_locale = None
        for locale, name, quality in voices_to_download:
            if locale != current_locale:
                print(f"\n{locale}:")
                current_locale = locale
            print(f"  {name} ({quality})")
        print(f"\nTotal: {len(voices_to_download)} voice models")
        return
    
    # Download mode
    print(f"Downloading {len(voices_to_download)} voice models to {args.dir}\n")
    
    success_count = 0
    fail_count = 0
    
    for locale, name, quality in voices_to_download:
        print(f"\n[{locale}] {name} ({quality})")
        
        def show_progress(filename, downloaded, total):
            if total > 0:
                pct = (downloaded / total) * 100
                mb_down = downloaded / (1024 * 1024)
                mb_total = total / (1024 * 1024)
                print(f"\r  {filename}: {mb_down:.1f}/{mb_total:.1f} MB ({pct:.0f}%)", end="", flush=True)
        
        if download_voice(args.dir, locale, name, quality, show_progress):
            print()  # Newline after progress
            success_count += 1
        else:
            print("\n  Failed!")
            fail_count += 1
    
    print(f"\n{'='*50}")
    print(f"Download complete!")
    print(f"  Success: {success_count}")
    print(f"  Failed:  {fail_count}")
    print(f"\nVoices saved to: {args.dir}")


if __name__ == "__main__":
    main()