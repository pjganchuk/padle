# PADLE Refactoring Guide for Windows Compatibility

## Summary of Issues Found

After reviewing the codebase, I found several areas that need modification for Windows compatibility:

### 1. **tts.py - Linux-Specific Paths** ❌ → ✅
**Problem:** 
- Hardcoded Linux paths like `~/Projects/leadr/venv/bin/piper`
- Uses `bin` instead of `Scripts` (Windows venv convention)
- Missing `.exe` extension on Windows

**Solution:** See updated `tts.py` with cross-platform path handling.

### 2. **app.py - Voice Preview Audio Playback** ❌ → ✅
**Problem:** 
- Lines 769-776 use Linux-only commands (`aplay`, `paplay`)
- No Windows audio playback support

**Solution:** Replace with cross-platform `play_audio_file()` function.

### 3. **config.py - Voices Directory Path** ⚠️ → ✅
**Problem:**
- Uses `~/piper-voices` which works but isn't Windows-conventional

**Solution:** Use `%LOCALAPPDATA%\piper-voices` on Windows.

### 4. **requirements.txt - Missing Dependencies** ⚠️ → ✅
**Problem:**
- Missing `pydub` and `screeninfo`
- No documentation of system dependencies

**Solution:** Updated requirements.txt with all dependencies and Windows install notes.

---

## Files to Update

### 1. Replace `tts.py` entirely
Use the new version from `padle_refactored/tts.py`

### 2. Replace `config.py` entirely  
Use the new version from `padle_refactored/config.py`

### 3. Add new file `platform_utils.py`
Copy from `padle_refactored/platform_utils.py`

### 4. Update `app.py` - Make these specific changes:

#### Change 1: Add import at top of file
```python
# Add after other imports (around line 36)
from platform_utils import play_audio_file
```

#### Change 2: Update VoiceSelectionDialog.preview_voice() method
Replace lines 758-791 with:

```python
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
            import os
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
```

### 5. Replace `requirements.txt`
Use the new version from `padle_refactored/requirements.txt`

---

## Windows Bundling with PyInstaller

### Step 1: Install PyInstaller
```bash
pip install pyinstaller
```

### Step 2: Create spec file
Create `padle.spec`:

```python
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('icon.png', '.'),
    ],
    hiddenimports=[
        'PIL._tkinter_finder',
        'moondream',
        'cv2',
        'numpy',
        'pytesseract',
        'pydub',
        'screeninfo',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PADLE',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to True for debugging
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',  # Need to convert PNG to ICO
)
```

### Step 3: Bundle
```bash
pyinstaller padle.spec
```

### Step 4: Create installer (optional)
Use NSIS or Inno Setup to create a proper Windows installer that:
- Installs PADLE
- Optionally installs/configures Tesseract
- Prompts user to install VLC and FFmpeg if not found
- Sets up default voices directory

---

## Testing Checklist

Before release, test on Windows:

- [ ] App launches without errors
- [ ] Moondream connection works
- [ ] Video loading works
- [ ] Video playback with audio works
- [ ] AI description generation works
- [ ] OCR mode works (requires Tesseract)
- [ ] Voice preview playback works
- [ ] Audio track export works (requires FFmpeg)
- [ ] WebVTT export works
- [ ] Project save/load works
- [ ] Auto-save works
- [ ] Window positioning on multi-monitor works

---

## Known Windows Considerations

1. **VLC Installation**: Users must install VLC separately. The python-vlc package will find it automatically if installed to default location.

2. **Tesseract Path**: May need to add Tesseract to PATH or configure pytesseract:
   ```python
   pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
   ```

3. **FFmpeg for MP3 Export**: pydub requires FFmpeg. Either add to PATH or set:
   ```python
   from pydub import AudioSegment
   AudioSegment.converter = r"C:\path\to\ffmpeg.exe"
   ```

4. **Console Window**: PyInstaller's `console=False` hides the console but may make debugging harder. Consider adding a log file.

5. **Icon File**: Windows .exe needs .ico format. Convert icon.png using:
   ```bash
   # Using ImageMagick
   convert icon.png -define icon:auto-resize=256,128,64,48,32,16 icon.ico
   ```
