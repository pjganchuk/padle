# -*- mode: python ; coding: utf-8 -*-
"""
PADLE PyInstaller Spec File
Panopto Audio Descriptions List Editor

Build with: pyinstaller padle.spec

This creates a single-folder distribution with all dependencies.
"""

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, collect_submodules

block_cipher = None

# Collect data files for transformers and other packages
datas = [
    ('icon.png', '.'),
    ('icon.ico', '.'),
]

# Collect transformers tokenizer files
datas += collect_data_files('transformers', include_py_files=False)

# Collect jaraco namespace package files (required by pkg_resources)
try:
    datas += collect_data_files('jaraco.text')
    datas += collect_data_files('jaraco.functools')
    datas += collect_data_files('jaraco.context')
    datas += collect_data_files('jaraco.classes')
except Exception:
    pass

# Hidden imports required by transformers and moondream
hiddenimports = [
    # Transformers
    'transformers',
    'transformers.models.phi',
    'transformers.models.auto',
    'transformers.tokenization_utils_base',
    'transformers.tokenization_utils_fast',
    
    # PyTorch
    'torch',
    'torch.cuda',
    'torchvision',
    'torchvision.transforms',
    
    # Other AI dependencies
    'accelerate',
    'safetensors',
    'safetensors.torch',
    'huggingface_hub',
    'einops',
    
    # Image processing
    'PIL',
    'PIL._tkinter_finder',
    'cv2',
    'numpy',
    
    # OCR
    'pytesseract',
    
    # Audio
    'pydub',
    'vlc',
    
    # TTS
    'piper',
    
    # UI
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
    
    # System
    'screeninfo',
    'logging',
    'json',
    'threading',
    'subprocess',
    
    # pkg_resources / setuptools dependencies (jaraco namespace packages)
    'pkg_resources',
    'jaraco',
    'jaraco.text',
    'jaraco.functools',
    'jaraco.context',
    'jaraco.classes',
]

# Collect all jaraco submodules
try:
    hiddenimports += collect_submodules('jaraco')
except Exception:
    pass

# Binaries to include
binaries = []

# Collect CUDA libraries if available
try:
    import torch
    if torch.cuda.is_available():
        cuda_libs = collect_dynamic_libs('torch')
        binaries += cuda_libs
except ImportError:
    pass

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'pandas',
        'notebook',
        'jupyter',
        'IPython',
        'pytest',
        'pip',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PADLE',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
    version='version_info.txt',  # Optional version info
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PADLE',
)
