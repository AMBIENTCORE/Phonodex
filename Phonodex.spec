# -*- mode: python ; coding: utf-8 -*-

import os
import sys
import site
from pathlib import Path

block_cipher = None

# Define the base directory for assets
assets_dir = 'assets'

# Find tkinterdnd2 directory in site-packages
site_packages = site.getsitepackages()
tkdnd_path = None

for site_dir in site_packages:
    potential_path = os.path.join(site_dir, 'tkinterdnd2')
    if os.path.exists(potential_path):
        tkdnd_path = potential_path
        break

if not tkdnd_path:
    raise FileNotFoundError("tkinterdnd2 package not found in site-packages")

# List all assets that need to be included
assets = [
    # The font file
    (os.path.join(assets_dir, 'I-pixel-u.ttf'), os.path.join('assets')),
    
    # The no_cover image
    (os.path.join(assets_dir, 'no_cover.png'), os.path.join('assets')),
]

# Add tkdnd TCL files - this is critical
tkdnd_files = []
for root, dirs, files in os.walk(os.path.join(tkdnd_path, 'tkdnd')):
    for file in files:
        source = os.path.join(root, file)
        # Preserve the directory structure under tkdnd
        dest_dir = os.path.join('tkinterdnd2', os.path.relpath(root, tkdnd_path))
        tkdnd_files.append((source, dest_dir))

# Define the entry point and other files to include
a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=assets + tkdnd_files,  # Include our assets and tkdnd files
    hiddenimports=[
        'tkinterdnd2',
        'mutagen',
        'mutagen.easyid3',
        'mutagen.mp3',
        'mutagen.flac',
        'mutagen.mp4',
        'mutagen.oggvorbis',
        'mutagen.asf',
        'mutagen.wave',
        'mutagen.id3',
        'requests',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'config',
        'utils.logging',
        'utils.file_operations',
        'utils.image_handling',
        'utils.metadata',
        'utils.table_operations',
        'services.api_client',
        'ui.dialogs',
        'ui.styles',
        'hashlib',
        'threading',
        'collections',
        'array',
        'tkinter',
        'tkinter.ttk',
        'tkinter.font',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'platform',
        'subprocess',
        'shutil',
        'os',
        'io',
        'json',
        're'
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

# Add TCL/TK DLLs without modifying a.binaries directly
# Each entry must be a tuple of (dest_name, source_name, typecode)
tcl_tk_dirs = [
    os.path.join(sys.base_prefix, 'DLLs'),
    os.path.join(sys.base_prefix, 'tcl', 'bin'),
    os.path.join(sys.base_prefix, 'Lib', 'lib-tk')
]

pyz = PYZ(
    a.pure, 
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Phonodex',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Temporarily true for debugging
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(assets_dir, 'no_cover.png'),
) 