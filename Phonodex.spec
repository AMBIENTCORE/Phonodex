# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('assets/I-pixel-u.ttf', 'assets'), ('assets/phonodex.ico', 'assets'), ('assets/no_cover.png', 'assets'), ('assets/no_cover_2.png', 'assets'), ('assets/no_cover_slime.png', 'assets'), ('assets/no_cover_yeoldetagger.png', 'assets'), ('assets/no_cover_amberterminal.png', 'assets'), ('assets/no_cover_sakura.png', 'assets'), ('assets/no_cover_eink.png', 'assets')]
binaries = []
hiddenimports = []
tmp_ret = collect_all('tkinterdnd2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Phonodex',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['assets\\phonodex.ico'],
)
