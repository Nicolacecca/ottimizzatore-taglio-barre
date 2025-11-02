# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['c:\\Users\\Utente\\Desktop\\PROGETTI\\TAGLIO\\ottimizzatore_taglio.py'],
    pathex=[],
    binaries=[],
    datas=[('icon.ico', '.'), ('icon.png', '.'), ('help_icon.ico', '.'), ('help_icon.png', '.')],
    hiddenimports=[],
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
    name='OttimizzatoreTaglioBarre',
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
    icon=['c:\\Users\\Utente\\Desktop\\PROGETTI\\TAGLIO\\icon.ico'],
)
