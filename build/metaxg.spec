# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules

hiddenimports = []
hiddenimports += collect_submodules("playwright")
hiddenimports += collect_submodules("win32com")

block_cipher = None

analysis = Analysis(
    ["main.py", "main_debug.py"],
    pathex=["."],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(analysis.pure, analysis.zipped_data, cipher=block_cipher)

exe_normal = EXE(
    pyz,
    analysis.scripts[0],
    analysis.binaries,
    analysis.zipfiles,
    analysis.datas,
    [],
    name="MetaXg",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
)

exe_debug = EXE(
    pyz,
    analysis.scripts[1],
    analysis.binaries,
    analysis.zipfiles,
    analysis.datas,
    [],
    name="MetaXg_debug",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
)
