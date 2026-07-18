# -*- mode: python ; coding: utf-8 -*-
"""
项目根目录下的 PyInstaller spec 文件。
从项目根目录运行：  pyinstaller --noconfirm merge_stock.spec
生成产物：  dist/merge_stock.app/    (macOS)
           dist/merge_stock/merge_stock.exe  (Windows)
"""
from pathlib import Path

import openpyxl
from PyInstaller.utils.hooks import collect_submodules

PROJECT_DIR = Path(SPECPATH).resolve()

hiddenimports = ['openpyxl']
hiddenimports += collect_submodules('PySide6')
hiddenimports += collect_submodules('openpyxl')


a = Analysis(
    ['gui.py'],
    pathex=[str(PROJECT_DIR)],
    binaries=[],
    datas=[],
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
    [],
    exclude_binaries=True,
    name='merge_stock',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='merge_stock',
)

app = BUNDLE(
    coll,
    name='merge_stock.app',
    icon=None,
    bundle_identifier='com.stock.merge',
)