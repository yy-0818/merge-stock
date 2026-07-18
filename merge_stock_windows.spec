# -*- mode: python ; coding: utf-8 -*-
"""
Windows 打包专用 PyInstaller spec 文件。
GitHub Actions 云端打包时使用此文件：
    pyinstaller --noconfirm merge_stock_windows.spec
生成产物：  dist/merge_stock/merge_stock.exe  (文件夹形式，需随 DLL 分发)

注意：Windows 下 PySide6 GUI 程序正常会显示控制台窗口，
      这是 Windows 上 PySide6 的标准行为，不影响 GUI 功能。
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
    console=True,          # Windows: 显示控制台窗口(PySide6 标准行为)
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
    name='merge_stock',     # 输出到 dist/merge_stock/ 文件夹
)
