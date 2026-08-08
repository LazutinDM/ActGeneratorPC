# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

project_dir = Path(SPEC).resolve().parent
icon_path = project_dir / "Data" / "Icons" / "app.ico"

a = Analysis(
    ["app.py"],
    pathex=[str(project_dir)],
    binaries=[],
    datas=[],
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
    [],
    exclude_binaries=True,
    name="ActGeneratorPC",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    # Keep the PyInstaller runtime beside application data in one Data folder.
    # PyInstaller supports exactly one contents-directory level in onedir mode.
    contents_directory="Data",
    icon=str(icon_path) if icon_path.exists() else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    name="ActGeneratorPC",
)
