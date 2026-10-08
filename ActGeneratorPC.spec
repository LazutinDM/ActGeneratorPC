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

# COLLEAGUE EDIT POINT: Windows executable name, icon and Data runtime folder.
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
