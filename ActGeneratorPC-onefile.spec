# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller recipe with the Python and Qt runtime embedded in one EXE."""

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
    runtime_hooks=[str(project_dir / "runtime_dll_search.py")],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

# COLLEAGUE EDIT POINT: a.binaries and a.datas are intentionally included in
# the executable.  This prevents a new EXE from loading stale Qt DLLs left by
# an older portable folder.
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="ActGeneratorPC",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=str(icon_path) if icon_path.exists() else None,
)
