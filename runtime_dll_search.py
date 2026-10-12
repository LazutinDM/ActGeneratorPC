"""Prefer DLLs unpacked by PyInstaller over libraries installed elsewhere.

PyInstaller executes this runtime hook before importing the application.  The
explicit directories prevent Qt from resolving an incompatible copy left by a
different program or an older ActGeneratorPC portable directory.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path


_DLL_DIRECTORY_HANDLES: list[object] = []


def configure_embedded_dll_search() -> None:
    """Register the one-file extraction directories with Windows' DLL loader."""
    if os.name != "nt" or not getattr(sys, "frozen", False):
        return

    bundle_root = Path(getattr(sys, "_MEIPASS", "")).resolve()
    for directory in (
        bundle_root,
        bundle_root / "PySide6",
        bundle_root / "shiboken6",
    ):
        if directory.is_dir():
            # Keep each handle alive for the lifetime of the process. Closing a
            # handle immediately removes its directory from the loader search.
            _DLL_DIRECTORY_HANDLES.append(os.add_dll_directory(str(directory)))


configure_embedded_dll_search()
