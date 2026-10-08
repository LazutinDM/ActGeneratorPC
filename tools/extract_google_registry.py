"""Build the portable factory-number registry from the approved Google export.

The source workbook is read-only.  Output is written to stdout so the caller
can review and install it explicitly.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from equipment_registry_update import extract_registry


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: extract_google_registry.py <exported-workbook.xlsx>")
    project_dir = Path(__file__).resolve().parents[1]
    stations = (project_dir / "Data" / "Variables" / "location_list.txt").read_text(
        encoding="utf-8-sig"
    ).splitlines()
    registry = extract_registry(Path(sys.argv[1]), stations)
    print(json.dumps(registry, ensure_ascii=False, indent=2))
