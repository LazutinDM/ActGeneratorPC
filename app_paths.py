"""Portable paths used by ActGeneratorPC.

This module contains no interface code.  Edit it when the folder layout of the
portable package changes.  Keeping all paths here prevents the main window
from depending on scattered string literals.
"""

from __future__ import annotations

import os
import sys


def resource_path(relative_path: str) -> str:
    """Return a bundled PyInstaller resource or a development-time file."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative_path)


# COLLEAGUE EDIT POINT: portable directory layout. Keep user data under Data or
# Acts so the updater can preserve it safely.
# APP_DIR is always the portable folder: beside app.py during development and
# beside ActGeneratorPC.exe in a packaged build.
APP_DIR = (
    os.path.dirname(os.path.abspath(sys.executable))
    if getattr(sys, "frozen", False)
    else os.path.dirname(os.path.abspath(__file__))
)
DATA_DIR = os.path.join(APP_DIR, "Data")
TEMPLATES_DIR = os.path.join(DATA_DIR, "Templates")
VARIABLES_DIR = os.path.join(DATA_DIR, "Variables")
ICONS_DIR = os.path.join(DATA_DIR, "Icons")
OTHER_DIR = os.path.join(DATA_DIR, "Other")
CONFIG_DIR = os.path.join(OTHER_DIR, "Configuration")
ICONS_GEN_DIR = os.path.join(OTHER_DIR, "GeneratedIcons")
DOCUMENTATION_DIR = os.path.join(OTHER_DIR, "Documentation")
HISTORY_DIR = os.path.join(OTHER_DIR, "History")
EXCEL_WORK_DIR = os.path.join(OTHER_DIR, "Excel")
OUTPUT_DIR = os.path.join(APP_DIR, "Acts")
LEGACY_OUTPUT_DIR = os.path.join(APP_DIR, "Act_Ready")
# These are old user-data folders that may be removed after migration.
# Never add ``_internal``: PyInstaller stores the embedded Python runtime there.
LEGACY_ROOT_FOLDERS = ("Templates", "icons", "_icons", "variables")
EXCEL_EXPORT_CONFIG = os.path.join(CONFIG_DIR, "excel_export.json")
EXCEL_TEMPLATE = os.path.join(TEMPLATES_DIR, "Tables.xlsx")
EMBEDDED_EXCEL_WORKBOOK = os.path.join(EXCEL_WORK_DIR, "Tables.xlsx")
APP_SETTINGS_CONFIG = os.path.join(CONFIG_DIR, "app_settings.json")
EQUIPMENT_REGISTRY = os.path.join(VARIABLES_DIR, "station_serials.json")

# COLLEAGUE EDIT POINT: user interface icon filenames.
# Missing optional icons are handled by app.py.
ICON_APP = os.path.join(ICONS_DIR, "app.ico")
ICON_BRAND = os.path.join(ICONS_DIR, "brand.png")
ICON_MODES = os.path.join(ICONS_DIR, "modes.png")
ICON_ARROW_DOWN = os.path.join(ICONS_DIR, "arrow_down.png")
ICON_CALENDAR = os.path.join(ICONS_DIR, "calendar.png")
ICON_DONE_WORK = os.path.join(ICONS_DIR, "done_work.png")
ICON_EXCEL = os.path.join(ICONS_DIR, "excel.png")
ICON_ISSUES = os.path.join(ICONS_DIR, "issues.png")
ICON_MATERIALS = os.path.join(ICONS_DIR, "materials.png")
ICON_QTY = os.path.join(ICONS_DIR, "qty.png")
ICON_SERIAL = os.path.join(ICONS_DIR, "serial.png")
ICON_SERVICES = os.path.join(ICONS_DIR, "services.png")
ICON_STATION = os.path.join(ICONS_DIR, "station.png")
ICON_USER = os.path.join(ICONS_DIR, "user.png")
ICON_XLS = os.path.join(ICONS_DIR, "xls.png")
ICON_EQUIPMENT = os.path.join(ICONS_DIR, "equipment.png")
ICON_LISTS = os.path.join(ICONS_DIR, "lists.png")
ICON_DOCX = os.path.join(ICONS_DIR, "docx.png")

SPIN_UP_LIGHT_PNG = os.path.join(ICONS_GEN_DIR, "spin_up_light.png")
SPIN_DOWN_LIGHT_PNG = os.path.join(ICONS_GEN_DIR, "spin_down_light.png")
SPIN_UP_DARK_PNG = os.path.join(ICONS_GEN_DIR, "spin_up_dark.png")
SPIN_DOWN_DARK_PNG = os.path.join(ICONS_GEN_DIR, "spin_down_dark.png")

# COLLEAGUE EDIT POINT: editable text-list filenames used by the form.
FILES = {
    "executors": os.path.join(VARIABLES_DIR, "executor_list.txt"),
    "locations": os.path.join(VARIABLES_DIR, "location_list.txt"),
    "models": os.path.join(VARIABLES_DIR, "model_list.txt"),
    "work": os.path.join(VARIABLES_DIR, "work_list.txt"),
    "issues": os.path.join(VARIABLES_DIR, "issues_list.txt"),
    "done_work": os.path.join(VARIABLES_DIR, "done_work_list.txt"),
    "materials": os.path.join(VARIABLES_DIR, "materials_list.txt"),
    "links": os.path.join(CONFIG_DIR, "tables.json"),
}

# COLLEAGUE EDIT POINT: DOCX templates selected by equipment type.
TEMPLATE_TYPE1 = os.path.join(TEMPLATES_DIR, "ABP_MKTF.docx")
TEMPLATE_TYPE2 = os.path.join(TEMPLATES_DIR, "Validator_MID.docx")
