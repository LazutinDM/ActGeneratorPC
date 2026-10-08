"""ActGeneratorPC user interface and document-generation workflow.

Folder paths live in app_paths.py; update/rollback code lives in
update_manager.py.  Keep this file focused on UI actions and DOCX generation.
"""

import os
import re
import sys
import json
import datetime
import subprocess
import ctypes
import ctypes.wintypes
import shutil
import tempfile
from typing import List, Dict
from functools import partial

from update_manager import UpdateManager, confirm_healthy_startup
from PySide6.QtCore import (
    Qt, QUrl, QSignalBlocker, QStringListModel, QEvent, QTimer,
    QByteArray, QBuffer, QIODevice, QThread, QDate, QSize, QPoint,
)
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtGui import (
    QAction, QPalette, QColor, QDesktopServices, QPixmap, QPainter, QPen, QIcon,
    QPainterPath, QRegion, QFont
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QMessageBox, QDialog, QListWidget, QListWidgetItem,
    QAbstractItemView, QAbstractSpinBox, QInputDialog, QSpinBox, QFormLayout,
    QStyle, QComboBox, QCompleter, QToolButton, QMenu, QFrame,
    QGraphicsDropShadowEffect, QFileDialog, QScrollArea, QTableWidget,
    QTableWidgetItem, QHeaderView, QProgressBar, QDateEdit, QCalendarWidget,
    QWidgetAction, QSizePolicy,
)

from docx import Document
from docx_preview import PreviewRenderError, render_docx_to_pdf
from excel_export import (
    ExcelExportError,
    append_act_to_workbook,
    area_for_station,
    ensure_embedded_excel_workbook,
    load_excel_path,
    note_for_materials,
    resolve_excel_workbook,
    save_excel_path,
)
from history_store import clear_history_records, load_history_records, save_history_record
from equipment_registry_update import (
    RegistryUpdateError,
    allowed_stations_through_alabushevo,
    canonical_station,
    update_registry,
)

from app_paths import (
    APP_DIR, APP_SETTINGS_CONFIG, CONFIG_DIR, DATA_DIR, DOCUMENTATION_DIR,
    EMBEDDED_EXCEL_WORKBOOK, EXCEL_EXPORT_CONFIG, EXCEL_TEMPLATE,
    EQUIPMENT_REGISTRY, EXCEL_WORK_DIR, FILES, ICONS_DIR,
    ICONS_GEN_DIR, ICON_APP, ICON_BRAND, ICON_ARROW_DOWN,
    ICON_CALENDAR, ICON_DOCX,
    ICON_DONE_WORK, ICON_EQUIPMENT, ICON_EXCEL, ICON_ISSUES, ICON_LISTS,
    ICON_MATERIALS, ICON_QTY, ICON_SERIAL, ICON_SERVICES, ICON_STATION,
    ICON_USER, ICON_XLS, LEGACY_OUTPUT_DIR, LEGACY_ROOT_FOLDERS, OTHER_DIR, OUTPUT_DIR,
    SPIN_DOWN_DARK_PNG, SPIN_DOWN_LIGHT_PNG, SPIN_UP_DARK_PNG,
    SPIN_UP_LIGHT_PNG, TEMPLATE_TYPE1, TEMPLATE_TYPE2, TEMPLATES_DIR,
    VARIABLES_DIR, HISTORY_DIR,
)


INVALID_FILENAME_RE = re.compile(r'[\\/:*?"<>|]+')
URL_SCHEME_RE = re.compile(r"^[a-zA-Z][a-zA-Z0-9+\-.]*://")
SEARCH_CLEAN_RE = re.compile(r"[^0-9a-zа-я]+")
# COLLEAGUE EDIT POINT: add normalized executor names here when they must be
# removed from existing lists and rejected from manual entry.
EXCLUDED_EXECUTOR_KEYS = {"чепельаа"}
# COLLEAGUE EDIT POINT: service period format used only by the MID/Validator act.
SERVICE_TIME_FORMAT = "%H:%M"
ACT_DATE_FORMAT = "%d.%m.%Y"
# Users enter only four digits; Qt keeps the time separator in place.
SERVICE_TIME_INPUT_MASK = "00:00;_"
SERVICE_MONTHS_GENITIVE = (
    "января", "февраля", "марта", "апреля", "мая", "июня",
    "июля", "августа", "сентября", "октября", "ноября", "декабря",
)
SERVICE_PERIOD_BLANK = "_______часов «__»____202__года"

# COLLEAGUE EDIT POINT: names in the source equipment table are not always the
# same as the values shown in the application's "Тип оборудования" list.
# Keys are accepted source spellings; values must exactly match model_list.txt.
EQUIPMENT_MODEL_ALIASES = {
    "АБП-БН": "АБП-БН",
    "АПБ-БН": "АБП-БН",
    "АБП": "АБП-09",
    "АПБ": "АБП-09",
    "АБП-09": "АБП-09",
    "АПБ-09": "АБП-09",
    "АБП-09-М3": "АБП-09-М3",
    "АПБ-09-М3": "АБП-09-М3",
    "АН-15": "АН-15",
    "АН 15": "АН-15",
    "БПА-20-БН1": "БПА-20-БН1",
    "МКТФ": "МКТФ",
    "ВАЛИДАТОР ВП-FT": "Валидатор ВП-FT",
    "МИД": "МИД",
}


# -----------------------------
# Helpers
# -----------------------------
def ensure_dirs():
    """Create required portable folders and migrate supported legacy data."""
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(TEMPLATES_DIR, exist_ok=True)
    os.makedirs(VARIABLES_DIR, exist_ok=True)
    os.makedirs(OTHER_DIR, exist_ok=True)
    os.makedirs(CONFIG_DIR, exist_ok=True)
    os.makedirs(DOCUMENTATION_DIR, exist_ok=True)
    os.makedirs(HISTORY_DIR, exist_ok=True)
    os.makedirs(EXCEL_WORK_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(ICONS_GEN_DIR, exist_ok=True)
    os.makedirs(ICONS_DIR, exist_ok=True)
    migrate_legacy_portable_layout()
    remove_excluded_executors_from_file()


def normalize_serial_key(value: str) -> str:
    """Normalize a factory number for exact registry lookup."""
    return re.sub(r"\s+", "", str(value or "")).casefold()


def normalize_registry_text(value: str) -> str:
    """Normalize visible registry text without changing its wording."""
    return re.sub(r"\s+", " ", str(value or "").replace("\xa0", " ")).strip()


def executor_name_with_initials(value: str) -> str:
    """Return a safe ``Surname I. O.`` value without digits/foreign signs."""
    normalized = normalize_registry_text(value)
    tokens = normalized.split()
    for index, token in enumerate(tokens):
        letters_only = "".join(character for character in token if character.isalpha())
        if letters_only:
            # Initials may be written as "Д. М." or "Д.М.". Extract the
            # remaining letter groups and normalize at most two initials.
            tail = " ".join(tokens[index + 1:])
            initial_groups = re.findall(r"[^\W\d_]+", tail, flags=re.UNICODE)
            initials = [f"{group[0].upper()}." for group in initial_groups[:2]]
            return " ".join([letters_only, *initials])
    return ""


def canonical_equipment_model(value: str) -> str:
    """Translate a source-table equipment label to the application's label."""
    normalized = normalize_registry_text(value).replace("–", "-").replace("—", "-")
    return EQUIPMENT_MODEL_ALIASES.get(normalized.upper(), normalized)


def load_equipment_registry(path: str) -> Dict[str, Dict[str, str]]:
    """Load unambiguous factory-number bindings from the portable JSON file."""
    try:
        with open(path, "r", encoding="utf-8-sig") as registry_file:
            raw_records = json.load(registry_file)
    except (OSError, UnicodeError, json.JSONDecodeError):
        return {}

    if not isinstance(raw_records, list):
        return {}

    records: Dict[str, Dict[str, str]] = {}
    ambiguous_keys = set()
    for raw_record in raw_records:
        if not isinstance(raw_record, dict):
            continue
        record = {
            "serial": normalize_registry_text(raw_record.get("serial", "")),
            "model": canonical_equipment_model(raw_record.get("model", "")),
            "station": canonical_station(raw_record.get("station", "")),
            "equipmentId": normalize_registry_text(
                raw_record.get("equipmentId", "")
            ),
            "stationNumber": normalize_registry_text(
                raw_record.get("stationNumber", "")
            ),
            "sourceSheet": normalize_registry_text(
                raw_record.get("sourceSheet", "")
            ),
        }
        key = normalize_serial_key(record["serial"])
        if not key or not record["model"] or not record["station"]:
            continue
        existing = records.get(key)
        if existing and (
            existing["model"].casefold(), existing["station"].casefold()
        ) != (record["model"].casefold(), record["station"].casefold()):
            ambiguous_keys.add(key)
            continue
        records[key] = record

    for key in ambiguous_keys:
        records.pop(key, None)
    return records


def is_validator_template(path: str) -> bool:
    """Handle is validator template."""
    return os.path.basename(str(path)).casefold() == os.path.basename(TEMPLATE_TYPE2).casefold()


def normalize_service_time_input(value: str) -> str:
    """Return entered time and accept legacy history values with an old date."""
    text = (value or "").strip()
    if not any(character.isdigit() for character in text):
        return ""
    legacy_match = re.match(r"^(\d{2}:\d{2})(?:\s+\d{2}\.\d{2}\.\d{4})?$", text)
    return legacy_match.group(1) if legacy_match else text


def parse_service_datetime(value: str, act_date: str) -> datetime.datetime:
    """Parse service datetime for this workflow."""
    time_text = normalize_service_time_input(value)
    date_text = (act_date or "").strip()
    try:
        parsed_time = datetime.datetime.strptime(time_text, SERVICE_TIME_FORMAT).time()
    except ValueError as error:
        raise ValueError(
            "Время услуги укажите в формате ЧЧ:ММ, например 08:30."
        ) from error
    if parsed_time.strftime(SERVICE_TIME_FORMAT) != time_text:
        raise ValueError(
            "Время услуги укажите в формате ЧЧ:ММ, например 08:30."
        )
    try:
        parsed_date = datetime.datetime.strptime(date_text, ACT_DATE_FORMAT).date()
    except ValueError as error:
        raise ValueError(
            "Дату акта укажите в формате ДД.ММ.ГГГГ, например 17.08.2026."
        ) from error
    if parsed_date.strftime(ACT_DATE_FORMAT) != date_text:
        raise ValueError(
            "Дату акта укажите в формате ДД.ММ.ГГГГ, например 17.08.2026."
        )
    return datetime.datetime.combine(parsed_date, parsed_time)


def service_datetime_for_doc(value: str, act_date: str) -> str:
    """Handle service datetime for doc."""
    value = normalize_service_time_input(value)
    if not value:
        return SERVICE_PERIOD_BLANK
    parsed = parse_service_datetime(value, act_date)
    return (
        f"{parsed:%H:%M} часов «{parsed:%d}» "
        f"{SERVICE_MONTHS_GENITIVE[parsed.month - 1]} {parsed:%Y} года"
    )


def validate_service_period(
    start: str, end: str, act_date: str
) -> tuple[str, str]:
    """Validate service period for this workflow."""
    start = normalize_service_time_input(start)
    end = normalize_service_time_input(end)
    if not start and not end:
        return SERVICE_PERIOD_BLANK, SERVICE_PERIOD_BLANK
    if not start or not end:
        raise ValueError("Заполните и начало, и окончание периода оказания услуги.")
    start_value = parse_service_datetime(start, act_date)
    end_value = parse_service_datetime(end, act_date)
    if end_value < start_value:
        raise ValueError("Окончание периода услуги не может быть раньше начала.")
    return (
        service_datetime_for_doc(start, act_date),
        service_datetime_for_doc(end, act_date),
    )


def migrate_legacy_portable_layout():
    """Move user documents and remove obsolete root items after an update."""
    # Older versions stored variable lists directly in Data.
    for name in os.listdir(DATA_DIR):
        source = os.path.join(DATA_DIR, name)
        if not os.path.isfile(source) or not name.lower().endswith(".txt"):
            continue
        destination = os.path.join(VARIABLES_DIR, name)
        try:
            if os.path.exists(destination):
                os.remove(source)
            else:
                os.replace(source, destination)
        except OSError:
            if not os.path.exists(destination):
                shutil.copy2(source, destination)
                os.remove(source)

    legacy_tables = os.path.join(DATA_DIR, "tables.json")
    current_tables = os.path.join(CONFIG_DIR, "tables.json")
    if os.path.isfile(legacy_tables):
        try:
            if os.path.exists(current_tables):
                os.remove(legacy_tables)
            else:
                os.replace(legacy_tables, current_tables)
        except OSError:
            if not os.path.exists(current_tables):
                shutil.copy2(legacy_tables, current_tables)
                os.remove(legacy_tables)

    legacy_generated_icons = os.path.join(DATA_DIR, "_icons")
    if os.path.isdir(legacy_generated_icons):
        for name in os.listdir(legacy_generated_icons):
            source = os.path.join(legacy_generated_icons, name)
            destination = os.path.join(ICONS_GEN_DIR, name)
            if os.path.isfile(source) and not os.path.exists(destination):
                try:
                    os.replace(source, destination)
                except OSError:
                    shutil.copy2(source, destination)
                    os.remove(source)
        try:
            os.rmdir(legacy_generated_icons)
        except OSError:
            pass

    legacy_output_dirs = (
        LEGACY_OUTPUT_DIR,
        os.path.join(APP_DIR, "\u0410\u043a\u0442\u044b"),
    )
    for legacy_output_dir in legacy_output_dirs:
        if not os.path.isdir(legacy_output_dir):
            continue
        for name in os.listdir(legacy_output_dir):
            source = os.path.join(legacy_output_dir, name)
            destination = os.path.join(OUTPUT_DIR, name)
            if not os.path.exists(destination):
                try:
                    os.replace(source, destination)
                except OSError:
                    if os.path.isdir(source):
                        shutil.copytree(source, destination)
                        shutil.rmtree(source)
                    else:
                        shutil.copy2(source, destination)
                        os.remove(source)
        try:
            os.rmdir(legacy_output_dir)
        except OSError:
            pass

    if not getattr(sys, "frozen", False):
        return

    # Normalize the spelling of portable user-data folders. PyInstaller runtime
    # files also live in Data, so unknown entries must never be deleted here.
    expected_data_folders = {
        name.casefold(): name
        for name in ("Icons", "Other", "Templates", "Variables")
    }
    # Windows paths are case-insensitive, but Explorer keeps the spelling used
    # by an old release. Rename through a temporary name to normalize it.
    for actual_name in os.listdir(DATA_DIR):
        canonical_name = expected_data_folders.get(actual_name.casefold())
        if not canonical_name or actual_name == canonical_name:
            continue
        actual_path = os.path.join(DATA_DIR, actual_name)
        temporary_path = os.path.join(DATA_DIR, f".{canonical_name}.rename")
        canonical_path = os.path.join(DATA_DIR, canonical_name)
        try:
            os.replace(actual_path, temporary_path)
            os.replace(temporary_path, canonical_path)
        except OSError:
            pass

    for legacy_folder in LEGACY_ROOT_FOLDERS:
        legacy_path = os.path.join(APP_DIR, legacy_folder)
        if os.path.isdir(legacy_path):
            shutil.rmtree(legacy_path, ignore_errors=True)

    legacy_config = os.path.join(APP_DIR, "update_config.json")
    new_config = os.path.join(CONFIG_DIR, "update_config.json")
    if os.path.isfile(legacy_config) and os.path.isfile(new_config):
        try:
            os.remove(legacy_config)
        except OSError:
            pass


def _make_placeholder_icon_png(path: str, label: str, bg: QColor):
    """Create placeholder icon png for this workflow."""
    if os.path.exists(path):
        return
    pm = QPixmap(64, 64)
    pm.fill(Qt.transparent)

    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing, True)

    p.setPen(Qt.NoPen)
    p.setBrush(bg)
    p.drawRoundedRect(4, 4, 56, 56, 14, 14)

    p.setPen(QColor("#ffffff"))
    f = QFont()
    f.setBold(True)
    f.setPointSize(16)
    p.setFont(f)
    p.drawText(pm.rect(), Qt.AlignCenter, (label or "?")[:3].upper())

    p.end()
    pm.save(path, "PNG")


def ensure_default_icons():
    """Ensure default icons for this workflow."""
    _make_placeholder_icon_png(ICON_ARROW_DOWN, "V", QColor("#4B5563"))
    _make_placeholder_icon_png(ICON_CALENDAR, "CAL", QColor("#10B981"))
    _make_placeholder_icon_png(ICON_DONE_WORK, "OK", QColor("#22C55E"))
    _make_placeholder_icon_png(ICON_EXCEL, "XL", QColor("#16A34A"))
    _make_placeholder_icon_png(ICON_ISSUES, "!", QColor("#F97316"))
    _make_placeholder_icon_png(ICON_MATERIALS, "MAT", QColor("#0EA5E9"))
    _make_placeholder_icon_png(ICON_QTY, "QTY", QColor("#6366F1"))
    _make_placeholder_icon_png(ICON_SERIAL, "SN", QColor("#8B5CF6"))
    _make_placeholder_icon_png(ICON_SERVICES, "SRV", QColor("#06B6D4"))
    _make_placeholder_icon_png(ICON_STATION, "ST", QColor("#64748B"))
    _make_placeholder_icon_png(ICON_USER, "USR", QColor("#3B82F6"))
    _make_placeholder_icon_png(ICON_XLS, "XLS", QColor("#16A34A"))
    _make_placeholder_icon_png(ICON_EQUIPMENT, "EQ", QColor("#A855F7"))
    _make_placeholder_icon_png(ICON_LISTS, "LST", QColor("#14B8A6"))
    _make_placeholder_icon_png(ICON_DOCX, "DOC", QColor("#2563EB"))


def make_spin_arrow_png(path: str, color_hex: str, direction: str):
    # COLLEAGUE EDIT POINT: change these drawing coordinates only when the
    # quantity step-arrow shape is intentionally redesigned.
    """Create spin arrow png for this workflow."""
    if os.path.exists(path):
        return
    pm = QPixmap(18, 18)
    pm.fill(Qt.transparent)

    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing, True)
    pen = QPen(QColor(color_hex))
    pen.setWidthF(3.2)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    p.setPen(pen)

    if direction == "up":
        p.drawLine(4, 11, 9, 6)
        p.drawLine(9, 6, 14, 11)
    else:
        p.drawLine(4, 7, 9, 12)
        p.drawLine(9, 12, 14, 7)

    p.end()
    pm.save(path, "PNG")


def ensure_arrow_icons():
    """Ensure arrow icons for this workflow."""
    make_spin_arrow_png(SPIN_UP_LIGHT_PNG, "#111111", "up")
    make_spin_arrow_png(SPIN_DOWN_LIGHT_PNG, "#111111", "down")
    make_spin_arrow_png(SPIN_UP_DARK_PNG, "#FFFFFF", "up")
    make_spin_arrow_png(SPIN_DOWN_DARK_PNG, "#FFFFFF", "down")


def resolve_portable_path(path: str) -> str:
    """Resolve portable path for this workflow."""
    path = (path or "").strip()
    if not path or os.path.isabs(path):
        return path
    return os.path.normpath(os.path.join(APP_DIR, path))


def portable_stored_path(path: str) -> str:
    """Handle portable stored path."""
    path = (path or "").strip()
    if not path:
        return ""
    try:
        relative = os.path.relpath(os.path.abspath(path), APP_DIR)
    except ValueError:
        return path
    if relative == ".." or relative.startswith(".." + os.sep):
        return path
    return relative.replace("\\", "/")


def icon_or_empty(path: str) -> QIcon:
    """Handle icon or empty."""
    resolved = resolve_portable_path(path)
    return QIcon(resolved) if resolved and os.path.exists(resolved) else QIcon()


def pixmap_or_empty(path: str, size: int = 18) -> QPixmap:
    """Handle pixmap or empty."""
    path = resolve_portable_path(path)
    if not path or not os.path.exists(path):
        return QPixmap()
    pm = QPixmap(path)
    if pm.isNull():
        return QPixmap()
    return pm.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation)


def make_form_label(text: str, icon_path: str) -> QWidget:
    # COLLEAGUE EDIT POINT: label icon size and spacing for every form row.
    """Create form label for this workflow."""
    w = QWidget()
    lay = QHBoxLayout(w)
    lay.setContentsMargins(0, 0, 0, 0)
    lay.setSpacing(10)

    icon = QLabel()
    icon.setPixmap(pixmap_or_empty(icon_path, 17))
    icon.setFixedSize(20, 20)
    icon.setAlignment(Qt.AlignCenter)

    lbl = QLabel(text)
    lbl.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)

    lay.addWidget(icon)
    lay.addWidget(lbl)
    lay.addStretch(1)
    return w


def normalized_person_key(text: str) -> str:
    """Handle normalized person key."""
    return SEARCH_CLEAN_RE.sub("", (text or "").strip().lower().replace("ё", "е"))


def is_excluded_executor(text: str) -> bool:
    """Handle is excluded executor."""
    return normalized_person_key(text) in EXCLUDED_EXECUTOR_KEYS


def read_lines(path: str) -> List[str]:
    """Read lines for this workflow."""
    if not os.path.exists(path):
        return []
    with open(path, "r", encoding="utf-8") as f:
        return [
            text for line in f
            if (text := line.strip()) and not is_excluded_executor(text)
        ]


def write_lines(path: str, items: List[str]):
    """Write lines for this workflow."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    clean = [
        text for x in items
        if x and (text := x.strip()) and not is_excluded_executor(text)
    ]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(clean) + ("\n" if clean else ""))


def remove_excluded_executors_from_file():
    """Remove excluded executors from file for this workflow."""
    path = FILES["executors"]
    if not os.path.isfile(path):
        return
    with open(path, "r", encoding="utf-8") as stream:
        original = [line.rstrip("\r\n") for line in stream]
    filtered = [line for line in original if not is_excluded_executor(line)]
    if filtered != original:
        write_lines(path, filtered)


def safe_filename(s: str) -> str:
    """Handle safe filename."""
    s = (s or "").strip()
    return INVALID_FILENAME_RE.sub("_", s)[:180]


def open_folder(path: str):
    """Open folder for this workflow."""
    if os.name == "nt":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.run(["open", path])
    else:
        subprocess.run(["xdg-open", path])


def normalize_url_or_path(s: str) -> QUrl:
    """Normalize url or path for this workflow."""
    s = (s or "").strip()
    if not s:
        return QUrl()
    if URL_SCHEME_RE.match(s):
        return QUrl(s)
    if os.path.exists(s):
        return QUrl.fromLocalFile(os.path.abspath(s))
    if s.startswith("www."):
        return QUrl("https://" + s)
    return QUrl(s)


def open_any(link_or_path: str):
    """Open any for this workflow."""
    url = normalize_url_or_path(link_or_path)
    if not url.isValid() or url.isEmpty():
        QMessageBox.warning(None, "Открытие", "Пустая/неверная ссылка или путь.")
        return
    QDesktopServices.openUrl(url)


def open_file(path: str):
    """Open file for this workflow."""
    if not path:
        return
    QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.abspath(path)))


def move_file_to_recycle_bin(path: str):
    """Move one file to the Windows Recycle Bin without extra dependencies."""
    if os.name != "nt":
        raise OSError("Отправка в Корзину поддерживается только в Windows.")

    class SHFILEOPSTRUCTW(ctypes.Structure):
        _fields_ = (
            ("hwnd", ctypes.wintypes.HWND),
            ("wFunc", ctypes.c_uint),
            ("pFrom", ctypes.c_wchar_p),
            ("pTo", ctypes.c_wchar_p),
            ("fFlags", ctypes.c_ushort),
            ("fAnyOperationsAborted", ctypes.wintypes.BOOL),
            ("hNameMappings", ctypes.c_void_p),
            ("lpszProgressTitle", ctypes.c_wchar_p),
        )

    operation = SHFILEOPSTRUCTW()
    operation.wFunc = 3  # FO_DELETE
    operation.pFrom = os.path.abspath(path) + "\0"
    operation.fFlags = 0x0040 | 0x0010 | 0x0004 | 0x0400
    result = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(operation))
    if result != 0 or operation.fAnyOperationsAborted:
        raise OSError(f"Windows не смогла переместить файл в Корзину (код {result}).")


def normalize_search(text: str) -> str:
    """Normalize search for this workflow."""
    return SEARCH_CLEAN_RE.sub("", (text or "").strip().lower().replace("ё", "е"))


def replace_in_paragraph(paragraph, mapping: Dict[str, str]):
    """Handle replace in paragraph."""
    full = "".join(run.text for run in paragraph.runs)
    if not full:
        return
    changed = False
    for key, val in mapping.items():
        if key in full:
            full = full.replace(key, val)
            changed = True
    if changed:
        if paragraph.runs:
            paragraph.runs[0].text = full
            for r in paragraph.runs[1:]:
                r.text = ""
        else:
            paragraph.add_run(full)


def replace_placeholders_docx(doc: Document, mapping: Dict[str, str]):
    """Replace placeholders throughout paragraphs, tables, headers and footers."""
    for p in doc.paragraphs:
        replace_in_paragraph(p, mapping)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p, mapping)


# -----------------------------
# Links storage (JSON)
# -----------------------------
def load_links() -> List[Dict[str, str]]:
    """Load links for this workflow."""
    if not os.path.exists(FILES["links"]):
        return []
    try:
        with open(FILES["links"], "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, list):
            return []
        out = [
            {
                "name": str(item.get("name", "")).strip(),
                "link": str(item.get("link", "")).strip(),
                "icon": str(item.get("icon", "")).strip(),
            }
            for item in data
            if isinstance(item, dict)
        ]
        return [item for item in out if any(item.values())]
    except (OSError, json.JSONDecodeError):
        return []


def save_links(links: List[Dict[str, str]]):
    """Save links for this workflow."""
    os.makedirs(os.path.dirname(FILES["links"]), exist_ok=True)
    clean = [
        item
        for link_data in links
        if any((item := {
            "name": (link_data.get("name") or "").strip(),
            "link": (link_data.get("link") or "").strip(),
            "icon": (link_data.get("icon") or "").strip(),
        }).values())
    ]
    with open(FILES["links"], "w", encoding="utf-8") as f:
        json.dump(clean, f, ensure_ascii=False, indent=2)


def load_app_settings() -> Dict[str, bool]:
    """Load app settings for this workflow."""
    settings = {"preview_before_save": False}
    try:
        with open(APP_SETTINGS_CONFIG, "r", encoding="utf-8") as stream:
            saved = json.load(stream)
        if isinstance(saved, dict) and isinstance(saved.get("preview_before_save"), bool):
            settings["preview_before_save"] = saved["preview_before_save"]
    except (OSError, json.JSONDecodeError):
        pass
    return settings


def save_app_settings(settings: Dict[str, bool]):
    """Save app settings for this workflow."""
    os.makedirs(os.path.dirname(APP_SETTINGS_CONFIG), exist_ok=True)
    temporary_path = APP_SETTINGS_CONFIG + ".tmp"
    with open(temporary_path, "w", encoding="utf-8") as stream:
        json.dump(settings, stream, ensure_ascii=False, indent=2)
    os.replace(temporary_path, APP_SETTINGS_CONFIG)


# -----------------------------
# Windows acrylic helper (safe)
# -----------------------------
def _is_windows() -> bool:
    """Handle is windows."""
    return os.name == "nt"


def _try_enable_acrylic(hwnd: int, dark: bool):
    """Handle try enable acrylic."""
    if not _is_windows() or not hwnd:
        return
    try:
        user32 = ctypes.windll.user32

        class ACCENT_POLICY(ctypes.Structure):
            _fields_ = [
                ("AccentState", ctypes.c_int),
                ("AccentFlags", ctypes.c_int),
                ("GradientColor", ctypes.c_int),
                ("AnimationId", ctypes.c_int),
            ]

        class WINDOWCOMPOSITIONATTRIBDATA(ctypes.Structure):
            _fields_ = [
                ("Attribute", ctypes.c_int),
                ("Data", ctypes.c_void_p),
                ("SizeOfData", ctypes.c_size_t),
            ]

        WCA_ACCENT_POLICY = 19
        ACCENT_ENABLE_ACRYLICBLURBEHIND = 4

        if dark:
            r, g, b, a = 22, 24, 28, 210
        else:
            r, g, b, a = 255, 255, 255, 200

        gradient = (a << 24) | (b << 16) | (g << 8) | r

        accent = ACCENT_POLICY()
        accent.AccentState = ACCENT_ENABLE_ACRYLICBLURBEHIND
        accent.AccentFlags = 2
        accent.GradientColor = gradient
        accent.AnimationId = 0

        data = WINDOWCOMPOSITIONATTRIBDATA()
        data.Attribute = WCA_ACCENT_POLICY
        data.Data = ctypes.cast(ctypes.pointer(accent), ctypes.c_void_p)
        data.SizeOfData = ctypes.sizeof(accent)

        set_wca = getattr(user32, "SetWindowCompositionAttribute", None)
        if set_wca:
            set_wca(hwnd, ctypes.byref(data))
    except Exception:
        pass


# -----------------------------
# Dialogs
# -----------------------------
class ListEditorDialog(QDialog):
    def __init__(self, title: str, file_path: str, parent=None):
        """Initialize the ListEditorDialog and its runtime state."""
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(720, 520)
        self.file_path = file_path

        self.listw = QListWidget()
        self.listw.setSelectionMode(QAbstractItemView.ExtendedSelection)

        self.btn_add = QPushButton("Добавить")
        self.btn_edit = QPushButton("Изменить")
        self.btn_del = QPushButton("Удалить")
        self.btn_up = QPushButton("Вверх")
        self.btn_down = QPushButton("Вниз")
        self.btn_save = QPushButton("Сохранить")
        self.btn_close = QPushButton("Закрыть")

        left = QVBoxLayout()
        left.addWidget(self.listw)

        right = QVBoxLayout()
        right.setSpacing(8)
        right.addWidget(self.btn_add)
        right.addWidget(self.btn_edit)
        right.addWidget(self.btn_del)
        right.addSpacing(10)
        right.addWidget(self.btn_up)
        right.addWidget(self.btn_down)
        right.addStretch()
        right.addWidget(self.btn_save)
        right.addWidget(self.btn_close)

        root = QVBoxLayout(self)
        wrapper = QWidget()
        row = QHBoxLayout(wrapper)
        row.setContentsMargins(14, 14, 14, 14)
        row.setSpacing(12)
        row.addLayout(left, 1)
        row.addLayout(right)
        root.addWidget(wrapper)

        self.btn_add.clicked.connect(self.add_item)
        self.btn_edit.clicked.connect(self.edit_item)
        self.btn_del.clicked.connect(self.del_item)
        self.btn_up.clicked.connect(self.move_up)
        self.btn_down.clicked.connect(self.move_down)
        self.btn_save.clicked.connect(self.save)
        self.btn_close.clicked.connect(self.close)

        self.load()

    def load(self):
        """Load the related value for this workflow."""
        self.listw.clear()
        for x in read_lines(self.file_path):
            self.listw.addItem(QListWidgetItem(x))

    def add_item(self):
        """Add item for this workflow."""
        text, ok = QInputDialog.getText(self, "Добавить", "Введите пункт:")
        if ok and text.strip():
            self.listw.addItem(QListWidgetItem(text.strip()))

    def edit_item(self):
        """Edit item for this workflow."""
        items = self.listw.selectedItems()
        if len(items) != 1:
            QMessageBox.information(self, "Изменить", "Выберите ровно один пункт.")
            return
        cur = items[0].text()
        text, ok = QInputDialog.getText(self, "Изменить", "Измените пункт:", text=cur)
        if ok and text.strip():
            items[0].setText(text.strip())

    def del_item(self):
        """Handle del item in ListEditorDialog."""
        for it in self.listw.selectedItems():
            row = self.listw.row(it)
            self.listw.takeItem(row)

    def move_up(self):
        """Handle move up in ListEditorDialog."""
        row = self.listw.currentRow()
        if row <= 0:
            return
        it = self.listw.takeItem(row)
        self.listw.insertItem(row - 1, it)
        self.listw.setCurrentRow(row - 1)

    def move_down(self):
        """Handle move down in ListEditorDialog."""
        row = self.listw.currentRow()
        if row < 0 or row >= self.listw.count() - 1:
            return
        it = self.listw.takeItem(row)
        self.listw.insertItem(row + 1, it)
        self.listw.setCurrentRow(row + 1)

    def save(self):
        """Save the related value for this workflow."""
        items = [self.listw.item(i).text().strip() for i in range(self.listw.count())]
        items = [x for x in items if x]
        write_lines(self.file_path, items)
        QMessageBox.information(self, "Сохранено", "Список сохранён.")


class LinksEditorDialog(QDialog):
    def __init__(self, links: List[Dict[str, str]], parent=None):
        """Initialize the LinksEditorDialog and its runtime state."""
        super().__init__(parent)
        self.setWindowTitle("Редактор ссылок")
        self.resize(860, 580)
        self.links = [dict(x) for x in links]

        self.listw = QListWidget()
        self.listw.setSelectionMode(QAbstractItemView.SingleSelection)

        self.btn_add = QPushButton("Добавить")
        self.btn_edit = QPushButton("Изменить")
        self.btn_icon = QPushButton("Иконка…")
        self.btn_clear_icon = QPushButton("Сбросить иконку")
        self.btn_del = QPushButton("Удалить")
        self.btn_up = QPushButton("Вверх")
        self.btn_down = QPushButton("Вниз")
        self.btn_save = QPushButton("Сохранить")
        self.btn_close = QPushButton("Закрыть")

        left = QVBoxLayout()
        left.addWidget(self.listw)

        right = QVBoxLayout()
        right.setSpacing(8)
        right.addWidget(self.btn_add)
        right.addWidget(self.btn_edit)
        right.addWidget(self.btn_icon)
        right.addWidget(self.btn_clear_icon)
        right.addWidget(self.btn_del)
        right.addSpacing(10)
        right.addWidget(self.btn_up)
        right.addWidget(self.btn_down)
        right.addStretch()
        right.addWidget(self.btn_save)
        right.addWidget(self.btn_close)

        wrapper = QWidget()
        row = QHBoxLayout(wrapper)
        row.setContentsMargins(14, 14, 14, 14)
        row.setSpacing(12)
        row.addLayout(left, 1)
        row.addLayout(right)

        root = QVBoxLayout(self)
        root.addWidget(wrapper)

        self.btn_add.clicked.connect(self.add_item)
        self.btn_edit.clicked.connect(self.edit_item)
        self.btn_icon.clicked.connect(self.change_icon_for_selected)
        self.btn_clear_icon.clicked.connect(self.clear_icon_for_selected)
        self.btn_del.clicked.connect(self.del_item)
        self.btn_up.clicked.connect(self.move_up)
        self.btn_down.clicked.connect(self.move_down)
        self.btn_save.clicked.connect(self.accept)
        self.btn_close.clicked.connect(self.reject)

        self.refresh()

    def _pick_icon(self, current: str = "") -> str:
        """Handle pick icon in LinksEditorDialog."""
        current_path = resolve_portable_path(current)
        start_dir = os.path.dirname(current_path) if current_path and os.path.exists(current_path) else ICONS_DIR
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите иконку",
            start_dir,
            "Images (*.png *.ico *.jpg *.jpeg *.bmp *.webp);;All files (*.*)"
        )
        return portable_stored_path(path)

    def refresh(self):
        """Refresh the related value for this workflow."""
        self.listw.clear()
        for t in self.links:
            name = (t.get("name") or "").strip() or "(без названия)"
            link = (t.get("link") or "").strip()
            it = QListWidgetItem(f"{name}  —  {link}")
            ico = icon_or_empty((t.get("icon") or "").strip())
            if not ico.isNull():
                it.setIcon(ico)
            self.listw.addItem(it)

    def add_item(self):
        """Add item for this workflow."""
        name, ok = QInputDialog.getText(self, "Название", "Название ссылки:")
        if not ok:
            return
        link, ok2 = QInputDialog.getText(self, "Ссылка/путь", "Ссылка или путь к файлу:")
        if not ok2:
            return
        icon = self._pick_icon("")
        self.links.append({"name": name.strip(), "link": link.strip(), "icon": icon})
        self.refresh()

    def edit_item(self):
        """Edit item for this workflow."""
        row = self.listw.currentRow()
        if row < 0:
            QMessageBox.information(self, "Изменить", "Выберите ссылку.")
            return

        cur = self.links[row]
        name, ok = QInputDialog.getText(self, "Название", "Название ссылки:", text=cur.get("name", ""))
        if not ok:
            return
        link, ok2 = QInputDialog.getText(self, "Ссылка/путь", "Ссылка или путь к файлу:", text=cur.get("link", ""))
        if not ok2:
            return

        self.links[row]["name"] = name.strip()
        self.links[row]["link"] = link.strip()
        self.refresh()
        self.listw.setCurrentRow(row)

    def change_icon_for_selected(self):
        """Handle change icon for selected in LinksEditorDialog."""
        row = self.listw.currentRow()
        if row < 0:
            QMessageBox.information(self, "Иконка", "Выберите ссылку.")
            return
        cur = self.links[row]
        picked = self._pick_icon(cur.get("icon", ""))
        if picked:
            self.links[row]["icon"] = picked
            self.refresh()
            self.listw.setCurrentRow(row)

    def clear_icon_for_selected(self):
        """Clear icon for selected for this workflow."""
        row = self.listw.currentRow()
        if row < 0:
            return
        self.links[row]["icon"] = ""
        self.refresh()
        self.listw.setCurrentRow(row)

    def del_item(self):
        """Handle del item in LinksEditorDialog."""
        row = self.listw.currentRow()
        if row < 0:
            return
        self.links.pop(row)
        self.refresh()

    def move_up(self):
        """Handle move up in LinksEditorDialog."""
        row = self.listw.currentRow()
        if row <= 0:
            return
        self.links[row - 1], self.links[row] = self.links[row], self.links[row - 1]
        self.refresh()
        self.listw.setCurrentRow(row - 1)

    def move_down(self):
        """Handle move down in LinksEditorDialog."""
        row = self.listw.currentRow()
        if row < 0 or row >= len(self.links) - 1:
            return
        self.links[row + 1], self.links[row] = self.links[row], self.links[row + 1]
        self.refresh()
        self.listw.setCurrentRow(row + 1)


# -----------------------------
# Search combo
# -----------------------------
class SearchCombo(QComboBox):
    def __init__(self, items: List[str], editor_title: str, editor_file: str, parent=None):
        """Initialize the SearchCombo and its runtime state."""
        super().__init__(parent)
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.NoInsert)
        self.setMaxVisibleItems(14)
        # Long list values must not enlarge the scroll area's hidden content
        # width.  The editor stays fluid and elides naturally inside the form.
        self.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.setMinimumContentsLength(1)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self._editor_title = editor_title
        self._editor_file = editor_file

        self._all_items: List[str] = []
        self._norm_map: List[tuple[str, str]] = []

        self._hint_model = QStringListModel(self)
        self._completer = QCompleter(self._hint_model, self)
        self._completer.setCompletionMode(QCompleter.PopupCompletion)
        self._completer.setCaseSensitivity(Qt.CaseInsensitive)
        self._completer.setFilterMode(Qt.MatchContains)

        self._completer.activated.connect(self._on_completer_activated)
        self.lineEdit().setCompleter(self._completer)

        self.lineEdit().textEdited.connect(self._on_text_edited)
        self.lineEdit().returnPressed.connect(self._on_return_pressed)

        gear_icon = self.style().standardIcon(QStyle.SP_FileDialogDetailedView)
        act = QAction(gear_icon, "Редактировать список", self)
        act.setToolTip("Редактировать список")
        act.setIconVisibleInMenu(False)
        act.triggered.connect(self.open_editor)
        self.lineEdit().addAction(act, QLineEdit.TrailingPosition)
        self.lineEdit().setClearButtonEnabled(False)

        self.set_items(items)

    def set_items(self, items: List[str]):
        """Set items for this workflow."""
        cur = (self.currentText() or "").strip()
        self._all_items = list(dict.fromkeys(
            text for item in (items or []) if item and (text := item.strip())
        ))
        self._norm_map = [(normalize_search(x), x) for x in self._all_items]

        with QSignalBlocker(self):
            self.clear()
            self.addItems(self._all_items)

        self._hint_model.setStringList(self._all_items)
        self.setEditText(cur)

    def open_editor(self):
        """Open editor for this workflow."""
        dlg = ListEditorDialog(self._editor_title, self._editor_file, self.window())
        dlg.exec()
        self.set_items(read_lines(self._editor_file))
        self._show_hints()

    def _build_hints(self, typed: str) -> List[str]:
        """Build hints for this workflow."""
        qn = normalize_search(typed)
        if not qn:
            return self._all_items[:]
        hits = [orig for n, orig in self._norm_map if qn in n]
        if not hits:
            parts = [normalize_search(p) for p in re.split(r"\s+", typed.strip()) if p.strip()]
            parts = [p for p in parts if p]
            if parts:
                hits = [orig for n, orig in self._norm_map if all(p in n for p in parts)]
        return hits

    def _on_text_edited(self, text: str):
        """Handle the on text edited interface event."""
        self._hint_model.setStringList(self._build_hints(text))
        self._show_hints()

    def _show_hints(self):
        """Show hints for this workflow."""
        if self._hint_model.rowCount() > 0:
            self._completer.complete()

    def _on_completer_activated(self, text: str):
        """Handle the on completer activated interface event."""
        text = (text or "").strip()
        if not text:
            return
        self.setEditText(text)
        idx = self.findText(text, Qt.MatchExactly)
        if idx >= 0:
            self.setCurrentIndex(idx)

    def _on_return_pressed(self):
        """Handle the on return pressed interface event."""
        typed = (self.lineEdit().text() or "").strip()
        if not typed:
            return
        hints = self._build_hints(typed)
        if hints:
            self._on_completer_activated(hints[0])
            return
        idx = self.findText(typed, Qt.MatchExactly)
        if idx >= 0:
            self.setCurrentIndex(idx)
        else:
            self.setEditText(typed)

    def text(self) -> str:
        """Handle text in SearchCombo."""
        return (self.currentText() or "").strip()


class SerialCombo(QComboBox):
    """Editable factory-number field with a real drop-down arrow."""

    def __init__(self, items: List[str], parent=None):
        """Initialize the SerialCombo and its runtime state."""
        super().__init__(parent)
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.NoInsert)
        self.setMaxVisibleItems(18)
        self.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.setMinimumContentsLength(1)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.set_items(items, current_text="")

        completer = self.completer()
        completer.setCompletionMode(QCompleter.PopupCompletion)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)

    def set_items(self, items: List[str], current_text: str | None = None):
        """Replace drop-down values while preserving requested manual text."""
        text = self.text() if current_text is None else str(current_text or "")
        values = list(dict.fromkeys(
            value for item in (items or [])
            if (value := normalize_registry_text(item))
        ))
        with QSignalBlocker(self):
            super().clear()
            self.addItems(values)
            self.setCurrentIndex(-1)
            self.setEditText(text)

    def setPlaceholderText(self, text: str):
        """Handle setPlaceholderText in SerialCombo."""
        self.lineEdit().setPlaceholderText(text)

    def setText(self, text: str):
        """Handle setText in SerialCombo."""
        self.setEditText(str(text or ""))

    def text(self) -> str:
        """Handle text in SerialCombo."""
        return (self.currentText() or "").strip()

    def clear(self):
        """Clear entered text without removing the drop-down registry."""
        self.setCurrentIndex(-1)
        self.setEditText("")


class QuantitySpinBox(QWidget):
    """Compact quantity field with step buttons contained inside its frame."""

    def __init__(self, parent=None):
        """Initialize the QuantitySpinBox and its runtime state."""
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setFixedHeight(42)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.container = QFrame(self)
        self.container.setObjectName("QuantityField")
        field_layout = QHBoxLayout(self.container)
        field_layout.setContentsMargins(0, 0, 0, 0)
        field_layout.setSpacing(0)
        layout.addWidget(self.container)

        self.editor = QSpinBox(self.container)
        self.editor.setObjectName("QuantityEditor")
        self.editor.setButtonSymbols(QAbstractSpinBox.NoButtons)
        field_layout.addWidget(self.editor, 1)

        step_column = QWidget(self.container)
        self.step_column = step_column
        step_column.setFixedWidth(30)
        step_layout = QVBoxLayout(step_column)
        step_layout.setContentsMargins(0, 0, 0, 0)
        step_layout.setSpacing(0)

        self.up_button = QToolButton(step_column)
        self.up_button.setObjectName("QuantityUpButton")
        self.up_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.up_button.setIconSize(QSize(12, 12))
        self.up_button.clicked.connect(self.editor.stepUp)
        step_layout.addWidget(self.up_button, 1)

        self.down_button = QToolButton(step_column)
        self.down_button.setObjectName("QuantityDownButton")
        self.down_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.down_button.setIconSize(QSize(12, 12))
        self.down_button.clicked.connect(self.editor.stepDown)
        step_layout.addWidget(self.down_button, 1)
        field_layout.addWidget(step_column)

    def set_arrow_icons(self, up_path: str, down_path: str):
        """Set arrow icons for this workflow."""
        self.up_button.setIcon(icon_or_empty(up_path))
        self.down_button.setIcon(icon_or_empty(down_path))

    def setRange(self, minimum: int, maximum: int):
        """Handle setRange in QuantitySpinBox."""
        self.editor.setRange(minimum, maximum)

    def setValue(self, value: int):
        """Handle setValue in QuantitySpinBox."""
        self.editor.setValue(value)

    def value(self) -> int:
        """Handle value in QuantitySpinBox."""
        return self.editor.value()


class TwoDigitSpinBox(QSpinBox):
    """Spin box that keeps hours and minutes visually zero-padded."""

    def textFromValue(self, value: int) -> str:
        """Handle textFromValue in TwoDigitSpinBox."""
        return f"{value:02d}"


class TimePickerSpinBox(QWidget):
    """Two-digit time selector with themed, centered step arrows."""

    def __init__(self, parent=None):
        """Initialize the TimePickerSpinBox and its runtime state."""
        super().__init__(parent)
        self.setFixedSize(88, 48)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.container = QFrame(self)
        self.container.setObjectName("TimePickerSpinField")
        field_layout = QHBoxLayout(self.container)
        field_layout.setContentsMargins(0, 0, 0, 0)
        field_layout.setSpacing(0)
        layout.addWidget(self.container)

        self.editor = TwoDigitSpinBox(self.container)
        self.editor.setObjectName("TimePickerSpinEditor")
        self.editor.setButtonSymbols(QAbstractSpinBox.NoButtons)
        self.editor.setAlignment(Qt.AlignCenter)
        field_layout.addWidget(self.editor, 1)

        step_column = QWidget(self.container)
        step_column.setFixedWidth(28)
        step_layout = QVBoxLayout(step_column)
        step_layout.setContentsMargins(0, 0, 0, 0)
        step_layout.setSpacing(0)

        self.up_button = QToolButton(step_column)
        self.up_button.setObjectName("TimePickerUpButton")
        uses_dark_theme = (
            QApplication.instance().palette().color(QPalette.WindowText).lightness()
            > 128
        )
        up_icon = SPIN_UP_DARK_PNG if uses_dark_theme else SPIN_UP_LIGHT_PNG
        down_icon = SPIN_DOWN_DARK_PNG if uses_dark_theme else SPIN_DOWN_LIGHT_PNG
        self.up_button.setIcon(icon_or_empty(up_icon))
        self.up_button.setIconSize(QSize(11, 11))
        self.up_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.up_button.setToolTip("Увеличить")
        self.up_button.clicked.connect(self.editor.stepUp)
        step_layout.addWidget(self.up_button, 1)

        self.down_button = QToolButton(step_column)
        self.down_button.setObjectName("TimePickerDownButton")
        self.down_button.setIcon(icon_or_empty(down_icon))
        self.down_button.setIconSize(QSize(11, 11))
        self.down_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.down_button.setToolTip("Уменьшить")
        self.down_button.clicked.connect(self.editor.stepDown)
        step_layout.addWidget(self.down_button, 1)
        field_layout.addWidget(step_column)

    def setRange(self, minimum: int, maximum: int):
        """Handle setRange in TimePickerSpinBox."""
        self.editor.setRange(minimum, maximum)

    def setValue(self, value: int):
        """Handle setValue in TimePickerSpinBox."""
        self.editor.setValue(value)

    def value(self) -> int:
        """Handle value in TimePickerSpinBox."""
        return self.editor.value()


class EditableTimeEdit(QWidget):
    """Manual HH:MM field with an optional graphical hour/minute picker."""

    def __init__(self, parent=None):
        """Initialize the EditableTimeEdit and its runtime state."""
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setFixedHeight(42)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.container = QFrame(self)
        self.container.setObjectName("TimeField")
        field_layout = QHBoxLayout(self.container)
        field_layout.setContentsMargins(0, 0, 0, 0)
        field_layout.setSpacing(0)
        layout.addWidget(self.container)

        self.editor = QLineEdit(self.container)
        self.editor.setObjectName("TimeEditor")
        self.editor.setInputMask(SERVICE_TIME_INPUT_MASK)
        self.editor.setPlaceholderText("__:__")
        self.editor.setToolTip(
            "Введите время цифрами или выберите часы и минуты"
        )
        field_layout.addWidget(self.editor, 1)

        self.picker_button = QToolButton(self.container)
        self.picker_button.setObjectName("TimePickerButton")
        self.picker_button.setIcon(icon_or_empty(ICON_ARROW_DOWN))
        self.picker_button.setIconSize(QSize(18, 18))
        self.picker_button.setFixedWidth(30)
        self.picker_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
        self.picker_button.setToolTip("Выбрать время")
        self.picker_button.clicked.connect(self.show_time_picker)
        field_layout.addWidget(self.picker_button)
        # The service-period row intentionally has no visible arrow buttons.
        # Manual HH:MM input remains available; the picker implementation is
        # retained for compatibility with old history/tests and future reuse.
        self.picker_button.hide()

        self._time_menu = None
        self._hour_spin = None
        self._minute_spin = None

    def show_time_picker(self):
        """Show time picker for this workflow."""
        menu = QMenu(self)
        panel = QWidget(menu)
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(12, 10, 12, 12)
        panel_layout.setSpacing(8)

        title = QLabel("Выберите время", panel)
        title.setObjectName("TimePickerTitle")
        panel_layout.addWidget(title)

        controls_layout = QHBoxLayout()
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(8)

        hour_column = QVBoxLayout()
        hour_column.setContentsMargins(0, 0, 0, 0)
        hour_column.setSpacing(4)
        hour_label = QLabel("Часы", panel)
        hour_label.setObjectName("TimePickerCaption")
        hour_label.setAlignment(Qt.AlignCenter)
        hour_column.addWidget(hour_label)
        hour_spin = TimePickerSpinBox(panel)
        hour_spin.setRange(0, 23)
        hour_spin.setToolTip("Часы")
        hour_column.addWidget(hour_spin)

        minute_column = QVBoxLayout()
        minute_column.setContentsMargins(0, 0, 0, 0)
        minute_column.setSpacing(4)
        minute_label = QLabel("Минуты", panel)
        minute_label.setObjectName("TimePickerCaption")
        minute_label.setAlignment(Qt.AlignCenter)
        minute_column.addWidget(minute_label)
        minute_spin = TimePickerSpinBox(panel)
        minute_spin.setRange(0, 59)
        minute_spin.setToolTip("Минуты")
        minute_column.addWidget(minute_spin)
        current = normalize_service_time_input(self.text())
        match = re.fullmatch(r"(\d{2}):(\d{2})", current)
        now = datetime.datetime.now()
        hour_spin.setValue(int(match.group(1)) if match else now.hour)
        minute_spin.setValue(int(match.group(2)) if match else now.minute)
        controls_layout.addLayout(hour_column)
        separator = QLabel(":", panel)
        separator.setObjectName("TimePickerSeparator")
        separator.setAlignment(Qt.AlignCenter)
        controls_layout.addWidget(separator)
        controls_layout.addLayout(minute_column)
        panel_layout.addLayout(controls_layout)

        select_button = QPushButton("Выбрать", panel)
        select_button.clicked.connect(
            lambda: self._select_time(
                hour_spin.value(), minute_spin.value(), menu
            )
        )
        panel_layout.addWidget(select_button)

        action = QWidgetAction(menu)
        action.setDefaultWidget(panel)
        menu.addAction(action)
        self._time_menu = menu
        self._hour_spin = hour_spin
        self._minute_spin = minute_spin
        menu.popup(
            self.picker_button.mapToGlobal(QPoint(0, self.picker_button.height()))
        )

    def _select_time(self, hour: int, minute: int, menu: QMenu):
        """Handle select time in EditableTimeEdit."""
        self.editor.setText(f"{hour:02d}:{minute:02d}")
        menu.close()

    def text(self) -> str:
        """Handle text in EditableTimeEdit."""
        return self.editor.text().strip()

    def setText(self, value: str):
        """Handle setText in EditableTimeEdit."""
        self.editor.setText(normalize_service_time_input(value))

    def clear(self):
        """Clear the related value for this workflow."""
        self.editor.clear()

    def insert(self, value: str):
        """Handle insert in EditableTimeEdit."""
        self.editor.insert(value)

    def inputMask(self) -> str:
        """Handle inputMask in EditableTimeEdit."""
        return self.editor.inputMask()

    def placeholderText(self) -> str:
        """Handle placeholderText in EditableTimeEdit."""
        return self.editor.placeholderText()


class EditableDateEdit(QWidget):
    """Calendar picker that intentionally keeps direct keyboard entry."""

    def __init__(self, parent=None):
        """Initialize the EditableDateEdit and its runtime state."""
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setFixedHeight(42)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.container = QFrame(self)
        self.container.setObjectName("DateField")
        field_layout = QHBoxLayout(self.container)
        field_layout.setContentsMargins(0, 0, 0, 0)
        field_layout.setSpacing(0)
        layout.addWidget(self.container)

        self.editor = QDateEdit(QDate.currentDate(), self.container)
        self.editor.setObjectName("DateEditor")
        self.editor.setDisplayFormat("dd.MM.yyyy")
        self.editor.setReadOnly(False)
        self.editor.setKeyboardTracking(False)
        self.editor.setButtonSymbols(QAbstractSpinBox.NoButtons)
        field_layout.addWidget(self.editor, 1)

        self.calendar_button = QToolButton(self.container)
        self.calendar_button.setObjectName("DatePickerButton")
        self.calendar_button.setIcon(icon_or_empty(ICON_ARROW_DOWN))
        self.calendar_button.setIconSize(QSize(18, 18))
        self.calendar_button.setFixedWidth(30)
        self.calendar_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
        self.calendar_button.setToolTip("Открыть календарь")
        self.calendar_button.clicked.connect(self.show_calendar)
        field_layout.addWidget(self.calendar_button)

        self._calendar_menu = None
        self.setToolTip("Выберите дату в календаре или введите её вручную")

    def show_calendar(self):
        """Show calendar for this workflow."""
        menu = QMenu(self)
        calendar = QCalendarWidget(menu)
        calendar.setGridVisible(True)
        calendar.setSelectedDate(self.editor.date())
        action = QWidgetAction(menu)
        action.setDefaultWidget(calendar)
        menu.addAction(action)
        calendar.clicked.connect(lambda date: self._select_calendar_date(date, menu))
        self._calendar_menu = menu
        menu.popup(self.calendar_button.mapToGlobal(QPoint(0, self.calendar_button.height())))

    def _select_calendar_date(self, date: QDate, menu: QMenu):
        """Handle select calendar date in EditableDateEdit."""
        self.editor.setDate(date)
        menu.close()

    def text(self) -> str:
        """Handle text in EditableDateEdit."""
        return self.editor.lineEdit().text().strip()

    def setText(self, value: str):
        """Handle setText in EditableDateEdit."""
        parsed = QDate.fromString(str(value).strip(), "dd.MM.yyyy")
        if parsed.isValid():
            self.editor.setDate(parsed)
        else:
            self.editor.lineEdit().setText(str(value))

    def setDate(self, date: QDate):
        """Handle setDate in EditableDateEdit."""
        self.editor.setDate(date)

    def calendarPopup(self) -> bool:
        """Compatibility with the previous QDateEdit-based field API."""
        return True

    def isReadOnly(self) -> bool:
        """Handle isReadOnly in EditableDateEdit."""
        return self.editor.isReadOnly()


class PreviewRenderThread(QThread):
    """Render a DOCX without blocking construction of the preview dialog."""

    def __init__(self, temporary_path: str, parent=None):
        """Initialize the PreviewRenderThread and its runtime state."""
        super().__init__(parent)
        self.temporary_path = temporary_path
        self.pdf_path = None
        self.error = None

    def run(self):
        """Run the related value for this workflow."""
        try:
            self.pdf_path = render_docx_to_pdf(self.temporary_path)
        except Exception as error:
            self.error = error


class ActPreviewDialog(QDialog):
    """Show the populated document as real rendered pages before saving."""

    def __init__(self, temporary_path: str, parent=None):
        """Initialize the ActPreviewDialog and its runtime state."""
        super().__init__(parent)
        self.temporary_path = temporary_path
        self.pdf_path = None
        self.pdf_bytes = None
        self.pdf_buffer = None
        self.pdf_document = None
        self._loading_step = 0
        self._cancel_requested = False
        self.setWindowTitle("Предпросмотр акта")
        self.resize(980, 780)

        layout = QVBoxLayout(self)
        info = QLabel(
            "Проверьте полностью заполненный акт. Можно открыть временный DOCX, "
            "вернуться к редактированию или подтвердить окончательное сохранение."
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        self.loading_label = QLabel("Подготавливаем визуальный предпросмотр…")
        self.loading_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.loading_label, 1)
        self.preview = QPdfView(self)
        self.preview.setPageMode(QPdfView.PageMode.MultiPage)
        self.preview.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        self.preview.hide()
        layout.addWidget(self.preview, 1)

        controls = QHBoxLayout()
        self.open_button = QPushButton("Открыть документ")
        self.open_button.setToolTip("Станет доступно после подготовки предпросмотра")
        self.open_button.clicked.connect(lambda: open_file(self.temporary_path))
        controls.addWidget(self.open_button)

        self.fit_button = QPushButton("По ширине")
        self.fit_button.clicked.connect(
            lambda: self.preview.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        )
        self.zoom_out_button = QPushButton("−")
        self.zoom_out_button.setFixedWidth(42)
        self.zoom_out_button.clicked.connect(lambda: self._zoom(0.85))
        self.zoom_in_button = QPushButton("+")
        self.zoom_in_button.setFixedWidth(42)
        self.zoom_in_button.clicked.connect(lambda: self._zoom(1.15))
        controls.addWidget(self.fit_button)
        controls.addWidget(self.zoom_out_button)
        controls.addWidget(self.zoom_in_button)
        controls.addStretch(1)

        self.back_button = QPushButton("Вернуться к редактированию")
        self.back_button.clicked.connect(self.reject)
        self.save_button = QPushButton("Подтвердить и сохранить")
        self.save_button.setDefault(True)
        self.save_button.clicked.connect(self.accept)
        controls.addWidget(self.back_button)
        controls.addWidget(self.save_button)
        layout.addLayout(controls)

        for button in (
            self.open_button,
            self.fit_button,
            self.zoom_out_button,
            self.zoom_in_button,
            self.save_button,
        ):
            button.setEnabled(False)

        self.loading_timer = QTimer(self)
        self.loading_timer.setInterval(350)
        self.loading_timer.timeout.connect(self._animate_loading)
        self.loading_timer.start()

        self.render_thread = PreviewRenderThread(temporary_path, self)
        self.render_thread.finished.connect(self._render_finished)
        self.render_thread.start()

    def _animate_loading(self):
        """Handle animate loading in ActPreviewDialog."""
        self._loading_step = (self._loading_step + 1) % 4
        self.loading_label.setText(
            "Подготавливаем визуальный предпросмотр" + "." * self._loading_step
        )

    def _render_finished(self):
        """Handle render finished in ActPreviewDialog."""
        self.loading_timer.stop()
        if self._cancel_requested:
            super().done(QDialog.Rejected)
            return
        if self.render_thread.error is not None:
            QMessageBox.critical(
                self,
                "Ошибка предпросмотра",
                str(self.render_thread.error),
            )
            super().done(QDialog.Rejected)
            return
        try:
            self._load_pdf(self.render_thread.pdf_path)
        except (OSError, PreviewRenderError) as error:
            QMessageBox.critical(self, "Ошибка предпросмотра", str(error))
            super().done(QDialog.Rejected)

    def _load_pdf(self, pdf_path: str):
        """Load pdf for this workflow."""
        self.pdf_path = pdf_path
        with open(self.pdf_path, "rb") as pdf_file:
            self.pdf_bytes = QByteArray(pdf_file.read())
        self.pdf_buffer = QBuffer(self)
        self.pdf_buffer.setData(self.pdf_bytes)
        if not self.pdf_buffer.open(QIODevice.ReadOnly):
            raise PreviewRenderError("Не удалось открыть PDF предпросмотра в памяти.")

        self.pdf_document = QPdfDocument(self)
        load_error = self.pdf_document.load(self.pdf_buffer)
        if load_error is not None and load_error != QPdfDocument.Error.None_:
            raise PreviewRenderError(f"Qt не смог открыть PDF предпросмотра: {load_error}")

        self.preview.setDocument(self.pdf_document)
        self.loading_label.hide()
        self.preview.show()
        self.open_button.setToolTip("")
        for button in (
            self.open_button,
            self.fit_button,
            self.zoom_out_button,
            self.zoom_in_button,
            self.save_button,
        ):
            button.setEnabled(True)

    def _zoom(self, factor: float):
        """Handle zoom in ActPreviewDialog."""
        self.preview.setZoomMode(QPdfView.ZoomMode.Custom)
        self.preview.setZoomFactor(max(0.25, min(4.0, self.preview.zoomFactor() * factor)))

    def done(self, result: int):
        """Handle done in ActPreviewDialog."""
        if self.render_thread.isRunning():
            self._cancel_requested = True
            self.loading_label.setText("Завершаем подготовку предпросмотра…")
            self.save_button.setEnabled(False)
            self.back_button.setEnabled(False)
            return
        if self.pdf_document is not None:
            self.preview.setDocument(None)
            self.pdf_document.close()
        if self.pdf_buffer is not None:
            self.pdf_buffer.close()
        super().done(result)


# -----------------------------
# Main window
# -----------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        """Initialize the MainWindow and its runtime state."""
        super().__init__()
        ensure_dirs()
        ensure_default_icons()
        ensure_arrow_icons()

        self.setWindowTitle("Генератор актов")
        self.resize(1100, 820)

        self.dark_theme = True
        self.links = load_links()
        self.app_settings = load_app_settings()

        self.data_lists = {
            key: read_lines(path)
            for key, path in FILES.items()
            if key != "links"
        }
        # COLLEAGUE EDIT POINT: location_list.txt is the authoritative ordered
        # station list. Everything after АЛАБУШЕВО is intentionally excluded.
        self.data_lists["locations"] = allowed_stations_through_alabushevo(
            self.data_lists["locations"]
        )
        self.equipment_registry = self._load_allowed_equipment_registry()
        # Equipment model values may be extended by the registry, but stations
        # are never extended from the remote table.
        known_models = {item.casefold() for item in self.data_lists["models"]}
        for record in self.equipment_registry.values():
            model = record["model"]
            if model.casefold() not in known_models:
                self.data_lists["models"].append(model)
                known_models.add(model.casefold())

        self._menus_for_mask = set()

        self._build_ui()
        self.apply_theme(True)
        self.refresh_links_menu()

    # ---- Menu round mask helpers ----
    def _apply_menu_round_mask(self, menu: QMenu):
        """Apply menu round mask for this workflow."""
        try:
            r = 14
            rect = menu.rect()
            if rect.width() <= 2 or rect.height() <= 2:
                return
            path = QPainterPath()
            path.addRoundedRect(rect, r, r)
            menu.setMask(QRegion(path.toFillPolygon().toPolygon()))
        except Exception:
            pass

    def eventFilter(self, obj, event):
        """Handle the eventFilter interface event."""
        if isinstance(obj, QMenu) and obj in getattr(self, "_menus_for_mask", set()):
            if event.type() in (QEvent.Show, QEvent.Resize):
                QTimer.singleShot(0, lambda m=obj: self._apply_menu_round_mask(m))
        return super().eventFilter(obj, event)

    def _beautify_topbar(self):
        """Handle beautify topbar in MainWindow."""
        eff = QGraphicsDropShadowEffect(self.topbar)
        eff.setBlurRadius(26)
        eff.setOffset(0, 10)
        eff.setColor(QColor(0, 0, 0, 110))
        self.topbar.setGraphicsEffect(eff)

    def _beautify_menu(self, menu: QMenu):
        """Handle beautify menu in MainWindow."""
        menu.setAttribute(Qt.WA_TranslucentBackground, True)
        menu.setAutoFillBackground(False)

        eff = QGraphicsDropShadowEffect(menu)
        eff.setBlurRadius(28)
        eff.setOffset(0, 10)
        eff.setColor(QColor(0, 0, 0, 140))
        menu.setGraphicsEffect(eff)

        def _apply_acrylic_now():
            """Apply acrylic now for this workflow."""
            try:
                hwnd = int(menu.winId())
                _try_enable_acrylic(hwnd, self.dark_theme)
            except Exception:
                pass
            QTimer.singleShot(0, lambda m=menu: self._apply_menu_round_mask(m))

        menu.aboutToShow.connect(_apply_acrylic_now)
        menu.installEventFilter(self)
        self._menus_for_mask.add(menu)

    def _menu_css(self) -> str:
        """Handle menu css in MainWindow."""
        if self.dark_theme:
            menu_bg = "rgba(22, 24, 28, 210)"
            menu_border = "rgba(255, 255, 255, 35)"
            menu_sep = "rgba(255, 255, 255, 28)"
            item_hover = "rgba(255, 255, 255, 18)"
            item_press = "rgba(45, 127, 249, 55)"
            text = "#e8eaed"
            disabled = "rgba(232, 234, 237, 120)"
        else:
            menu_bg = "rgba(255, 255, 255, 220)"
            menu_border = "rgba(0, 0, 0, 38)"
            menu_sep = "rgba(0, 0, 0, 26)"
            item_hover = "rgba(0, 0, 0, 6)"
            item_press = "rgba(45, 127, 249, 45)"
            text = "#111111"
            disabled = "rgba(17, 17, 17, 120)"

        return f"""
            QMenu {{
                background: {menu_bg};
                color: {text};
                border: 1px solid {menu_border};
                border-radius: 14px;
                padding: 8px;
                background-clip: padding;
            }}
            QMenu::icon {{
                width: 18px;
                height: 18px;
                padding-left: 2px;
            }}
            QMenu::item {{
                padding: 10px 14px;
                padding-left: 34px;
                border-radius: 10px;
                margin: 2px 2px;
                spacing: 10px;
            }}
            QMenu::item:selected {{
                background: {item_hover};
            }}
            QMenu::item:pressed {{
                background: {item_press};
            }}
            QMenu::item:disabled {{
                color: {disabled};
            }}
            QMenu::separator {{
                height: 1px;
                background: {menu_sep};
                margin: 6px 6px;
            }}
        """

    # ---------------- UI ----------------
    def _load_allowed_equipment_registry(self) -> Dict[str, Dict[str, str]]:
        """Load only records belonging to the approved visible station list."""
        allowed = {
            canonical_station(station).casefold()
            for station in self.data_lists.get("locations", [])
        }
        return {
            serial: record
            for serial, record in load_equipment_registry(EQUIPMENT_REGISTRY).items()
            if canonical_station(record["station"]).casefold() in allowed
        }

    def update_equipment_registry_from_tables(self):
        """Refresh station/model/serial bindings from the approved tables."""
        answer = QMessageBox.question(
            self,
            "Обновление оборудования",
            "Скачать свежие станции, типы оборудования и заводские номера "
            "из утверждённых таблиц?\n\nТекущий реестр будет сохранён как резервная копия.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            result = update_registry(
                EQUIPMENT_REGISTRY,
                self.data_lists["locations"],
            )
            self.equipment_registry = self._load_allowed_equipment_registry()
            current_station = self.cb_location.text()
            self.cb_location.set_items(self.data_lists["locations"])
            self.cb_location.setEditText(current_station)
            self.refresh_serials_for_station(current_station)
        except RegistryUpdateError as exc:
            QMessageBox.critical(
                self,
                "Обновление оборудования",
                f"Реестр не изменён.\n\n{exc}",
            )
            return
        finally:
            QApplication.restoreOverrideCursor()
        message = (
            f"Обновление завершено: {result['count']} записей, "
            f"{result['stations']} станций."
        )
        self.lbl_status.setText(message)
        QMessageBox.information(self, "Обновление оборудования", message)

    def check_for_updates(self):
        """Check for updates for this workflow."""
        manager = getattr(self, "update_manager", None)
        if manager is None:
            QMessageBox.warning(
                self,
                "Проверка обновлений",
                "Модуль обновления ещё не запущен. Повторите попытку через несколько секунд.",
            )
            return
        manager.check_async(manual=True)

    def choose_excel_workbook(self) -> str:
        """Choose excel workbook for this workflow."""
        current = load_excel_path(EXCEL_EXPORT_CONFIG)
        start_at = current if current and os.path.exists(current) else APP_DIR
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите Excel-таблицу для автозаполнения",
            start_at,
            "Excel (*.xlsx)",
        )
        if not path:
            return ""
        try:
            save_excel_path(EXCEL_EXPORT_CONFIG, path)
        except OSError as exc:
            QMessageBox.critical(
                self,
                "Настройка Excel",
                f"Не удалось сохранить путь к таблице:\n{exc}",
            )
            return ""
        self.lbl_status.setText(f"Excel-таблица автозаполнения: {path}")
        return path

    def use_embedded_excel_workbook(self) -> str:
        """Handle use embedded excel workbook in MainWindow."""
        try:
            path = ensure_embedded_excel_workbook(
                EXCEL_TEMPLATE, EMBEDDED_EXCEL_WORKBOOK
            )
            save_excel_path(EXCEL_EXPORT_CONFIG, "")
        except (ExcelExportError, OSError) as exc:
            QMessageBox.critical(
                self,
                "Встроенная Excel-таблица",
                f"Не удалось включить встроенную таблицу:\n{exc}",
            )
            return ""
        self.lbl_status.setText(f"Используется встроенная Excel-таблица: {path}")
        return path

    def active_excel_workbook(self) -> str:
        """Handle active excel workbook in MainWindow."""
        return resolve_excel_workbook(
            EXCEL_EXPORT_CONFIG, EXCEL_TEMPLATE, EMBEDDED_EXCEL_WORKBOOK
        )

    def open_excel_workbook(self):
        """Open excel workbook for this workflow."""
        try:
            path = self.active_excel_workbook()
        except ExcelExportError as exc:
            QMessageBox.critical(self, "Excel-таблица", str(exc))
            return
        open_file(path)

    def set_preview_enabled(self, enabled: bool):
        """Set preview enabled for this workflow."""
        self.app_settings["preview_before_save"] = bool(enabled)
        try:
            save_app_settings(self.app_settings)
        except OSError as error:
            QMessageBox.warning(
                self,
                "Настройки",
                f"Не удалось сохранить настройку предпросмотра:\n{error}",
            )

    def _build_topbar(self) -> QWidget:
        # COLLEAGUE EDIT POINT: add or reorder top-level menus in this method.
        """Build topbar for this workflow."""
        bar = QFrame()
        bar.setObjectName("TopBar")
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(10, 8, 10, 8)
        lay.setSpacing(8)

        brand = QWidget()
        brand_layout = QHBoxLayout(brand)
        brand_layout.setContentsMargins(6, 0, 14, 0)
        brand_layout.setSpacing(9)
        brand_icon = QLabel()
        brand_icon.setPixmap(icon_or_empty(ICON_BRAND).pixmap(48, 48))
        brand_icon.setFixedSize(50, 50)
        brand_title = QLabel("ГЕНЕРАТОР АКТОВ")
        brand_title.setObjectName("BrandTitle")
        brand_layout.addWidget(brand_icon)
        brand_layout.addWidget(brand_title)
        lay.addWidget(brand)

        brand_separator = QFrame()
        brand_separator.setObjectName("BrandSeparator")
        brand_separator.setFrameShape(QFrame.VLine)
        lay.addWidget(brand_separator)

        # --- File menu
        self.menu_file = QMenu(self)
        self.act_check_updates = QAction("Проверить обновления", self)
        self.act_check_updates.triggered.connect(self.check_for_updates)
        self.menu_file.addAction(self.act_check_updates)
        self.act_update_equipment_registry = QAction(
            "Обновить станции и оборудование из таблиц", self
        )
        self.act_update_equipment_registry.triggered.connect(
            self.update_equipment_registry_from_tables
        )
        self.menu_file.addAction(self.act_update_equipment_registry)
        self.menu_file.addSeparator()

        self.act_batch_panel = QAction("Панель массового создания актов", self)
        self.act_batch_panel.setCheckable(True)
        self.act_batch_panel.setChecked(False)
        self.act_batch_panel.toggled.connect(self.set_batch_panel_visible)

        self.act_preview_enabled = QAction("Предпросмотр перед сохранением", self)
        self.act_preview_enabled.setCheckable(True)
        self.act_preview_enabled.setChecked(
            self.app_settings.get("preview_before_save", False)
        )
        self.act_preview_enabled.toggled.connect(self.set_preview_enabled)

        self.act_history_panel = QAction("История актов", self)
        self.act_history_panel.setCheckable(True)
        self.act_history_panel.setChecked(False)
        self.act_history_panel.toggled.connect(self.set_history_panel_visible)

        # COLLEAGUE EDIT POINT: optional panels live in File. Keep their
        # default checked state False unless a product decision changes it.
        # All
        # optional modes remain disabled by default.
        self.menu_file.addAction(self.act_preview_enabled)
        self.menu_file.addAction(self.act_batch_panel)
        self.menu_file.addAction(self.act_history_panel)
        self.menu_file.addSeparator()

        self.act_select_excel = QAction("Выбрать другую Excel-таблицу…", self)
        self.act_select_excel.triggered.connect(self.choose_excel_workbook)
        self.menu_file.addAction(self.act_select_excel)
        self.act_use_embedded_excel = QAction(
            "Использовать встроенную Excel-таблицу", self
        )
        self.act_use_embedded_excel.triggered.connect(
            self.use_embedded_excel_workbook
        )
        self.menu_file.addAction(self.act_use_embedded_excel)
        self.act_open_excel = QAction("Открыть Excel-таблицу", self)
        self.act_open_excel.triggered.connect(self.open_excel_workbook)
        self.menu_file.addAction(self.act_open_excel)
        self.menu_file.addSeparator()

        act_exit = QAction("Выход", self)
        act_exit.triggered.connect(self.close)
        self.menu_file.addAction(act_exit)

        self.btn_file = QToolButton()
        self.btn_file.setText("Файл")
        self.btn_file.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.btn_file.setPopupMode(QToolButton.InstantPopup)
        self.btn_file.setMenu(self.menu_file)

        # --- Links menu (renamed from tables)
        self.menu_links = QMenu(self)
        self.act_edit_links = QAction("Добавить / редактировать ссылки", self)
        self.act_edit_links.triggered.connect(self.edit_links)
        self.menu_links.addAction(self.act_edit_links)
        self.menu_links.addSeparator()

        self.btn_links = QToolButton()
        self.btn_links.setText("Ссылки")
        self.btn_links.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.btn_links.setPopupMode(QToolButton.InstantPopup)
        self.btn_links.setMenu(self.menu_links)

        # --- Lists menu
        self.menu_lists = QMenu(self)

        def add_list(title: str, key: str, icon_path: str):
            """Add list for this workflow."""
            a = QAction(icon_or_empty(icon_path), f"Редактировать: {title}", self)
            a.setIconVisibleInMenu(True)
            a.triggered.connect(lambda: self.open_list_editor(title, key))
            self.menu_lists.addAction(a)

        add_list("Исполнители", "executors", ICON_USER)
        add_list("Станции", "locations", ICON_STATION)
        add_list("Типы оборудования", "models", ICON_EQUIPMENT)
        add_list("Выполненные работы", "work", ICON_SERVICES)
        add_list("Несоответствие", "issues", ICON_ISSUES)
        add_list("Проделанная работа", "done_work", ICON_DONE_WORK)
        add_list("Расходные материалы", "materials", ICON_MATERIALS)

        self.btn_lists = QToolButton()
        self.btn_lists.setText("Списки")
        self.btn_lists.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.btn_lists.setPopupMode(QToolButton.InstantPopup)
        self.btn_lists.setMenu(self.menu_lists)

        # --- Theme button
        self.btn_theme = QToolButton()
        self.btn_theme.setText("Тема")
        self.btn_theme.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.btn_theme.setCheckable(True)
        self.btn_theme.setChecked(True)
        self.btn_theme.clicked.connect(self._toggle_theme_from_button)

        for b in (self.btn_file, self.btn_links, self.btn_lists, self.btn_theme):
            b.setMinimumHeight(40)
            b.setIconSize(QSize(19, 19))

        self.btn_file.setMinimumWidth(104)
        self.btn_links.setMinimumWidth(104)
        self.btn_lists.setMinimumWidth(104)
        self.btn_theme.setMinimumWidth(100)

        lay.addWidget(self.btn_file)
        lay.addWidget(self.btn_links)
        lay.addWidget(self.btn_lists)
        lay.addStretch(1)
        lay.addWidget(self.btn_theme)

        self._beautify_menu(self.menu_file)
        self._beautify_menu(self.menu_links)
        self._beautify_menu(self.menu_lists)

        return bar

    def _build_ui(self):
        """Build ui for this workflow."""
        central = QWidget()
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(18, 16, 18, 16)
        root.setSpacing(14)

        self.topbar = self._build_topbar()
        root.addWidget(self.topbar)
        self._beautify_topbar()

        content = QHBoxLayout()
        content.setSpacing(12)
        root.addLayout(content, 1)

        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(10)
        content.addWidget(left_panel, 1)

        form_card = QFrame()
        form_card.setObjectName("FormCard")
        form_card_layout = QVBoxLayout(form_card)
        form_card_layout.setContentsMargins(18, 14, 18, 14)
        form_card_layout.setSpacing(0)

        form_host = QWidget()
        form = QFormLayout(form_host)
        # COLLEAGUE EDIT POINT: compact row spacing used at 1366px and 2K.
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(8)
        form.setFormAlignment(Qt.AlignTop)
        # Reserve the scrollbar lane for every row equally.  This keeps the
        # date picker, combo arrows and plain fields on one right-hand edge.
        form.setContentsMargins(0, 0, 6, 0)
        self.form_scroll = QScrollArea()
        self.form_scroll.setWidgetResizable(True)
        self.form_scroll.setFrameShape(QFrame.NoFrame)
        self.form_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.form_scroll.setWidget(form_host)
        form_card_layout.addWidget(self.form_scroll, 1)
        left_layout.addWidget(form_card, 1)

        self.ed_date = EditableDateEdit()
        # COLLEAGUE EDIT POINT: form field labels, icons and ordering start here.
        form.addRow(make_form_label("Дата:", ICON_CALENDAR), self.ed_date)

        self.cb_executor = SearchCombo(self.data_lists["executors"], "Исполнители", FILES["executors"])
        form.addRow(make_form_label("Исполнитель:", ICON_USER), self.cb_executor)

        self.cb_location = SearchCombo(self.data_lists["locations"], "Станции", FILES["locations"])
        form.addRow(make_form_label("Станция:", ICON_STATION), self.cb_location)

        self.cb_model = SearchCombo(self.data_lists["models"], "Типы оборудования", FILES["models"])
        self.cb_model.lineEdit().textChanged.connect(lambda _: self.update_template_hint())
        self.cb_model.currentTextChanged.connect(
            lambda _model: self.refresh_serials_for_station(
                self.cb_location.text()
            )
        )
        form.addRow(make_form_label("Тип оборудования:", ICON_EQUIPMENT), self.cb_model)

        self.ed_serial = SerialCombo([])
        self.ed_serial.setPlaceholderText("№ ККТ")
        self.serial_completer = self.ed_serial.completer()
        self.ed_serial.lineEdit().editingFinished.connect(
            self.apply_equipment_for_serial
        )
        # Keep the equipment field in sync immediately when a factory number
        # is selected from the drop-down or entered manually.
        self.ed_serial.currentTextChanged.connect(
            self.apply_equipment_for_serial
        )
        self.ed_serial.textActivated.connect(
            self.apply_equipment_for_serial
        )
        self.ed_serial.activated.connect(self.apply_equipment_for_serial)
        self.cb_location.currentTextChanged.connect(
            self.refresh_serials_for_station
        )
        self.refresh_serials_for_station(self.cb_location.text())
        form.addRow(make_form_label("№ ККТ:", ICON_SERIAL), self.ed_serial)

        self.cb_work = SearchCombo(self.data_lists["work"], "Выполненные работы", FILES["work"])
        form.addRow(make_form_label("Выполненные работы:", ICON_SERVICES), self.cb_work)

        self.cb_issues = SearchCombo(self.data_lists["issues"], "Несоответствие", FILES["issues"])
        form.addRow(make_form_label("Несоответствие:", ICON_ISSUES), self.cb_issues)

        self.cb_done_work = SearchCombo(self.data_lists["done_work"], "Проделанная работа", FILES["done_work"])
        form.addRow(make_form_label("Проделанная работа:", ICON_DONE_WORK), self.cb_done_work)

        self.cb_materials = SearchCombo(self.data_lists["materials"], "Расходные материалы", FILES["materials"])
        form.addRow(make_form_label("Расходные материалы:", ICON_MATERIALS), self.cb_materials)

        self.sp_qty = QuantitySpinBox()
        self.sp_qty.setRange(0, 999999)
        self.sp_qty.setValue(0)
        form.addRow(make_form_label("Количество (цифрами):", ICON_QTY), self.sp_qty)

        self.cb_template = QComboBox()
        self.cb_template.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.cb_template.setMinimumContentsLength(1)
        self.cb_template.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.cb_template.addItem("Авто", None)
        self.cb_template.addItem(
            f"АБП / МКТФ ({os.path.basename(TEMPLATE_TYPE1)})",
            TEMPLATE_TYPE1,
        )
        self.cb_template.addItem(
            f"Валидатор / МИД ({os.path.basename(TEMPLATE_TYPE2)})",
            TEMPLATE_TYPE2,
        )
        self.cb_template.currentIndexChanged.connect(lambda _: self.update_template_hint())
        form.addRow(make_form_label("Выбор шаблона:", ICON_DOCX), self.cb_template)

        # MID/Validator-only fields are intentionally the last row of the form.
        self.service_period_widget = QWidget()
        service_period_layout = QHBoxLayout(self.service_period_widget)
        service_period_layout.setContentsMargins(0, 0, 0, 0)
        service_period_layout.setSpacing(8)
        self.ed_service_from = EditableTimeEdit()
        self.ed_service_from.setToolTip(
            "Введите только время цифрами, например 0830. Дата берётся из поля «Дата»."
        )
        self.ed_service_to = EditableTimeEdit()
        self.ed_service_to.setToolTip(
            "Введите только время цифрами, например 1045. Дата берётся из поля «Дата»."
        )
        service_period_layout.addWidget(self.ed_service_from, 1)
        service_period_separator = QLabel("по")
        service_period_separator.setAlignment(Qt.AlignCenter)
        service_period_layout.addWidget(service_period_separator)
        service_period_layout.addWidget(self.ed_service_to, 1)
        self.service_period_label = make_form_label("Период услуги:", ICON_CALENDAR)
        form.addRow(self.service_period_label, self.service_period_widget)
        self.update_template_hint()

        action_bar = QFrame()
        # COLLEAGUE EDIT POINT: fixed bottom action layout and button ordering.
        action_bar.setObjectName("ActionBar")
        self.action_bar = action_bar
        action_layout = QHBoxLayout(action_bar)
        action_layout.setContentsMargins(10, 8, 10, 8)
        action_layout.setSpacing(10)
        self.btn_clear = QPushButton("Очистить")
        self.btn_clear.setProperty("secondary", True)
        self.btn_clear.setMinimumHeight(38)
        self.btn_clear.clicked.connect(self.clear_form)
        self.btn_batch_add = QPushButton("Добавить в очередь")
        self.btn_batch_add.setMinimumHeight(38)
        self.btn_batch_add.clicked.connect(self.add_current_act_to_batch)
        self.btn_batch_add.hide()
        self.btn_open_output = QPushButton("Открыть папку с актами")
        self.btn_open_output.setProperty("secondary", True)
        self.btn_open_output.setMinimumHeight(38)
        self.btn_open_output.clicked.connect(lambda: open_folder(OUTPUT_DIR))
        self.btn_surname_act = QPushButton("Акт по фамилии")
        self.btn_surname_act.setProperty("secondary", True)
        self.btn_surname_act.setMinimumHeight(38)
        self.btn_surname_act.setToolTip(
            "Создать чистый акт, заполнив только фамилию исполнителя"
        )
        surname_menu = QMenu(self.btn_surname_act)
        surname_type1 = surname_menu.addAction("АБП / МКТФ")
        surname_type1.triggered.connect(
            lambda: self.create_surname_only_act(TEMPLATE_TYPE1)
        )
        surname_type2 = surname_menu.addAction("Валидатор / МИД")
        surname_type2.triggered.connect(
            lambda: self.create_surname_only_act(TEMPLATE_TYPE2)
        )
        self.btn_surname_act.setMenu(surname_menu)
        for button in (
            self.btn_clear, self.btn_batch_add, self.btn_open_output,
            self.btn_surname_act,
        ):
            button.setIconSize(QSize(18, 18))

        self.btn_create = QPushButton("Создать акт")
        self.btn_create.setProperty("primaryAction", True)
        self.btn_create.setMinimumSize(300, 46)
        self.btn_create.setMaximumWidth(380)
        create_shadow = QGraphicsDropShadowEffect(self.btn_create)
        create_shadow.setBlurRadius(24)
        create_shadow.setOffset(0, 6)
        create_shadow.setColor(QColor(37, 99, 235, 105))
        self.btn_create.setGraphicsEffect(create_shadow)
        self.btn_create.clicked.connect(self.create_act)
        action_layout.addWidget(self.btn_clear)
        action_layout.addStretch(1)
        action_layout.addWidget(self.btn_surname_act)
        action_layout.addWidget(self.btn_create, 0, Qt.AlignHCenter)
        action_layout.addStretch(1)
        action_layout.addWidget(self.btn_batch_add)
        action_layout.addWidget(self.btn_open_output)
        left_layout.addWidget(action_bar)

        self.lbl_status = QLabel("")
        self.lbl_status.setObjectName("StatusLabel")
        self.lbl_status.setWordWrap(True)
        left_layout.addWidget(self.lbl_status)

        self.batch_panel = QFrame()
        self.batch_panel.setObjectName("BatchPanel")
        self.batch_panel.setMinimumWidth(520)
        batch_layout = QVBoxLayout(self.batch_panel)
        batch_layout.setContentsMargins(12, 12, 12, 12)
        batch_layout.setSpacing(8)

        batch_header = QHBoxLayout()
        batch_title = QLabel("Массовое создание")
        batch_title.setObjectName("BatchTitle")
        self.batch_count_label = QLabel("В очереди: 0")
        self.batch_count_label.setObjectName("BatchCount")
        batch_header.addWidget(batch_title)
        batch_header.addStretch(1)
        batch_header.addWidget(self.batch_count_label)
        batch_layout.addLayout(batch_header)

        batch_hint = QLabel(
            "Добавляйте заполненную форму слева. Двойной щелчок по строке — редактирование."
        )
        batch_hint.setWordWrap(True)
        batch_layout.addWidget(batch_hint)

        self.batch_table = QTableWidget(0, 6)
        self.batch_table.setHorizontalHeaderLabels(
            ("Дата", "Станция", "Тип", "№ ККТ", "Исполнитель", "Статус")
        )
        self.batch_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.batch_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.batch_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.batch_table.setAlternatingRowColors(True)
        self.batch_table.setShowGrid(False)
        self.batch_table.verticalHeader().setVisible(False)
        self.batch_table.verticalHeader().setDefaultSectionSize(38)
        header = self.batch_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        self.batch_table.cellDoubleClicked.connect(self.edit_batch_act)
        batch_layout.addWidget(self.batch_table, 1)

        self.batch_progress = QProgressBar()
        self.batch_progress.setTextVisible(True)
        self.batch_progress.hide()
        batch_layout.addWidget(self.batch_progress)

        batch_buttons = QHBoxLayout()
        self.btn_batch_duplicate = QPushButton("Дублировать")
        self.btn_batch_duplicate.setProperty("secondary", True)
        self.btn_batch_duplicate.clicked.connect(self.duplicate_selected_batch_act)
        self.btn_batch_remove = QPushButton("Удалить выбранные")
        self.btn_batch_remove.setProperty("dangerSecondary", True)
        self.btn_batch_remove.clicked.connect(self.remove_selected_batch_acts)
        self.btn_batch_clear = QPushButton("Очистить очередь")
        self.btn_batch_clear.setProperty("dangerSecondary", True)
        self.btn_batch_clear.clicked.connect(self.clear_batch_acts)
        for button in (
            self.btn_batch_duplicate, self.btn_batch_remove, self.btn_batch_clear
        ):
            button.setIconSize(QSize(18, 18))
        batch_buttons.addWidget(self.btn_batch_duplicate)
        batch_buttons.addWidget(self.btn_batch_remove)
        batch_buttons.addWidget(self.btn_batch_clear)
        batch_layout.addLayout(batch_buttons)

        self.btn_batch_create = QPushButton("Создать все акты")
        self.btn_batch_create.setIconSize(QSize(20, 20))
        self.btn_batch_create.setMinimumHeight(44)
        self.btn_batch_create.clicked.connect(self.create_batch_acts)
        batch_layout.addWidget(self.btn_batch_create)

        self.batch_edit_row = None
        self.batch_panel.hide()
        content.addWidget(self.batch_panel)

        self.history_panel = self._build_history_panel()
        self.history_panel.hide()
        content.addWidget(self.history_panel)

    def _build_history_panel(self) -> QWidget:
        """Build history panel for this workflow."""
        panel = QFrame()
        panel.setObjectName("HistoryPanel")
        panel.setMinimumWidth(620)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        header = QHBoxLayout()
        title = QLabel("История актов")
        title.setObjectName("BatchTitle")
        self.history_count_label = QLabel("Актов: 0")
        self.history_count_label.setObjectName("BatchCount")
        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(self.history_count_label)
        layout.addLayout(header)

        filters = QHBoxLayout()
        self.history_search = QLineEdit()
        self.history_search.setPlaceholderText("Поиск по станции, № ККТ, исполнителю или имени файла…")
        self.history_search.textChanged.connect(self.filter_history)
        self.history_station = QComboBox()
        self.history_station.addItem("Все станции")
        self.history_station.currentTextChanged.connect(self.filter_history)
        filters.addWidget(self.history_search, 1)
        filters.addWidget(self.history_station)
        layout.addLayout(filters)

        self.history_table = QTableWidget(0, 6)
        self.history_table.setHorizontalHeaderLabels(
            ("Дата", "Станция", "Тип", "№ ККТ", "Исполнитель", "Акт")
        )
        self.history_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.history_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.history_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.history_table.setAlternatingRowColors(True)
        self.history_table.setShowGrid(False)
        self.history_table.verticalHeader().setVisible(False)
        self.history_table.verticalHeader().setDefaultSectionSize(38)
        history_header = self.history_table.horizontalHeader()
        history_header.setSectionResizeMode(QHeaderView.Interactive)
        self.history_table.setColumnWidth(0, 86)
        self.history_table.setColumnWidth(2, 82)
        self.history_table.setColumnWidth(3, 105)
        self.history_table.setColumnWidth(5, 150)
        history_header.setSectionResizeMode(1, QHeaderView.Stretch)
        history_header.setSectionResizeMode(4, QHeaderView.Stretch)
        self.history_table.cellDoubleClicked.connect(self.open_history_row)
        layout.addWidget(self.history_table, 1)

        buttons = QHBoxLayout()
        self.btn_history_refresh = QPushButton("Обновить")
        self.btn_history_refresh.setProperty("secondary", True)
        self.btn_history_refresh.clicked.connect(self.refresh_history)
        self.btn_history_load = QPushButton("Загрузить в форму")
        self.btn_history_load.setProperty("secondary", True)
        self.btn_history_load.clicked.connect(self.load_selected_history)
        self.btn_history_clear = QPushButton("Очистить историю")
        self.btn_history_clear.setProperty("dangerSecondary", True)
        self.btn_history_clear.clicked.connect(self.clear_history)
        self.btn_history_delete = QPushButton("Удалить выбранный")
        self.btn_history_delete.setProperty("dangerSecondary", True)
        self.btn_history_delete.clicked.connect(self.delete_selected_history)
        self.btn_history_open = QPushButton("Открыть акт DOCX")
        self.btn_history_open.clicked.connect(self.open_selected_history)
        for button in (
            self.btn_history_refresh, self.btn_history_load,
            self.btn_history_delete, self.btn_history_clear,
        ):
            button.setIconSize(QSize(18, 18))
        self.btn_history_open.setIconSize(QSize(18, 18))
        buttons.addWidget(self.btn_history_refresh)
        buttons.addWidget(self.btn_history_load)
        buttons.addWidget(self.btn_history_delete)
        buttons.addWidget(self.btn_history_clear)
        buttons.addStretch(1)
        buttons.addWidget(self.btn_history_open)
        layout.addLayout(buttons)
        return panel

    # ---------------- Theme / Styles ----------------
    def _toggle_theme_from_button(self):
        """Toggle theme from button for this workflow."""
        self.dark_theme = self.btn_theme.isChecked()
        self.apply_theme(self.dark_theme)

    def apply_theme(self, dark: bool):
        """Apply theme for this workflow."""
        QApplication.instance().setStyle("Fusion")

        pal = QPalette()
        if not dark:
            pal.setColor(QPalette.Window, QColor("#f5f6f7"))
            pal.setColor(QPalette.WindowText, QColor("#111111"))
            pal.setColor(QPalette.Base, QColor("#ffffff"))
            pal.setColor(QPalette.Text, QColor("#111111"))
            pal.setColor(QPalette.Button, QColor("#ffffff"))
            pal.setColor(QPalette.ButtonText, QColor("#111111"))
            pal.setColor(QPalette.Highlight, QColor("#2d7ff9"))
            pal.setColor(QPalette.HighlightedText, QColor("#ffffff"))
            border = "#cfd3d7"
            bg = "#ffffff"
            dd_border = "#e3e6ea"
            popup_bg = "#ffffff"
            popup_fg = "#111111"
            spin_up = SPIN_UP_LIGHT_PNG
            spin_down = SPIN_DOWN_LIGHT_PNG
        else:
            pal.setColor(QPalette.Window, QColor("#0b1016"))
            pal.setColor(QPalette.WindowText, QColor("#edf3f8"))
            pal.setColor(QPalette.Base, QColor("#0f161e"))
            pal.setColor(QPalette.Text, QColor("#edf3f8"))
            pal.setColor(QPalette.Button, QColor("#121922"))
            pal.setColor(QPalette.ButtonText, QColor("#edf3f8"))
            pal.setColor(QPalette.Highlight, QColor("#3b82f6"))
            pal.setColor(QPalette.HighlightedText, QColor("#ffffff"))
            border = "#2d3a48"
            bg = "#0f161e"
            dd_border = "#2d3a48"
            popup_bg = "#121922"
            popup_fg = "#edf3f8"
            spin_up = SPIN_UP_DARK_PNG
            spin_down = SPIN_DOWN_DARK_PNG

        if dark:
            # COLLEAGUE EDIT POINT: dark theme palette and field colors.
            tb_bg = "rgba(18, 25, 34, 235)"
            tb_border = "rgba(255, 255, 255, 22)"
            btn_bg = "rgba(255, 255, 255, 10)"
            btn_hover = "rgba(255, 255, 255, 18)"
            btn_press = "rgba(255, 255, 255, 26)"
            btn_checked = "rgba(45, 127, 249, 40)"
            btn_checked_border = "rgba(45, 127, 249, 120)"
        else:
            # COLLEAGUE EDIT POINT: light theme palette and field colors.
            tb_bg = "rgba(255, 255, 255, 235)"
            tb_border = "rgba(0, 0, 0, 18)"
            btn_bg = "rgba(0, 0, 0, 4)"
            btn_hover = "rgba(0, 0, 0, 7)"
            btn_press = "rgba(0, 0, 0, 10)"
            btn_checked = "rgba(45, 127, 249, 22)"
            btn_checked_border = "rgba(45, 127, 249, 120)"

        QApplication.instance().setPalette(pal)

        arrow_png = ICON_ARROW_DOWN if os.path.exists(ICON_ARROW_DOWN) else ""
        arrow_png_qss = arrow_png.replace("\\", "/") if arrow_png else ""
        self.sp_qty.set_arrow_icons(spin_up, spin_down)

        down_arrow_rule = f"""
        QComboBox::down-arrow {{
            image: url("{arrow_png_qss}");
            width: 18px;
            height: 18px;
            subcontrol-position: center;
            subcontrol-origin: content;
        }}
        """ if arrow_png_qss else ""

        self.setStyleSheet(f"""
            #TopBar {{
                background: {tb_bg};
                border: 1px solid {tb_border};
                border-radius: 16px;
                padding: 2px;
            }}
            #BrandTitle {{
                color: #65aefb;
                font-size: 12px;
                font-weight: 700;
                letter-spacing: 1.15px;
            }}
            #BrandSeparator {{
                color: {tb_border};
                margin: 4px 2px;
            }}

            #FormCard, #ActionBar {{
                background: {tb_bg};
                border: 1px solid {tb_border};
                border-radius: 16px;
            }}
            #FormTitle {{
                font-size: 20px;
                font-weight: 750;
            }}
            #FormSubtitle {{
                color: #8190a3;
                font-size: 12px;
            }}
            #StatusLabel {{
                color: #8190a3;
                padding: 0 4px;
                min-height: 18px;
            }}
            QScrollArea {{ background: transparent; }}
            QScrollArea > QWidget > QWidget {{ background: transparent; }}

            #BatchPanel, #HistoryPanel {{
                background: {tb_bg};
                border: 1px solid {tb_border};
                border-radius: 14px;
            }}
            #BatchTitle {{
                font-size: 16px;
                font-weight: 700;
            }}
            #BatchCount {{
                color: #2d7ff9;
                font-weight: 700;
            }}

            QToolButton {{
                color: {popup_fg};
                background: {btn_bg};
                border: 1px solid rgba(255,255,255,0);
                padding: 8px 12px;
                border-radius: 10px;
                font-weight: 600;
            }}
            QToolButton:hover {{ background: {btn_hover}; }}
            QToolButton:pressed {{ background: {btn_press}; }}
            QToolButton:checked {{
                background: {btn_checked};
                border: 1px solid {btn_checked_border};
            }}
            QToolButton::menu-indicator {{
                image: none;
                width: 0px;
            }}

            {self._menu_css()}

            QComboBox {{
                padding: 5px 34px 5px 12px;
                border-radius: 11px;
                border: 1px solid {border};
                min-height: 30px;
                background: {bg};
            }}
            QComboBox:hover, QLineEdit:hover, QSpinBox:hover, QDateEdit:hover {{
                border-color: #4b6480;
            }}
            QComboBox::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 30px;
                border-left: 1px solid {dd_border};
            }}
            {down_arrow_rule}

            QComboBox QLineEdit {{
                padding-right: 8px;
                background: transparent;
                border: none;
            }}

            QLineEdit, QSpinBox, QDateEdit {{
                padding: 5px 11px;
                border-radius: 11px;
                border: 1px solid {border};
                min-height: 30px;
                background: {bg};
            }}
            QLineEdit:focus, QSpinBox:focus, QDateEdit:focus, QComboBox:focus {{
                border: 1px solid #3b82f6;
            }}
            #DateField, #TimeField {{
                background: {bg};
                border: 1px solid {border};
                border-radius: 11px;
            }}
            #DateField:hover, #TimeField:hover {{
                border-color: #4b6480;
            }}
            #DateEditor, #TimeEditor {{
                background: transparent;
                border: none;
                border-radius: 10px;
                padding: 5px 11px;
                min-height: 30px;
            }}
            #DateEditor:focus, #TimeEditor:focus {{
                border: none;
            }}
            #DatePickerButton, #TimePickerButton {{
                background: rgba(59, 130, 246, 22);
                border-left: 1px solid {dd_border};
                border-top: none;
                border-right: none;
                border-bottom: none;
                border-top-right-radius: 10px;
                border-bottom-right-radius: 10px;
                border-top-left-radius: 0px;
                border-bottom-left-radius: 0px;
                padding: 0px;
            }}
            #DatePickerButton:hover, #TimePickerButton:hover {{
                background: rgba(59, 130, 246, 55);
            }}
            #TimePickerCaption {{
                color: #8190a3;
                font-size: 11px;
                font-weight: 600;
            }}
            #TimePickerSeparator {{
                min-width: 12px;
                padding-top: 19px;
                font-size: 20px;
                font-weight: 700;
            }}
            #TimePickerSpinField {{
                background: {bg};
                border: 1px solid {border};
                border-radius: 10px;
            }}
            #TimePickerSpinEditor {{
                background: transparent;
                border: none;
                border-radius: 9px;
                min-height: 0px;
                padding: 0 6px;
                font-size: 15px;
                font-weight: 700;
            }}
            #TimePickerUpButton, #TimePickerDownButton {{
                background: transparent;
                border-left: 1px solid {dd_border};
                border-top: none;
                border-right: none;
                border-bottom: none;
                border-radius: 0px;
                padding: 0px;
            }}
            #TimePickerUpButton {{
                border-bottom: 1px solid {dd_border};
                border-top-right-radius: 9px;
            }}
            #TimePickerDownButton {{
                border-bottom-right-radius: 9px;
            }}
            #TimePickerUpButton:hover, #TimePickerDownButton:hover {{
                background: rgba(59, 130, 246, 35);
            }}

            QComboBox QAbstractItemView {{
                background: {popup_bg};
                color: {popup_fg};
                selection-background-color: #2d7ff9;
                selection-color: #ffffff;
            }}

            #QuantityField {{
                background: {bg};
                border: 1px solid {border};
                border-radius: 11px;
            }}
            #QuantityField:hover {{
                border-color: #4b6480;
            }}
            #QuantityEditor {{
                background: transparent;
                border: none;
                border-radius: 10px;
                padding: 5px 11px;
                min-height: 30px;
            }}
            #QuantityEditor:focus {{
                border: none;
            }}
            #QuantityUpButton, #QuantityDownButton {{
                background: transparent;
                border-left: 1px solid {dd_border};
                border-top: none;
                border-right: none;
                border-bottom: none;
                border-radius: 0px;
                padding: 0px 0px 0px 1px;
            }}
            #QuantityUpButton {{
                border-bottom: 1px solid {dd_border};
                border-top-right-radius: 10px;
            }}
            #QuantityDownButton {{
                border-bottom-right-radius: 10px;
            }}
            #QuantityUpButton:hover, #QuantityDownButton:hover {{
                background: rgba(59, 130, 246, 35);
            }}

            QPushButton {{
                background: #3b82f6;
                color: #ffffff;
                font-weight: 700;
                border: none;
                border-radius: 11px;
                padding: 10px 14px;
            }}
            QPushButton:hover {{ background: #2f73df; }}
            QPushButton:pressed {{ background: #2563c7; }}
            QPushButton[primaryAction="true"] {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #2f73ed, stop:0.52 #3b82f6, stop:1 #2563eb);
                border: 1px solid rgba(147, 197, 253, 105);
                border-radius: 14px;
                min-height: 44px;
                font-size: 14px;
                font-weight: 750;
                letter-spacing: 0.25px;
                padding: 0 30px;
            }}
            QPushButton[primaryAction="true"]:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #3b82f6, stop:0.52 #4b8df8, stop:1 #2f6fe4);
                border-color: rgba(191, 219, 254, 155);
            }}
            QPushButton[primaryAction="true"]:pressed {{
                background: #245fc7;
                border-color: rgba(147, 197, 253, 125);
            }}
            QPushButton[secondary="true"] {{
                background: {btn_bg};
                color: {popup_fg};
                border: 1px solid {tb_border};
            }}
            QPushButton[secondary="true"]:hover {{ background: {btn_hover}; }}
            QPushButton[dangerSecondary="true"] {{
                background: {btn_bg};
                color: #ef6464;
                border: 1px solid #a94444;
            }}
            QPushButton[dangerSecondary="true"]:hover {{ background: {btn_hover}; }}
            QTableWidget {{
                background: {bg};
                alternate-background-color: {tb_bg};
                border: 1px solid {border};
                border-radius: 10px;
                gridline-color: transparent;
                selection-background-color: rgba(59, 130, 246, 72);
                selection-color: {popup_fg};
            }}
            QTableWidget::item {{
                padding: 7px 8px;
                border-bottom: 1px solid {tb_border};
            }}
            QTableWidget::item:selected {{
                background: rgba(59, 130, 246, 72);
                color: {popup_fg};
            }}
            QHeaderView::section {{
                background: {tb_bg};
                color: {popup_fg};
                border: none;
                border-bottom: 1px solid {border};
                padding: 9px 8px;
                font-weight: 600;
            }}
            QScrollBar:vertical {{
                background: transparent;
                width: 10px;
                margin: 3px;
            }}
            QScrollBar::handle:vertical {{
                background: rgba(120, 140, 165, 95);
                min-height: 32px;
                border-radius: 4px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0px;
            }}
        """)

    # ---------------- Menus data ----------------
    def refresh_links_menu(self):
        """Refresh links menu for this workflow."""
        self.menu_links.clear()
        self.menu_links.addAction(self.act_edit_links)
        self.menu_links.addSeparator()

        self.links = load_links()
        if not self.links:
            a = QAction("(ссылки не добавлены)", self)
            a.setEnabled(False)
            self.menu_links.addAction(a)
            return

        default_icon = icon_or_empty(ICON_XLS) if os.path.exists(ICON_XLS) else QIcon()

        for t in self.links:
            name = (t.get("name") or "").strip() or "Ссылка"
            link = (t.get("link") or "").strip()

            ico_path = (t.get("icon") or "").strip()
            ico = icon_or_empty(ico_path) if ico_path else default_icon

            a = QAction(ico, name, self)
            a.setIconVisibleInMenu(True)
            a.triggered.connect(partial(open_any, link))
            self.menu_links.addAction(a)

    def open_list_editor(self, title: str, key: str):
        """Open list editor for this workflow."""
        dlg = ListEditorDialog(title, FILES[key], self)
        dlg.exec()
        self.data_lists[key] = read_lines(FILES[key])
        combos = {
            "executors": self.cb_executor,
            "locations": self.cb_location,
            "models": self.cb_model,
            "work": self.cb_work,
            "issues": self.cb_issues,
            "done_work": self.cb_done_work,
            "materials": self.cb_materials,
        }
        combos[key].set_items(self.data_lists[key])
        if key == "models":
            self.update_template_hint()

    def edit_links(self):
        """Edit links for this workflow."""
        dlg = LinksEditorDialog(load_links(), self)
        if dlg.exec() == QDialog.Accepted:
            save_links(dlg.links)
            self.refresh_links_menu()
            QMessageBox.information(self, "Ссылки", "Ссылки сохранены.")

    # ---------------- Act generation ----------------
    def _canonical_list_value(self, key: str, value: str) -> str:
        """Reuse list capitalization when the registry contains the same value."""
        wanted = str(value or "").strip().casefold()
        for item in self.data_lists.get(key, []):
            if item.strip().casefold() == wanted:
                return item
        return str(value or "").strip()

    def refresh_serials_for_station(self, station: str = ""):
        """Show numbers matching both the selected station and equipment."""
        station_key = normalize_registry_text(station).casefold()
        model_key = canonical_equipment_model(self.cb_model.text()).casefold()
        current_serial = self.ed_serial.text()
        current_record = self.equipment_registry.get(
            normalize_serial_key(current_serial)
        )
        if current_record and (
            normalize_registry_text(current_record["station"]).casefold()
            != station_key
            or (
                model_key
                and canonical_equipment_model(current_record["model"]).casefold()
                != model_key
            )
        ):
            current_serial = ""

        serials = []
        if station_key:
            serials = sorted(
                (
                    record["serial"]
                    for record in self.equipment_registry.values()
                    if normalize_registry_text(record["station"]).casefold()
                    == station_key
                    and (
                        not model_key
                        or canonical_equipment_model(record["model"]).casefold()
                        == model_key
                    )
                ),
                key=lambda value: (len(value), value),
            )
        self.ed_serial.set_items(serials, current_text=current_serial)

    def apply_equipment_for_serial(self, selected_value=None, *_args) -> bool:
        """Fill equipment and station for an exact known factory number."""
        # textActivated provides the selected value directly. Using it avoids
        # reading the previous edit text while the drop-down is still closing.
        serial = (
            selected_value
            if isinstance(selected_value, str)
            else self.ed_serial.text()
        )
        key = normalize_serial_key(serial)
        record = self.equipment_registry.get(key)
        if record is None:
            return False
        model = self._canonical_list_value("models", record["model"])
        station = self._canonical_list_value("locations", record["station"])
        # Change both dependent fields as one transaction. If their signals are
        # allowed to run between these two assignments, the serial filter can
        # briefly see a new model with the old station (or vice versa) and clear
        # the number that the user has just selected.
        model_blocker = QSignalBlocker(self.cb_model)
        station_blocker = QSignalBlocker(self.cb_location)
        try:
            self.cb_model.setEditText(model)
            self.cb_location.setEditText(station)
        finally:
            del station_blocker
            del model_blocker
        self.refresh_serials_for_station(station)
        self.update_template_hint()
        self.lbl_status.setText(
            f"По № {record['serial']} определено: {model}, {station}."
        )
        return True

    def pick_template_for_equipment_type(self, equipment_type: str) -> str:
        """
        ✅ Один тип (TEMPLATE_TYPE1) для:
        МКТФ, АБП-БН, АБП-09, АБП-09-М3, АН-15, БПА-20-БН1
        """
        t = (equipment_type or "").upper().replace(" ", "")

        # Все варианты, которые должны идти в TEMPLATE_TYPE1
        type1_tokens = (
            "МКТФ", "MKTF",
            "АБП",          # покроет АБП-БН, АБП-09, АБП-09-М3 и т.п.
            "АН-15", "АН15",
            "БПА",          # покроет БПА-20-БН1
        )

        if any(tok in t for tok in type1_tokens):
            return TEMPLATE_TYPE1

        return TEMPLATE_TYPE2

    def selected_template(self, equipment_type: str) -> str:
        """Возвращает выбранный вручную шаблон или результат автовыбора."""
        manual_template = self.cb_template.currentData()
        if manual_template:
            return manual_template
        return self.pick_template_for_equipment_type(equipment_type)

    def update_template_hint(self):
        """Update template hint for this workflow."""
        tpl = self.selected_template(self.cb_model.text())
        self.cb_template.setToolTip(f"Будет использован: {os.path.basename(tpl)}")
        if hasattr(self, "service_period_widget"):
            visible = is_validator_template(tpl)
            self.service_period_label.setVisible(visible)
            self.service_period_widget.setVisible(visible)

    def set_batch_panel_visible(self, visible: bool):
        """Set batch panel visible for this workflow."""
        if visible and getattr(self, "act_history_panel", None) is not None:
            self.act_history_panel.setChecked(False)
        panel = getattr(self, "batch_panel", None)
        if panel is not None:
            panel.setVisible(bool(visible))
            self.btn_batch_add.setVisible(bool(visible))
            if visible:
                self.resize(max(self.width(), 1450), max(self.height(), 820))

    def set_history_panel_visible(self, visible: bool):
        """Set history panel visible for this workflow."""
        if visible and getattr(self, "act_batch_panel", None) is not None:
            self.act_batch_panel.setChecked(False)
        panel = getattr(self, "history_panel", None)
        if panel is not None:
            panel.setVisible(bool(visible))
            if visible:
                self.refresh_history()
                self.resize(max(self.width(), 1450), max(self.height(), 820))

    def refresh_history(self):
        """Refresh history for this workflow."""
        records = load_history_records(HISTORY_DIR, APP_DIR, OUTPUT_DIR)
        self.history_table.setRowCount(0)
        stations = {"Все станции"}
        for record in records:
            data = dict(record.get("data") or {})
            row = self.history_table.rowCount()
            self.history_table.insertRow(row)
            values = (
                data.get("date", ""), data.get("location", ""),
                data.get("model", ""), data.get("serial", ""),
                data.get("executor", ""), record.get("document_name", ""),
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value or "—"))
                if column == 0:
                    item.setData(Qt.UserRole, dict(record))
                self.history_table.setItem(row, column, item)
            if data.get("location"):
                stations.add(str(data["location"]))
        with QSignalBlocker(self.history_station):
            current = self.history_station.currentText()
            self.history_station.clear()
            self.history_station.addItems(sorted(stations, key=lambda value: (value != "Все станции", value)))
            index = self.history_station.findText(current)
            self.history_station.setCurrentIndex(index if index >= 0 else 0)
        self.history_count_label.setText(f"Актов: {len(records)}")
        self.filter_history()

    def history_row_record(self, row: int) -> Dict[str, object]:
        """Handle history row record in MainWindow."""
        item = self.history_table.item(row, 0)
        return dict(item.data(Qt.UserRole)) if item is not None else {}

    def selected_history_row(self) -> int | None:
        """Return history row for this workflow."""
        rows = self.history_table.selectionModel().selectedRows()
        return rows[0].row() if rows else None

    def filter_history(self):
        """Filter history for this workflow."""
        search = (self.history_search.text() or "").strip().casefold()
        station = self.history_station.currentText()
        for row in range(self.history_table.rowCount()):
            values = [
                self.history_table.item(row, column).text()
                for column in range(self.history_table.columnCount())
                if self.history_table.item(row, column) is not None
            ]
            row_station = self.history_table.item(row, 1).text()
            visible = (not search or search in " ".join(values).casefold()) and (
                station == "Все станции" or station == row_station
            )
            self.history_table.setRowHidden(row, not visible)

    def load_selected_history(self):
        """Load selected history for this workflow."""
        row = self.selected_history_row()
        if row is None:
            QMessageBox.information(self, "История актов", "Выберите акт.")
            return
        record = self.history_row_record(row)
        data = dict(record.get("data") or {})
        if not data:
            return
        self.apply_act_data_to_form(data)
        self.lbl_status.setText(f"Данные загружены из истории: {record.get('document_name', '')}")

    def open_history_row(self, row: int, _column: int):
        """Open history row for this workflow."""
        self.history_table.selectRow(row)
        self.open_selected_history()

    def open_selected_history(self):
        """Open selected history for this workflow."""
        row = self.selected_history_row()
        if row is None:
            QMessageBox.information(self, "История актов", "Выберите акт.")
            return
        record = self.history_row_record(row)
        document = str(record.get("document", ""))
        if not document or not os.path.isfile(document):
            QMessageBox.warning(self, "История актов", "Файл DOCX не найден.")
            return
        open_file(document)

    def clear_history(self):
        """Clear history for this workflow."""
        if self.history_table.rowCount() == 0:
            return
        answer = QMessageBox.question(
            self,
            "Очистить историю",
            "Удалить все записи из истории?\n\nСозданные файлы DOCX в папке Acts останутся.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        try:
            removed = clear_history_records(HISTORY_DIR)
        except OSError as error:
            QMessageBox.critical(
                self, "Очистить историю", f"Не удалось очистить историю:\n{error}"
            )
            return
        self.refresh_history()
        self.lbl_status.setText(
            f"История очищена: {removed}. Файлы DOCX в папке Acts сохранены."
        )

    def delete_selected_history(self):
        """Delete selected history for this workflow."""
        row = self.selected_history_row()
        if row is None:
            QMessageBox.information(
                self, "Удалить акт", "Выберите строку с актом для удаления."
            )
            return
        record = self.history_row_record(row)
        document = str(record.get("document", ""))
        if not document or not os.path.isfile(document):
            self.refresh_history()
            QMessageBox.warning(self, "Удалить акт", "Файл DOCX уже не найден.")
            return
        document_path = os.path.realpath(document)
        output_path = os.path.realpath(OUTPUT_DIR)
        if os.path.dirname(document_path) != output_path:
            QMessageBox.critical(
                self,
                "Удалить акт",
                "Удаление отменено: выбранный файл находится вне папки Acts.",
            )
            return
        answer = QMessageBox.question(
            self,
            "Удалить выбранный акт",
            f"Отправить файл «{os.path.basename(document_path)}» в Корзину?\n\n"
            "Остальные акты и очередь не изменятся.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        try:
            move_file_to_recycle_bin(document_path)
            metadata = str(record.get("metadata_path", ""))
            if metadata:
                metadata_path = os.path.realpath(metadata)
                if (
                    os.path.dirname(metadata_path) == os.path.realpath(HISTORY_DIR)
                    and metadata_path.lower().endswith(".json")
                    and os.path.isfile(metadata_path)
                ):
                    os.remove(metadata_path)
        except OSError as error:
            QMessageBox.critical(
                self, "Удалить акт", f"Не удалось удалить выбранный акт:\n{error}"
            )
            return
        self.refresh_history()
        self.lbl_status.setText(
            f"Акт перемещён в Корзину: {os.path.basename(document_path)}"
        )

    def clear_form(self):
        """Clear form for this workflow."""
        self.ed_date.setDate(QDate.currentDate())
        self.ed_service_from.clear()
        self.ed_service_to.clear()
        for combo in (
            self.cb_executor, self.cb_location, self.cb_model, self.cb_work,
            self.cb_issues, self.cb_done_work, self.cb_materials,
        ):
            combo.setEditText("")
        self.ed_serial.clear()
        self.sp_qty.setValue(0)
        self.cb_template.setCurrentIndex(0)
        self.lbl_status.setText("Форма очищена.")

    def current_act_data(self) -> Dict[str, object]:
        """Return act data for this workflow."""
        self.apply_equipment_for_serial()
        model = self.cb_model.text()
        executor = self.cb_executor.text()
        if is_excluded_executor(executor):
            executor = ""
            self.cb_executor.setEditText("")
            self.lbl_status.setText("Исполнитель исключён из списка.")
        return {
            "date": (self.ed_date.text() or "").strip(),
            "service_from": normalize_service_time_input(
                self.ed_service_from.text()
            ),
            "service_to": normalize_service_time_input(
                self.ed_service_to.text()
            ),
            "executor": executor,
            "location": self.cb_location.text(),
            "model": model,
            "serial": (self.ed_serial.text() or "").strip(),
            "work": self.cb_work.text(),
            "issues": self.cb_issues.text(),
            "done_work": self.cb_done_work.text(),
            "materials": self.cb_materials.text(),
            "qty": int(self.sp_qty.value()),
            "template": self.selected_template(model),
        }

    @staticmethod
    def batch_act_caption(data: Dict[str, object]) -> str:
        """Handle batch act caption in MainWindow."""
        return (
            f"{data['date']} | {data['location'] or 'Без станции'} | "
            f"{data['model'] or 'Без типа'} | № {data['serial'] or '—'} | "
            f"{data['executor'] or 'Без исполнителя'}"
        )

    def batch_row_data(self, row: int) -> Dict[str, object]:
        """Handle batch row data in MainWindow."""
        item = self.batch_table.item(row, 0)
        return dict(item.data(Qt.UserRole)) if item is not None else {}

    def set_batch_row(
        self, row: int, data: Dict[str, object], status: str = "Готов"
    ):
        """Set batch row for this workflow."""
        values = (
            data["date"], data["location"], data["model"], data["serial"],
            data["executor"], status,
        )
        for column, value in enumerate(values):
            item = QTableWidgetItem(str(value or "—"))
            if column == 0:
                item.setData(Qt.UserRole, dict(data))
            self.batch_table.setItem(row, column, item)

    def update_batch_count(self):
        """Update batch count for this workflow."""
        count = self.batch_table.rowCount()
        self.batch_count_label.setText(f"В очереди: {count}")
        self.lbl_status.setText(f"В очереди актов: {count}")

    def apply_act_data_to_form(self, data: Dict[str, object]):
        """Apply act data to form for this workflow."""
        self.ed_date.setText(str(data.get("date", "")))
        self.ed_service_from.setText(
            normalize_service_time_input(str(data.get("service_from", "")))
        )
        self.ed_service_to.setText(
            normalize_service_time_input(str(data.get("service_to", "")))
        )
        self.cb_executor.setEditText(str(data.get("executor", "")))
        self.cb_location.setEditText(str(data.get("location", "")))
        self.cb_model.setEditText(str(data.get("model", "")))
        self.ed_serial.setText(str(data.get("serial", "")))
        self.cb_work.setEditText(str(data.get("work", "")))
        self.cb_issues.setEditText(str(data.get("issues", "")))
        self.cb_done_work.setEditText(str(data.get("done_work", "")))
        self.cb_materials.setEditText(str(data.get("materials", "")))
        self.sp_qty.setValue(int(data.get("qty", 0)))
        template_index = self.cb_template.findData(data.get("template"))
        self.cb_template.setCurrentIndex(template_index if template_index >= 0 else 0)

    def edit_batch_act(self, row: int, _column: int):
        """Edit batch act for this workflow."""
        data = self.batch_row_data(row)
        if not data:
            return
        self.apply_act_data_to_form(data)
        self.batch_edit_row = row
        self.btn_batch_add.setText("Сохранить изменения")
        self.lbl_status.setText(
            "Позиция загружена в форму. Измените данные и нажмите "
            "«Сохранить изменения»."
        )

    def cancel_batch_edit(self):
        """Handle cancel batch edit in MainWindow."""
        self.batch_edit_row = None
        self.btn_batch_add.setText("Добавить в очередь")

    def add_current_act_to_batch(self):
        """Add current act to batch for this workflow."""
        data = self.current_act_data()
        template = str(data["template"])
        if not os.path.isfile(template):
            QMessageBox.critical(
                self,
                "Нет шаблона",
                f"Не найден шаблон:\n{template}\n\nПоложи его в папку Data\\Templates.",
            )
            return
        if self.batch_edit_row is not None and self.batch_edit_row < self.batch_table.rowCount():
            self.set_batch_row(self.batch_edit_row, data)
            self.cancel_batch_edit()
        else:
            row = self.batch_table.rowCount()
            self.batch_table.insertRow(row)
            self.set_batch_row(row, data)
        self.update_batch_count()

    def selected_batch_rows(self) -> list[int]:
        """Return batch rows for this workflow."""
        return sorted(
            {index.row() for index in self.batch_table.selectionModel().selectedRows()},
            reverse=True,
        )

    def duplicate_selected_batch_act(self):
        """Handle duplicate selected batch act in MainWindow."""
        rows = self.selected_batch_rows()
        if not rows:
            QMessageBox.information(
                self, "Массовое создание", "Выберите строку для дублирования."
            )
            return
        data = self.batch_row_data(rows[-1])
        row = self.batch_table.rowCount()
        self.batch_table.insertRow(row)
        self.set_batch_row(row, data)
        self.update_batch_count()

    def remove_selected_batch_acts(self):
        """Remove selected batch acts for this workflow."""
        for row in self.selected_batch_rows():
            self.batch_table.removeRow(row)
        self.cancel_batch_edit()
        self.update_batch_count()

    def clear_batch_acts(self):
        """Clear batch acts for this workflow."""
        if self.batch_table.rowCount() == 0:
            return
        answer = QMessageBox.question(
            self,
            "Очистить очередь",
            "Удалить все позиции из очереди массового генератора?\n\n"
            "Уже созданные файлы DOCX не затрагиваются.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        self.batch_table.setRowCount(0)
        self.cancel_batch_edit()
        self.update_batch_count()

    def generate_act_document(
        self, data: Dict[str, object], *, preview: bool,
        surname_only: bool = False,
    ) -> Dict[str, str] | None:
        """Fill one DOCX template and return its final output path."""
        date = str(data["date"])
        executor = str(data["executor"])
        location = str(data["location"])
        model = str(data["model"])
        serial = str(data["serial"])
        work = str(data["work"])
        issues = str(data["issues"])
        done_work = str(data["done_work"])
        materials = str(data["materials"])
        qty_int = int(data["qty"])
        template = str(data["template"])

        if not os.path.isfile(template):
            raise FileNotFoundError(f"Не найден шаблон: {template}")

        service_from_doc = SERVICE_PERIOD_BLANK
        service_to_doc = SERVICE_PERIOD_BLANK
        if is_validator_template(template) and not surname_only:
            service_from_doc, service_to_doc = validate_service_period(
                str(data.get("service_from", "")),
                str(data.get("service_to", "")),
                date,
            )
        elif surname_only:
            service_from_doc = ""
            service_to_doc = ""

        mapping = {}
        if surname_only:
            # Remove the complete equipment value, including the literal №
            # between its two placeholders. Replacing MODEL and SERIAL alone
            # would leave an orphan number sign in the otherwise blank row.
            mapping["{MODEL} № {SERIAL}"] = ""
        mapping.update({
            # COLLEAGUE EDIT POINT: DOCX placeholder-to-form mappings.
            "{DATE}": date,
            "{SERVICE_FROM}": service_from_doc,
            "{SERVICE_TO}": service_to_doc,
            "{EXECUTOR}": executor,
            "{LOCATION}": location,
            "{MODEL}": model,
            "{SERIAL}": serial,
            "{WORK}": work,
            "{ISSUES}": issues,
            "{DONE_WORK}": done_work,
            "{MATERIALS}": materials,
            "{MATERIALS_QTY}": str(qty_int) if qty_int > 0 else "",
            "{UNIT}": "шт." if qty_int > 0 else "",
        })

        doc = Document(template)
        replace_placeholders_docx(doc, mapping)
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S-%f")
        if surname_only:
            template_label = os.path.splitext(os.path.basename(template))[0]
            base = safe_filename(f"Акт_{executor}_{template_label}_{timestamp}")
        else:
            base = safe_filename(
                f"Акт_{location or 'БезСтанции'}_{serial or 'БезККТ'}_{timestamp}"
            )
        out_path = os.path.join(OUTPUT_DIR, base + ".docx")
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        if preview:
            with tempfile.TemporaryDirectory(
                prefix="actgenerator-preview-", ignore_cleanup_errors=True
            ) as preview_dir:
                preview_path = os.path.join(preview_dir, base + ".docx")
                doc.save(preview_path)
                dialog = ActPreviewDialog(preview_path, self)
                if dialog.exec() != QDialog.Accepted:
                    return None
                shutil.copy2(preview_path, out_path)
        else:
            doc.save(out_path)

        history_error = ""
        try:
            save_history_record(HISTORY_DIR, APP_DIR, out_path, data)
        except OSError as error:
            history_error = str(error)

        excel_error = ""
        if surname_only:
            excel_status = "Excel не изменялся: создан чистый акт по фамилии."
        else:
            try:
                excel_result = append_act_to_workbook(
                self.active_excel_workbook(),
                {
                    "date": date,
                    "executor": executor,
                    "model": model,
                    "serial": serial,
                    "location": location,
                    "work": work,
                    "issues": issues,
                    "materials": materials,
                    "done_work": done_work,
                    "qty": qty_int if qty_int > 0 else "",
                    "area": area_for_station(location),
                    "notes": note_for_materials(materials),
                },
                )
                excel_status = (
                    f"Excel: лист «{excel_result.sheet_name}», "
                    f"строка {excel_result.row_number}."
                )
            except ExcelExportError as error:
                excel_error = str(error)
                excel_status = f"Excel не обновлён: {error}"

        return {
            "out_path": out_path,
            "excel_status": excel_status,
            "excel_error": excel_error,
            "history_error": history_error,
        }

    def create_surname_only_act(self, template: str):
        """Create a blank template with only the executor name and initials."""
        executor_name = executor_name_with_initials(self.cb_executor.text())
        if not executor_name:
            QMessageBox.warning(
                self,
                "Не указана фамилия",
                "Выберите исполнителя или введите фамилию буквами.",
            )
            return

        data = {
            "date": "",
            "service_from": "",
            "service_to": "",
            "executor": executor_name,
            "location": "",
            "model": "",
            "serial": "",
            "work": "",
            "issues": "",
            "done_work": "",
            "materials": "",
            "qty": 0,
            "template": template,
        }
        try:
            result = self.generate_act_document(
                data,
                preview=self.app_settings.get("preview_before_save", False),
                surname_only=True,
            )
            if result is None:
                self.lbl_status.setText(
                    "Создание чистого акта отменено."
                )
                return
            if self.act_history_panel.isChecked():
                self.refresh_history()
            self.lbl_status.setText(
                f"✅ Чистый акт создан: {result['out_path']}\n"
                f"{result['excel_status']}"
            )
            msg = QMessageBox(self)
            msg.setWindowTitle("Готово")
            msg.setText(f"Создан акт по фамилии: {executor_name}.")
            msg.setInformativeText(result["out_path"])
            btn_open = msg.addButton("Открыть", QMessageBox.AcceptRole)
            msg.addButton("OK", QMessageBox.RejectRole)
            msg.exec()
            if msg.clickedButton() == btn_open:
                open_file(result["out_path"])
        except Exception as error:
            QMessageBox.critical(self, "Ошибка", str(error))

    def create_batch_acts(self):
        """Generate every queued act without opening individual previews."""
        if self.batch_table.rowCount() == 0:
            QMessageBox.information(
                self, "Массовое создание", "Очередь актов пуста."
            )
            return

        self.cancel_batch_edit()
        total = self.batch_table.rowCount()
        self.batch_progress.setRange(0, total)
        self.batch_progress.setValue(0)
        self.batch_progress.setFormat("Обработано: %v из %m")
        self.batch_progress.show()
        self.btn_batch_create.setEnabled(False)
        successful_rows = []
        failures = []
        excel_warnings = []
        history_warnings = []
        for row in range(total):
            data = self.batch_row_data(row)
            caption = self.batch_act_caption(data)
            status_item = self.batch_table.item(row, 5)
            status_item.setText("Создаётся…")
            QApplication.processEvents()
            try:
                result = self.generate_act_document(data, preview=False)
            except Exception as error:
                status_item.setText("Ошибка")
                status_item.setForeground(QColor("#e05252"))
                status_item.setToolTip(str(error))
                failures.append(f"{caption}: {error}")
                self.batch_progress.setValue(row + 1)
                continue
            if result is None:
                status_item.setText("Ошибка")
                failures.append(f"{caption}: создание отменено")
                self.batch_progress.setValue(row + 1)
                continue
            status_item.setText("Создан")
            status_item.setForeground(QColor("#39a96b"))
            successful_rows.append(row)
            if result["excel_error"]:
                excel_warnings.append(f"{caption}: {result['excel_error']}")
            if result["history_error"]:
                history_warnings.append(f"{caption}: {result['history_error']}")
            self.batch_progress.setValue(row + 1)
            QApplication.processEvents()

        for row in reversed(successful_rows):
            self.batch_table.removeRow(row)

        created = len(successful_rows)
        self.btn_batch_create.setEnabled(True)
        self.batch_progress.hide()
        self.batch_count_label.setText(f"В очереди: {self.batch_table.rowCount()}")
        self.lbl_status.setText(
            f"✅ Массовое создание завершено: создано {created}, "
            f"ошибок {len(failures)}. В очереди: {self.batch_table.rowCount()}."
        )
        details = []
        if failures:
            details.append("Ошибки:\n" + "\n".join(failures[:10]))
        if excel_warnings:
            details.append("Excel:\n" + "\n".join(excel_warnings[:10]))
        if history_warnings:
            details.append("История:\n" + "\n".join(history_warnings[:10]))
        if self.act_history_panel.isChecked():
            self.refresh_history()
        QMessageBox.information(
            self,
            "Массовое создание завершено",
            f"Создано актов: {created}.\nОшибок: {len(failures)}."
            + ("\n\n" + "\n\n".join(details) if details else ""),
        )

    def create_act(self):
        """Validate the form, optionally preview the act, and save the confirmed document."""
        try:
            result = self.generate_act_document(
                self.current_act_data(),
                preview=self.app_settings.get("preview_before_save", False),
            )
            if result is None:
                self.lbl_status.setText(
                    "Сохранение отменено. Можно изменить данные и снова открыть предпросмотр."
                )
                return

            out_path = result["out_path"]
            excel_status = result["excel_status"]
            if result["excel_error"]:
                QMessageBox.warning(
                    self,
                    "Акт сохранён, Excel не обновлён",
                    f"DOCX создан:\n{out_path}\n\n{result['excel_error']}",
                )

            if result["history_error"]:
                QMessageBox.warning(
                    self,
                    "Акт сохранён, история не обновлена",
                    f"DOCX создан:\n{out_path}\n\n{result['history_error']}",
                )
            if self.act_history_panel.isChecked():
                self.refresh_history()

            self.lbl_status.setText(f"✅ Файл создан: {out_path}\n{excel_status}")
            msg = QMessageBox(self)
            msg.setWindowTitle("Готово")
            msg.setText("Акт создан.")
            msg.setInformativeText(f"{out_path}\n\n{excel_status}")
            btn_open = msg.addButton("Открыть", QMessageBox.AcceptRole)
            msg.addButton("OK", QMessageBox.RejectRole)
            msg.exec()
            if msg.clickedButton() == btn_open:
                open_file(out_path)
        except Exception as error:
            QMessageBox.critical(
                self, "Ошибка", f"Не удалось создать акт:\n{error}"
            )


def main():
    """Initialize the application, update service and main window event loop."""
    app = QApplication(sys.argv)
    app.setApplicationName("Генератор актов")
    app.setOrganizationName("ActGeneratorPC")
    app.setWindowIcon(icon_or_empty(ICON_APP))
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    QTimer.singleShot(2500, confirm_healthy_startup)

    # COLLEAGUE EDIT POINT: startup delay for the automatic update check.
    # Keep the updater alive for the lifetime of the main window.
    window.update_manager = UpdateManager(window)
    QTimer.singleShot(1500, window.update_manager.check_async)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
