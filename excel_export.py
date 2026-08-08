"""Append act data to an existing XLSX table while preserving its structure."""

from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
import datetime
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Mapping


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


class ExcelExportError(RuntimeError):
    """Raised when an act cannot be appended safely."""


@dataclass(frozen=True)
class ExcelAppendResult:
    workbook_path: str
    sheet_name: str
    row_number: int
    backup_path: str


HEADER_TO_FIELD = {
    "участок": "area",
    "число": "date",
    "представитель": "executor",
    "типккт": "model",
    "номерккт": "serial",
    "станция": "location",
    "выполненныеработы": "work",
    "выполненыеработы": "work",
    "описание": "issues",
    "запчасти": "materials",
    "подробноеописание": "done_work",
    "колво": "qty",
    "примечания": "notes",
}
IGNORED_HEADERS = {"проверил", "наличиеакта", "заявка"}
REQUIRED_FIELDS = set(HEADER_TO_FIELD.values())
MONTH_NAMES = (
    "январь", "февраль", "март", "апрель", "май", "июнь",
    "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь",
)

# COLLEAGUE EDIT POINT: station-to-area rules used by the Excel export.
# Values intentionally match the wording already used in the supplied table.
KRYUKOVO_STATIONS = {
    "левобережная",
    "химки",
    "молжаниново",
    "новоподрезково",
    "подрезково",
    "сходня",
    "фирсановская",
    "малино",
    "зеленоградкрюково",
    "алабушево",
}
MOSCOW_STATIONS = {
    "москва",
    "останкино",
    "петровскоразумовская",
    "лихоборы",
    "моссельмаш",
    "грачевская",
    "грачевскаябывшховрино",
    "ховрино",
}
MATERIAL_NOTE_TRIGGER = re.compile(r"(?<![0-9A-ZА-ЯЁ])МТТПК(?![0-9A-ZА-ЯЁ])", re.IGNORECASE)
MATERIAL_NOTE_TEXT = "выданно МТППК"


def load_excel_path(config_path: str) -> str:
    try:
        with open(config_path, "r", encoding="utf-8") as handle:
            value = json.load(handle).get("workbook_path", "")
    except (OSError, ValueError, AttributeError):
        return ""
    return str(value).strip()


def save_excel_path(config_path: str, workbook_path: str) -> None:
    path = Path(config_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    configured_path = str(workbook_path).strip()
    if configured_path:
        configured_path = str(Path(configured_path).resolve())
    with open(temporary, "w", encoding="utf-8") as handle:
        json.dump({"workbook_path": configured_path}, handle, ensure_ascii=False, indent=2)
    os.replace(temporary, path)


def ensure_embedded_excel_workbook(template_path: str, workbook_path: str) -> str:
    """Create the user's editable workbook from the bundled template once."""
    template = Path(template_path).resolve()
    workbook = Path(workbook_path).resolve()
    if workbook.is_file():
        return str(workbook)
    if template.suffix.casefold() != ".xlsx" or not template.is_file():
        raise ExcelExportError(f"Встроенный шаблон Excel не найден: {template}")

    workbook.parent.mkdir(parents=True, exist_ok=True)
    temporary_handle = tempfile.NamedTemporaryFile(
        prefix=f".{workbook.stem}-initial-",
        suffix=".xlsx",
        dir=workbook.parent,
        delete=False,
    )
    temporary_path = Path(temporary_handle.name)
    temporary_handle.close()
    try:
        shutil.copy2(template, temporary_path)
        with zipfile.ZipFile(temporary_path, "r") as archive:
            broken = archive.testzip()
            if broken:
                raise ExcelExportError(
                    f"Встроенный шаблон Excel повреждён: {broken}"
                )
            archive.getinfo("xl/workbook.xml")
        # os.replace keeps creation atomic. Existing user data is never copied
        # over because the method returns above when the workbook already exists.
        if not workbook.exists():
            os.replace(temporary_path, workbook)
    except (OSError, zipfile.BadZipFile, KeyError) as exc:
        raise ExcelExportError(
            f"Не удалось создать встроенную Excel-таблицу: {exc}"
        ) from exc
    finally:
        temporary_path.unlink(missing_ok=True)
    return str(workbook)


def resolve_excel_workbook(
    config_path: str,
    template_path: str,
    embedded_workbook_path: str,
) -> str:
    """Return an external override or the persistent built-in workbook."""
    configured = load_excel_path(config_path)
    if configured and Path(configured).is_file():
        return str(Path(configured).resolve())
    return ensure_embedded_excel_workbook(template_path, embedded_workbook_path)


def _normalize_header(value: str) -> str:
    normalized = (value or "").strip().casefold().replace("ё", "е")
    normalized = normalized.replace("cтанция", "станция")
    normalized = normalized.replace("№", "номер ")
    return re.sub(r"[^0-9a-zа-я]+", "", normalized)


def _normalize_station(value: str) -> str:
    normalized = (value or "").strip().casefold().replace("ё", "е")
    normalized = re.sub(r"^\s*(?:пл|платформа)\.?\s*", "", normalized)
    return re.sub(r"[^0-9a-zа-я]+", "", normalized)


def area_for_station(station: str) -> str:
    """Return the workbook area name for a known station, else an empty value."""
    normalized = _normalize_station(station)
    if normalized in KRYUKOVO_STATIONS:
        return "Крюковский"
    if normalized in MOSCOW_STATIONS:
        return "Московский"
    return ""


def note_for_materials(materials: str) -> str:
    """Return the requested note when materials contain the standalone word МТТПК."""
    return MATERIAL_NOTE_TEXT if MATERIAL_NOTE_TRIGGER.search(materials or "") else ""


def _column_number(reference: str) -> int:
    letters = re.match(r"[A-Za-z]+", reference)
    if not letters:
        raise ExcelExportError(f"Некорректная ссылка на ячейку: {reference}")
    number = 0
    for character in letters.group(0).upper():
        number = number * 26 + ord(character) - 64
    return number


def _column_name(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _resolve_part(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return str(PurePosixPath(source_part).parent.joinpath(target))


def _shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    values = []
    for item in root.findall(f"{{{MAIN_NS}}}si"):
        values.append("".join(node.text or "" for node in item.iter(f"{{{MAIN_NS}}}t")))
    return values


def _cell_value(cell: ET.Element, shared: list[str]) -> str:
    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(f"{{{MAIN_NS}}}t"))
    value = cell.find(f"{{{MAIN_NS}}}v")
    if value is None or value.text is None:
        return ""
    if cell_type == "s":
        try:
            return shared[int(value.text)]
        except (ValueError, IndexError):
            return ""
    return value.text


def _workbook_sheets(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        item.get("Id"): _resolve_part("xl/workbook.xml", item.get("Target", ""))
        for item in relationships.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    result = []
    for sheet in workbook.findall(f".//{{{MAIN_NS}}}sheet"):
        relationship_id = sheet.get(f"{{{DOC_REL_NS}}}id")
        target = targets.get(relationship_id)
        if target:
            result.append((sheet.get("name", "Лист"), target))
    return result


def _preferred_sheets(sheets: list[tuple[str, str]], date_value: object) -> list[tuple[str, str]]:
    try:
        act_date = datetime.datetime.strptime(str(date_value).strip(), "%d.%m.%Y").date()
    except ValueError:
        return sheets
    month = MONTH_NAMES[act_date.month - 1]
    normalized_candidates = {
        _normalize_header(f"{month} {act_date:%y}"),
        _normalize_header(f"{month} {act_date:%Y}"),
    }
    matching = [item for item in sheets if _normalize_header(item[0]) in normalized_candidates]
    return matching + [item for item in sheets if item not in matching]


def _header_map(root: ET.Element, shared: list[str]) -> dict[str, int]:
    first_row = root.find(f".//{{{MAIN_NS}}}sheetData/{{{MAIN_NS}}}row[@r='1']")
    if first_row is None:
        return {}
    result = {}
    for cell in first_row.findall(f"{{{MAIN_NS}}}c"):
        header = _normalize_header(_cell_value(cell, shared))
        if header:
            result[header] = _column_number(cell.get("r", ""))
    return result


def _row_cells(row: ET.Element) -> dict[int, ET.Element]:
    return {
        _column_number(cell.get("r", "")): cell
        for cell in row.findall(f"{{{MAIN_NS}}}c")
        if cell.get("r")
    }


def _find_empty_row(root: ET.Element, columns: set[int], shared: list[str]) -> ET.Element:
    sheet_data = root.find(f".//{{{MAIN_NS}}}sheetData")
    if sheet_data is None:
        raise ExcelExportError("В листе отсутствует область данных.")
    for row in sheet_data.findall(f"{{{MAIN_NS}}}row"):
        try:
            row_number = int(row.get("r", "0"))
        except ValueError:
            continue
        if row_number < 2:
            continue
        cells = _row_cells(row)
        if all(not _cell_value(cells[column], shared) if column in cells else True for column in columns):
            return row
    raise ExcelExportError("В таблице нет свободной строки для нового акта.")


def _set_cell(row: ET.Element, column: int, value: object) -> None:
    row_number = int(row.get("r", "0"))
    reference = f"{_column_name(column)}{row_number}"
    cells = _row_cells(row)
    cell = cells.get(column)
    if cell is None:
        cell = ET.Element(f"{{{MAIN_NS}}}c", {"r": reference})
        insert_at = len(row)
        for index, sibling in enumerate(row.findall(f"{{{MAIN_NS}}}c")):
            if _column_number(sibling.get("r", "")) > column:
                insert_at = list(row).index(sibling)
                break
        row.insert(insert_at, cell)
    for child in list(cell):
        cell.remove(child)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        cell.attrib.pop("t", None)
        node = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
        node.text = str(value)
    else:
        cell.set("t", "inlineStr")
        inline = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
        text = ET.SubElement(inline, f"{{{MAIN_NS}}}t")
        text.text = str(value)


def _register_namespaces(xml_bytes: bytes) -> None:
    for _, (prefix, uri) in ET.iterparse(__import__("io").BytesIO(xml_bytes), events=("start-ns",)):
        if prefix != "xml":
            try:
                ET.register_namespace(prefix or "", uri)
            except ValueError:
                pass


def _build_updated_workbook(source: Path, destination: Path, values: Mapping[str, object]) -> tuple[str, int]:
    with zipfile.ZipFile(source, "r") as archive:
        shared = _shared_strings(archive)
        selected = None
        sheets = _preferred_sheets(_workbook_sheets(archive), values.get("date", ""))
        for sheet_name, sheet_part in sheets:
            xml_bytes = archive.read(sheet_part)
            root = ET.fromstring(xml_bytes)
            headers = _header_map(root, shared)
            field_columns = {
                HEADER_TO_FIELD[header]: column
                for header, column in headers.items()
                if header in HEADER_TO_FIELD
            }
            if REQUIRED_FIELDS.issubset(field_columns):
                selected = (sheet_name, sheet_part, xml_bytes, root, field_columns)
                break
        if selected is None:
            raise ExcelExportError(
                "Не найден лист с ожидаемыми заголовками: Число, Представитель, "
                "Тип ККТ, № ККТ, Станция и данные выполненных работ."
            )

        sheet_name, sheet_part, xml_bytes, root, field_columns = selected
        target_row = _find_empty_row(root, set(field_columns.values()), shared)
        for field, column in field_columns.items():
            value = values.get(field, "")
            if value not in (None, "", 0):
                _set_cell(target_row, column, value)

        _register_namespaces(xml_bytes)
        replacement = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(destination, "w") as output:
            for item in archive.infolist():
                output.writestr(item, replacement if item.filename == sheet_part else archive.read(item.filename))

    with zipfile.ZipFile(destination, "r") as check:
        broken = check.testzip()
        if broken:
            raise ExcelExportError(f"Проверка XLSX выявила повреждённую часть: {broken}")
        ET.fromstring(check.read(sheet_part))
    return sheet_name, int(target_row.get("r", "0"))


def append_act_to_workbook(workbook_path: str, values: Mapping[str, object]) -> ExcelAppendResult:
    """Append an act to the first empty row, leaving ignored columns untouched."""
    source = Path(workbook_path).resolve()
    if source.suffix.casefold() != ".xlsx" or not source.is_file():
        raise ExcelExportError(f"Файл Excel не найден или имеет неверный формат: {source}")

    temporary_handle = tempfile.NamedTemporaryFile(
        prefix=f".{source.stem}-actgenerator-",
        suffix=".xlsx",
        dir=source.parent,
        delete=False,
    )
    temporary_path = Path(temporary_handle.name)
    temporary_handle.close()
    backup_path = source.with_suffix(source.suffix + ".before-actgenerator.bak")
    try:
        sheet_name, row_number = _build_updated_workbook(source, temporary_path, values)
        shutil.copy2(source, backup_path)
        os.replace(temporary_path, source)
    except PermissionError as exc:
        raise ExcelExportError(
            "Excel-таблица занята или доступна только для чтения. Закройте её в Excel и повторите создание акта."
        ) from exc
    except (OSError, zipfile.BadZipFile, ET.ParseError) as exc:
        raise ExcelExportError(f"Не удалось безопасно обновить Excel-таблицу: {exc}") from exc
    finally:
        temporary_path.unlink(missing_ok=True)

    return ExcelAppendResult(str(source), sheet_name, row_number, str(backup_path))
