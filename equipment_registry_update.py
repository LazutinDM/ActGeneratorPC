"""Download and validate the factory-number registry used by ActGeneratorPC.

The source is the user-approved Google Sheets workbook.  The application reads
its XLSX export with the Python standard library, so the portable build does not
depend on Microsoft Excel, Google credentials, or an additional XLSX package.
The installed JSON is replaced only after the whole workbook has been parsed and
validated successfully.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
import urllib.request
import zipfile
from decimal import Decimal, InvalidOperation
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree


# COLLEAGUE EDIT POINT: replace this ID only if the approved source workbook
# changes.  The three worksheet names below are the public tabs previously
# provided by the user (МКТФ, Валидатор ВП-FT and БПА).
GOOGLE_WORKBOOK_ID = "1Y5YLY2sqokyPYJw-_Sr2Cl9UK99MbW4NslK1b5LBWS8"
GOOGLE_WORKBOOK_EXPORT_URL = (
    f"https://docs.google.com/spreadsheets/d/{GOOGLE_WORKBOOK_ID}/export?format=xlsx"
)
SOURCE_SHEETS = ("МКТФ", "Валидатор ВП-FT", "БПА")
MAX_DOWNLOAD_BYTES = 30 * 1024 * 1024

# Source-table spelling -> exact value used by the application.
MODEL_ALIASES = {
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

# COLLEAGUE EDIT POINT: station spellings in the source may differ from the
# station list shown in the application.
STATION_ALIASES = {
    "зеленоград-крюково": "КРЮКОВО",
    "зеленоград крюково": "КРЮКОВО",
}


class RegistryUpdateError(RuntimeError):
    """Raised when downloaded data cannot safely replace the current registry."""


def clean(value) -> str:
    """Normalize a worksheet value without changing meaningful characters."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()


def clean_identifier(value) -> str:
    """Keep Excel identifiers as plain integers instead of ``123.0``/E notation."""
    text = clean(value)
    if not text:
        return ""
    try:
        number = Decimal(text)
    except InvalidOperation:
        return text
    if number == number.to_integral_value():
        return format(number.quantize(Decimal(1)), "f")
    return text


def canonical_model(value: str) -> str:
    """Return the model spelling used by the program's equipment list."""
    normalized = clean(value).replace("–", "-").replace("—", "-")
    return MODEL_ALIASES.get(normalized.upper(), normalized)


def canonical_station(value: str) -> str:
    """Return the approved station spelling used by the program."""
    normalized = clean(value).replace("–", "-").replace("—", "-")
    return STATION_ALIASES.get(normalized.casefold(), normalized)


def allowed_stations_through_alabushevo(stations: list[str]) -> list[str]:
    """Normalize a station list and discard every entry after АЛАБУШЕВО."""
    result: list[str] = []
    seen: set[str] = set()
    found_limit = False
    for station in stations:
        value = canonical_station(station)
        key = value.casefold()
        if not value or key in seen or found_limit:
            continue
        result.append(value)
        seen.add(key)
        if key == "алабушево":
            found_limit = True
    if not found_limit:
        raise RegistryUpdateError("В списке станций не найдена станция АЛАБУШЕВО.")
    return result


def _column_number(reference: str) -> int:
    """Handle column number."""
    letters = re.match(r"[A-Za-z]+", reference or "")
    if not letters:
        return 0
    result = 0
    for character in letters.group(0).upper():
        result = result * 26 + ord(character) - ord("A") + 1
    return result


def _shared_strings(archive: zipfile.ZipFile) -> list[str]:
    """Handle shared strings."""
    try:
        payload = archive.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ElementTree.fromstring(payload)
    values = []
    for item in root:
        values.append("".join(node.text or "" for node in item.iter() if node.tag.endswith("}t")))
    return values


def _sheet_paths(archive: zipfile.ZipFile) -> dict[str, str]:
    """Handle sheet paths."""
    workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    relationships = ElementTree.fromstring(
        archive.read("xl/_rels/workbook.xml.rels")
    )
    target_by_id = {
        relation.attrib["Id"]: relation.attrib["Target"]
        for relation in relationships
    }
    result = {}
    for sheet in workbook.iter():
        if not sheet.tag.endswith("}sheet"):
            continue
        relation_id = next(
            (value for key, value in sheet.attrib.items() if key.endswith("}id")),
            "",
        )
        target = target_by_id.get(relation_id, "")
        if not target:
            continue
        if target.startswith("/"):
            path = target.lstrip("/")
        else:
            path = str(PurePosixPath("xl") / target)
        result[sheet.attrib.get("name", "")] = str(PurePosixPath(path))
    return result


def _worksheet_rows(
    archive: zipfile.ZipFile, sheet_path: str, shared_strings: list[str]
):
    """Handle worksheet rows."""
    with archive.open(sheet_path) as stream:
        for _event, element in ElementTree.iterparse(stream, events=("end",)):
            if not element.tag.endswith("}row"):
                continue
            values: dict[int, str] = {}
            for cell in element:
                if not cell.tag.endswith("}c"):
                    continue
                column = _column_number(cell.attrib.get("r", ""))
                cell_type = cell.attrib.get("t", "")
                value_node = next(
                    (node for node in cell if node.tag.endswith("}v")), None
                )
                if cell_type == "inlineStr":
                    raw = "".join(
                        node.text or "" for node in cell.iter() if node.tag.endswith("}t")
                    )
                else:
                    raw = value_node.text if value_node is not None else ""
                    if cell_type == "s" and raw:
                        try:
                            raw = shared_strings[int(raw)]
                        except (ValueError, IndexError):
                            raw = ""
                values[column] = clean(raw)
            yield values
            element.clear()


def _add_record(records: dict[str, dict[str, str]], record: dict[str, str]):
    """Add record for this workflow."""
    serial = clean_identifier(record.get("serial", ""))
    if not serial:
        return
    record["serial"] = serial
    existing = records.get(serial.casefold())
    if existing and (
        existing["station"].casefold(), existing["model"].casefold()
    ) != (record["station"].casefold(), record["model"].casefold()):
        raise RegistryUpdateError(
            f"Заводской номер {serial} относится к разным станциям или моделям."
        )
    records[serial.casefold()] = record


def extract_registry(
    workbook_path: str | os.PathLike[str], allowed_stations: list[str]
) -> list[dict[str, str]]:
    """Extract and validate records from the three approved worksheet tabs."""
    allowed = allowed_stations_through_alabushevo(allowed_stations)
    allowed_keys = {station.casefold() for station in allowed}
    records: dict[str, dict[str, str]] = {}
    try:
        archive = zipfile.ZipFile(workbook_path)
    except (OSError, zipfile.BadZipFile) as exc:
        raise RegistryUpdateError("Загруженный файл не является исправной XLSX-книгой.") from exc

    with archive:
        try:
            sheet_paths = _sheet_paths(archive)
            shared_strings = _shared_strings(archive)
        except (KeyError, ElementTree.ParseError) as exc:
            raise RegistryUpdateError("В XLSX-книге повреждена структура листов.") from exc
        missing = [name for name in SOURCE_SHEETS if name not in sheet_paths]
        if missing:
            raise RegistryUpdateError(
                "В книге отсутствуют обязательные листы: " + ", ".join(missing)
            )

        def add(station, model, serial, equipment_id="", station_number="", source=""):
            """Add the related value for this workflow."""
            station = canonical_station(station)
            model = canonical_model(model)
            if not station or station.casefold() not in allowed_keys or not model or not clean_identifier(serial):
                return
            _add_record(records, {
                "station": station,
                "model": model,
                "serial": clean_identifier(serial),
                "equipmentId": clean_identifier(equipment_id),
                "stationNumber": clean_identifier(station_number),
                "sourceSheet": source,
            })

        for row_number, row in enumerate(
            _worksheet_rows(archive, sheet_paths["МКТФ"], shared_strings), start=1
        ):
            if row_number >= 2:
                add(row.get(3), "МКТФ", row.get(6), row.get(7), source="МКТФ")

        for row_number, row in enumerate(
            _worksheet_rows(archive, sheet_paths["Валидатор ВП-FT"], shared_strings), start=1
        ):
            if row_number >= 2:
                add(
                    row.get(3), "Валидатор ВП-FT", row.get(4), row.get(5),
                    source="Валидатор ВП-FT",
                )

        # The current БПА data block is in columns AJ:AM, while the station is C.
        for row_number, row in enumerate(
            _worksheet_rows(archive, sheet_paths["БПА"], shared_strings), start=1
        ):
            if row_number >= 3:
                add(
                    row.get(3), row.get(36), row.get(37), row.get(38), row.get(39),
                    source="БПА",
                )

    station_rank = {station.casefold(): index for index, station in enumerate(allowed)}
    result = sorted(
        records.values(),
        key=lambda item: (
            station_rank[item["station"].casefold()],
            item["model"].casefold(),
            len(item["serial"]),
            item["serial"],
        ),
    )
    if not result:
        raise RegistryUpdateError("После проверки таблиц не найдено ни одной записи.")
    return result


def download_workbook(destination: str | os.PathLike[str]) -> None:
    """Download the public XLSX export with a strict size limit."""
    request = urllib.request.Request(
        GOOGLE_WORKBOOK_EXPORT_URL,
        headers={"User-Agent": "ActGeneratorPC/1.0"},
    )
    try:
        with urllib.request.urlopen(request, timeout=30) as response, open(destination, "wb") as output:
            total = 0
            while chunk := response.read(1024 * 256):
                total += len(chunk)
                if total > MAX_DOWNLOAD_BYTES:
                    raise RegistryUpdateError("Размер загружаемой таблицы превышает 30 МБ.")
                output.write(chunk)
    except RegistryUpdateError:
        raise
    except OSError as exc:
        raise RegistryUpdateError(f"Не удалось скачать таблицу: {exc}") from exc


def update_registry(
    registry_path: str | os.PathLike[str], allowed_stations: list[str]
) -> dict[str, object]:
    """Download, validate, back up and atomically install the registry JSON."""
    registry_path = Path(registry_path)
    registry_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="actgenerator-registry-") as directory:
        workbook_path = Path(directory) / "registry.xlsx"
        download_workbook(workbook_path)
        records = extract_registry(workbook_path, allowed_stations)
        staged_path = Path(directory) / "station_serials.json"
        staged_path.write_text(
            json.dumps(records, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        # Re-read staged JSON before touching the installed registry.
        json.loads(staged_path.read_text(encoding="utf-8"))

        backup_path = registry_path.with_name("station_serials.previous.json")
        if registry_path.exists():
            shutil.copy2(registry_path, backup_path)
        replacement = registry_path.with_suffix(".json.new")
        shutil.copy2(staged_path, replacement)
        os.replace(replacement, registry_path)

    return {
        "count": len(records),
        "stations": len({record["station"] for record in records}),
        "backup": str(backup_path) if backup_path.exists() else "",
    }
