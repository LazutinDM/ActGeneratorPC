import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from excel_export import (
    MAIN_NS,
    append_act_to_workbook,
    area_for_station,
    ensure_embedded_excel_workbook,
    load_excel_path,
    note_for_materials,
    resolve_excel_workbook,
    save_excel_path,
)


HEADERS = [
    "Проверил", "Наличие акта", "Заявка", "Участок", "Число",
    "Представитель", "Тип ККТ", "№ ККТ", "Cтанция",
    "Выполненые работы", "Описание", "Запчасти",
    "Подробное описание", "Кол-во", "Примечания",
]


def cell(column, row, value):
    reference = f"{column}{row}"
    if isinstance(value, bool):
        return f'<c r="{reference}" t="b"><v>{int(value)}</v></c>'
    return f'<c r="{reference}" t="inlineStr"><is><t>{value}</t></is></c>'


def make_workbook(path: Path):
    columns = [chr(ord("A") + index) for index in range(len(HEADERS))]
    header_xml = "".join(cell(column, 1, value) for column, value in zip(columns, HEADERS))
    row2_xml = cell("A", 2, False) + cell("B", 2, False) + cell("C", 2, "REQ-KEEP")
    worksheet = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="{MAIN_NS}"><sheetData>
<row r="1">{header_xml}</row><row r="2">{row2_xml}</row><row r="3"/>
</sheetData></worksheet>'''
    workbook = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{MAIN_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Август 26" sheetId="1" r:id="rId1"/></sheets></workbook>'''
    relationships = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>'''
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", relationships)
        archive.writestr("xl/worksheets/sheet1.xml", worksheet)


def values_by_reference(path: Path):
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    result = {}
    for item in root.findall(f".//{{{MAIN_NS}}}c"):
        text = item.find(f"{{{MAIN_NS}}}v")
        if text is None:
            text = item.find(f"{{{MAIN_NS}}}is/{{{MAIN_NS}}}t")
        result[item.get("r")] = "" if text is None else text.text
    return result


class ExcelExportTests(unittest.TestCase):
    def test_appends_values_and_does_not_touch_ignored_columns(self):
        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory) / "Tables.xlsx"
            make_workbook(workbook)

            result = append_act_to_workbook(
                str(workbook),
                {
                    "date": "08.08.2026",
                    "executor": "Петров",
                    "model": "МКТФ",
                    "serial": "52881",
                    "location": "ХИМКИ",
                    "work": "Замена блока",
                    "issues": "Сбой",
                    "materials": "Блок питания",
                    "done_work": "Блок заменён",
                    "qty": 1,
                    "area": "Крюковский",
                    "notes": "выданно МТППК",
                },
            )

            self.assertEqual(2, result.row_number)
            values = values_by_reference(workbook)
            self.assertEqual("0", values["A2"])
            self.assertEqual("0", values["B2"])
            self.assertEqual("REQ-KEEP", values["C2"])
            self.assertEqual("Крюковский", values["D2"])
            self.assertEqual("08.08.2026", values["E2"])
            self.assertEqual("Петров", values["F2"])
            self.assertEqual("1", values["N2"])
            self.assertEqual("выданно МТППК", values["O2"])
            self.assertTrue(Path(result.backup_path).is_file())

    def test_station_area_mapping(self):
        kryukovo = (
            "пл. Левобережная", "Химки", "пл. Молжаниново",
            "пл. Новоподрезково", "пл. Подрезково", "Сходня",
            "пл. Фирсановская", "пл. Малино", "Зеленоград-Крюково",
            "пл. Алабушево",
        )
        moscow = (
            "Москва", "пл. Останкино", "пл. Петровско-Разумовская",
            "пл. Лихоборы", "пл. Моссельмаш",
            "Грачёвская (бывш. Ховрино)", "пл. Ховрино",
        )
        for station in kryukovo:
            with self.subTest(station=station):
                self.assertEqual("Крюковский", area_for_station(station))
        for station in moscow:
            with self.subTest(station=station):
                self.assertEqual("Московский", area_for_station(station))
        self.assertEqual("", area_for_station("Неизвестная станция"))

    def test_material_note_requires_the_standalone_trigger_word(self):
        self.assertEqual("выданно МТППК", note_for_materials("Комплект МТТПК"))
        self.assertEqual("выданно МТППК", note_for_materials("мттпк"))
        self.assertEqual("", note_for_materials("XМТТПК"))
        self.assertEqual("", note_for_materials("МТППК"))

    def test_settings_round_trip(self):
        with tempfile.TemporaryDirectory() as directory:
            config = Path(directory) / "excel_export.json"
            workbook = Path(directory) / "Tables.xlsx"
            save_excel_path(str(config), str(workbook))
            self.assertEqual(str(workbook.resolve()), load_excel_path(str(config)))

    def test_embedded_workbook_is_created_once_and_not_overwritten(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            template = root / "template.xlsx"
            workbook = root / "Data" / "Other" / "Excel" / "Tables.xlsx"
            make_workbook(template)

            created = ensure_embedded_excel_workbook(str(template), str(workbook))
            self.assertEqual(str(workbook.resolve()), created)
            workbook.write_bytes(b"user data")
            ensure_embedded_excel_workbook(str(template), str(workbook))
            self.assertEqual(b"user data", workbook.read_bytes())

    def test_resolver_prefers_external_override_and_can_reset_to_embedded(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = root / "excel_export.json"
            template = root / "template.xlsx"
            embedded = root / "embedded" / "Tables.xlsx"
            external = root / "external.xlsx"
            make_workbook(template)
            make_workbook(external)

            save_excel_path(str(config), str(external))
            self.assertEqual(
                str(external.resolve()),
                resolve_excel_workbook(str(config), str(template), str(embedded)),
            )
            save_excel_path(str(config), "")
            self.assertEqual(
                str(embedded.resolve()),
                resolve_excel_workbook(str(config), str(template), str(embedded)),
            )
            self.assertTrue(embedded.is_file())


if __name__ == "__main__":
    unittest.main()
