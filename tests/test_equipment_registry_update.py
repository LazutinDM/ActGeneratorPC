import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from equipment_registry_update import (
    RegistryUpdateError,
    allowed_stations_through_alabushevo,
    canonical_model,
    canonical_station,
    clean_identifier,
    update_registry,
)


class EquipmentRegistryUpdateTests(unittest.TestCase):
    def test_station_alias_and_cutoff(self):
        self.assertEqual("КРЮКОВО", canonical_station("Зеленоград-Крюково"))
        self.assertEqual(
            ["КРЮКОВО", "АЛАБУШЕВО"],
            allowed_stations_through_alabushevo(
                ["Зеленоград-Крюково", "АЛАБУШЕВО", "РЕДКИНО"]
            ),
        )

    def test_equipment_aliases(self):
        expected = {
            "АБП-БН": "АБП-БН",
            "АПБ": "АБП-09",
            "АН 15": "АН-15",
            "БПА-20-БН1": "БПА-20-БН1",
            "АПБ-09-М3": "АБП-09-М3",
            "МКТФ": "МКТФ",
        }
        for source, result in expected.items():
            with self.subTest(source=source):
                self.assertEqual(result, canonical_model(source))

    def test_excel_numeric_identifiers_are_plain_digits(self):
        self.assertEqual("52840", clean_identifier("52840.0"))
        self.assertEqual("180300010", clean_identifier("1.8030001E8"))

    def test_failed_download_preserves_installed_registry(self):
        with tempfile.TemporaryDirectory() as directory:
            registry = Path(directory) / "station_serials.json"
            original = [{"serial": "1", "model": "МКТФ", "station": "КРЮКОВО"}]
            registry.write_text(json.dumps(original), encoding="utf-8")
            with patch(
                "equipment_registry_update.download_workbook",
                side_effect=RegistryUpdateError("нет сети"),
            ):
                with self.assertRaises(RegistryUpdateError):
                    update_registry(registry, ["КРЮКОВО", "АЛАБУШЕВО"])
            self.assertEqual(original, json.loads(registry.read_text(encoding="utf-8")))
            self.assertFalse(registry.with_name("station_serials.previous.json").exists())


if __name__ == "__main__":
    unittest.main()
