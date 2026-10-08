import unittest
from pathlib import Path

from app_paths import (
    ICON_APP, ICON_BRAND, ICON_MODES, ICON_CALENDAR, ICON_DOCX,
    ICON_DONE_WORK, ICON_EQUIPMENT,
    ICON_EXCEL, ICON_ISSUES, ICON_LISTS, ICON_MATERIALS, ICON_QTY,
    ICON_SERIAL, ICON_SERVICES, ICON_STATION, ICON_USER, ICON_XLS,
)


class UiIconTests(unittest.TestCase):
    def test_original_user_icon_set_is_present(self):
        for icon_path in (
            ICON_CALENDAR, ICON_DOCX, ICON_DONE_WORK, ICON_EQUIPMENT,
            ICON_EXCEL, ICON_ISSUES, ICON_LISTS, ICON_MATERIALS, ICON_QTY,
            ICON_SERIAL, ICON_SERVICES, ICON_STATION, ICON_USER, ICON_XLS,
            ICON_BRAND, ICON_MODES,
        ):
            icon = Path(icon_path)
            self.assertTrue(icon.is_file(), icon.name)
            self.assertEqual(".png", icon.suffix.lower())

    def test_windows_application_icon_exists(self):
        icon = Path(ICON_APP)
        self.assertTrue(icon.is_file())
        self.assertGreater(icon.stat().st_size, 1_000)


if __name__ == "__main__":
    unittest.main()
