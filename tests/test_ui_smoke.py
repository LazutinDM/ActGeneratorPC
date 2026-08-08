import os
import tempfile
import threading
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QRectF
from PySide6.QtGui import QPainter, QPdfWriter
from PySide6.QtWidgets import QApplication, QDialog, QMessageBox

from app import ActPreviewDialog, MainWindow


class UiSmokeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.application = QApplication.instance() or QApplication([])

    def test_main_window_builds_all_primary_controls(self):
        window = MainWindow()
        try:
            self.assertEqual("ActGeneratorPC", window.windowTitle())
            self.assertIsNotNone(window.btn_create)
            self.assertIsNotNone(window.cb_executor)
            self.assertIsNotNone(window.cb_location)
            self.assertIsNotNone(window.cb_model)
            self.assertIsNotNone(window.act_select_excel)
            self.assertIsNotNone(window.act_open_excel)
            self.assertTrue(window.act_preview_enabled.isCheckable())
        finally:
            window.close()

    def test_preview_menu_setting_is_loaded_and_persisted(self):
        with (
            patch("app.load_app_settings", return_value={"preview_before_save": False}),
            patch("app.save_app_settings") as save_settings,
        ):
            window = MainWindow()
            try:
                self.assertFalse(window.act_preview_enabled.isChecked())
                window.act_preview_enabled.setChecked(True)
                save_settings.assert_called_once_with({"preview_before_save": True})
            finally:
                window.close()

    def test_disabled_preview_saves_without_opening_dialog(self):
        class UnexpectedPreview:
            def __init__(self, *_args, **_kwargs):
                raise AssertionError("Preview dialog must not be created")

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "Acts"
            window = MainWindow()
            try:
                window.app_settings["preview_before_save"] = False
                window.cb_location.setEditText("Химки")
                window.cb_model.setEditText("МКТФ")
                window.cb_template.setCurrentIndex(1)

                excel_result = SimpleNamespace(sheet_name="Август 26", row_number=2)
                with (
                    patch("app.OUTPUT_DIR", str(output)),
                    patch("app.ActPreviewDialog", UnexpectedPreview),
                    patch("app.load_excel_path", return_value="Tables.xlsx"),
                    patch("app.append_act_to_workbook", return_value=excel_result),
                    patch.object(QMessageBox, "exec", return_value=0),
                    patch.object(QMessageBox, "clickedButton", return_value=None),
                ):
                    window.create_act()

                self.assertEqual(1, len(list(output.glob("*.docx"))))
            finally:
                window.close()

    def test_visual_preview_loads_a_real_pdf_page(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pdf_path = root / "preview.pdf"
            dummy_docx = root / "preview.docx"
            dummy_docx.write_bytes(b"test")

            writer = QPdfWriter(str(pdf_path))
            painter = QPainter(writer)
            painter.drawText(QRectF(100, 100, 800, 200), "Act preview smoke test")
            painter.end()
            del painter
            del writer

            release_renderer = threading.Event()

            def delayed_renderer(_path):
                release_renderer.wait(2)
                return str(pdf_path)

            with patch("app.render_docx_to_pdf", side_effect=delayed_renderer):
                dialog = ActPreviewDialog(str(dummy_docx))
                try:
                    self.assertTrue(dialog.render_thread.isRunning())
                    self.assertIsNone(dialog.pdf_document)
                    self.assertFalse(dialog.open_button.isEnabled())
                    release_renderer.set()
                    self.assertTrue(dialog.render_thread.wait(3000))
                    self.application.processEvents()
                    self.assertEqual(1, dialog.pdf_document.pageCount())
                    self.assertIs(dialog.pdf_document, dialog.preview.document())
                    self.assertTrue(dialog.open_button.isEnabled())
                    self.assertTrue(dialog.save_button.isEnabled())
                finally:
                    dialog.reject()

    def test_form_creates_docx_and_passes_derived_excel_values(self):
        class AcceptedPreview:
            def __init__(self, _path, _parent=None):
                pass

            def exec(self):
                return QDialog.Accepted

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "Acts"
            window = MainWindow()
            try:
                window.ed_date.setText("08.08.2026")
                window.cb_executor.setEditText("Петров")
                window.cb_location.setEditText("Химки")
                window.cb_model.setEditText("МКТФ")
                window.ed_serial.setText("52881")
                window.cb_work.setEditText("Замена блока питания")
                window.cb_issues.setEditText("Сбой")
                window.cb_done_work.setEditText("Блок заменён")
                window.cb_materials.setEditText("Комплект МТТПК")
                window.sp_qty.setValue(1)
                window.cb_template.setCurrentIndex(1)

                excel_result = SimpleNamespace(sheet_name="Август 26", row_number=2)
                with (
                    patch("app.OUTPUT_DIR", str(output)),
                    patch("app.ActPreviewDialog", AcceptedPreview),
                    patch("app.load_excel_path", return_value="Tables.xlsx"),
                    patch("app.append_act_to_workbook", return_value=excel_result) as append,
                    patch.object(QMessageBox, "exec", return_value=0),
                    patch.object(QMessageBox, "clickedButton", return_value=None),
                ):
                    window.create_act()

                documents = list(output.glob("*.docx"))
                self.assertEqual(1, len(documents))
                values = append.call_args.args[1]
                self.assertEqual("Крюковский", values["area"])
                self.assertEqual("выданно МТППК", values["notes"])
                self.assertEqual("Химки", values["location"])
            finally:
                window.close()


if __name__ == "__main__":
    unittest.main()
