import tempfile
import unittest
from pathlib import Path

from history_store import clear_history_records, load_history_records, save_history_record


class HistoryStoreTests(unittest.TestCase):
    def test_each_act_uses_an_independent_metadata_file(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            history = root / "Data" / "Other" / "History"
            acts = root / "Acts"
            acts.mkdir(parents=True)
            first = acts / "first.docx"
            second = acts / "second.docx"
            first.write_bytes(b"one")
            second.write_bytes(b"two")
            data = {
                "date": "11.08.2026", "executor": "Петров",
                "location": "Химки", "model": "МКТФ", "serial": "100",
                "work": "Проверка", "issues": "", "done_work": "Готово",
                "materials": "", "qty": 0, "template": "",
            }
            save_history_record(history, root, first, data)
            save_history_record(history, root, second, {**data, "serial": "200"})

            self.assertEqual(2, len(list(history.glob("*.json"))))
            records = load_history_records(history, root, acts)
            self.assertEqual({"100", "200"}, {item["data"]["serial"] for item in records})
            self.assertTrue(all(Path(item["document"]).is_file() for item in records))

    def test_unindexed_docx_is_still_visible(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            acts = root / "Acts"
            acts.mkdir()
            document = acts / "legacy.docx"
            document.write_bytes(b"legacy")
            records = load_history_records(root / "History", root, acts)
            self.assertEqual(1, len(records))
            self.assertEqual("legacy.docx", records[0]["document_name"])

    def test_missing_docx_disappears_from_history(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            acts = root / "Acts"
            acts.mkdir()
            document = acts / "deleted.docx"
            document.write_bytes(b"docx")
            save_history_record(
                root / "History", root, document,
                {"date": "11.08.2026", "serial": "42"},
            )
            document.unlink()

            self.assertEqual([], load_history_records(root / "History", root, acts))

    def test_clear_history_keeps_documents_and_hides_old_records(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            acts = root / "Acts"
            history = root / "History"
            acts.mkdir()
            document = acts / "kept.docx"
            document.write_bytes(b"docx")
            save_history_record(
                history, root, document,
                {"date": "11.08.2026", "serial": "42"},
            )

            self.assertEqual(1, clear_history_records(history))
            self.assertTrue(document.is_file())
            self.assertEqual([], load_history_records(history, root, acts))


if __name__ == "__main__":
    unittest.main()
