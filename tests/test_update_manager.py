import tempfile
import unittest
import zipfile
from pathlib import Path

from update_manager import safe_extract


class UpdateValidationTests(unittest.TestCase):
    def test_safe_extract_rejects_path_traversal(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            archive = root / "bad.zip"
            with zipfile.ZipFile(archive, "w") as bundle:
                bundle.writestr("../outside.txt", "bad")
            with self.assertRaises(ValueError):
                safe_extract(archive, root / "stage")

    def test_safe_extract_accepts_portable_payload(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            archive = root / "good.zip"
            with zipfile.ZipFile(archive, "w") as bundle:
                bundle.writestr("ActGeneratorPC.exe", b"x" * 2048)
                bundle.writestr("Data/Other/readme.txt", "ok")
            stage = root / "stage"
            stage.mkdir()
            safe_extract(archive, stage)
            self.assertTrue((stage / "ActGeneratorPC.exe").is_file())


if __name__ == "__main__":
    unittest.main()
