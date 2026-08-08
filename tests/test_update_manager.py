import hashlib
import tempfile
import unittest
import zipfile
from pathlib import Path

from update_manager import expected_sha256, safe_extract


class UpdateValidationTests(unittest.TestCase):
    def test_checksum_parser_and_hash(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            archive = root / "ActGeneratorPC-portable.zip"
            archive.write_bytes(b"verified payload")
            digest = hashlib.sha256(archive.read_bytes()).hexdigest()
            checksum = root / (archive.name + ".sha256")
            checksum.write_text(f"{digest}  {archive.name}\n", encoding="ascii")
            self.assertEqual(digest, expected_sha256(checksum, archive.name))

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
