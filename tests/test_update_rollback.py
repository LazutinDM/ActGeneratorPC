import os
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

from update_manager import UpdateManager


@unittest.skipUnless(os.name == "nt", "Windows updater integration test")
class UpdateRollbackTests(unittest.TestCase):
    def test_failed_startup_restores_previous_version_and_user_data(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as directory:
            root = Path(directory)
            destination = root / "ActGeneratorPC"
            work = root / "work"
            payload = work / "payload"
            destination.mkdir()
            payload.mkdir(parents=True)

            # where.exe exits immediately because updater health arguments are
            # not valid search patterns, simulating a failed application start.
            system_exe = Path(os.environ["WINDIR"]) / "System32" / "where.exe"
            shutil.copy2(system_exe, destination / "ActGeneratorPC.exe")
            shutil.copy2(system_exe, payload / "ActGeneratorPC.exe")
            (destination / "old-version.txt").write_text("old", encoding="utf-8")
            (payload / "new-version.txt").write_text("new", encoding="utf-8")
            user_data = destination / "Data" / "Variables"
            user_data.mkdir(parents=True)
            (user_data / "user.txt").write_text("keep", encoding="utf-8")
            payload_defaults = payload / "Data" / "Variables"
            payload_defaults.mkdir(parents=True)
            (payload_defaults / "defaults.txt").write_text("replace", encoding="utf-8")
            (destination / "Data" / "python312.dll").write_text(
                "old runtime", encoding="utf-8"
            )
            (payload / "Data" / "python312.dll").write_text(
                "new runtime", encoding="utf-8"
            )

            script = UpdateManager()._write_apply_script(work)
            result = subprocess.run(
                [
                    "powershell.exe", "-NoProfile", "-NonInteractive",
                    "-ExecutionPolicy", "Bypass", "-File", str(script),
                    "-ProcessId", "2147483647",
                    "-Payload", str(payload),
                    "-Destination", str(destination),
                    "-Executable", "ActGeneratorPC.exe",
                    "-WorkDirectory", str(work),
                ],
                capture_output=True,
                timeout=15,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )

            self.assertNotEqual(0, result.returncode)
            self.assertTrue((destination / "old-version.txt").is_file())
            self.assertFalse((destination / "new-version.txt").exists())
            self.assertEqual(
                "keep",
                (destination / "Data" / "Variables" / "user.txt").read_text(
                    encoding="utf-8"
                ),
            )
            self.assertFalse(
                (destination / "Data" / "Variables" / "defaults.txt").exists()
            )
            self.assertEqual(
                "old runtime",
                (destination / "Data" / "python312.dll").read_text(encoding="utf-8"),
            )


if __name__ == "__main__":
    unittest.main()
