import unittest
from pathlib import Path

from app_paths import LEGACY_ROOT_FOLDERS


class PortableLayoutTests(unittest.TestCase):
    def test_pyinstaller_runtime_is_never_treated_as_legacy_data(self):
        normalized = {name.casefold() for name in LEGACY_ROOT_FOLDERS}
        self.assertNotIn("_internal", normalized)

    def test_pyside_version_is_pinned_for_windows_portable_compatibility(self):
        project = Path(__file__).resolve().parents[1]
        requirements = (project / "requirements.txt").read_text(encoding="utf-8")
        build_requirements = (project / "requirements-build.txt").read_text(
            encoding="utf-8"
        )
        build_script = (project / "build.ps1").read_text(encoding="utf-8")
        self.assertIn("PySide6==6.8.3", requirements.splitlines())
        self.assertIn("PyInstaller==6.21.0", build_requirements.splitlines())
        self.assertIn(
            "pyinstaller-hooks-contrib==2026.6",
            build_requirements.splitlines(),
        )
        self.assertIn('$BuildPythonVersion -ne "3.12.10"', build_script)
        self.assertIn('if ($LASTEXITCODE -ne 0)', build_script)
        self.assertIn("ACTGENERATOR_COMPAT_RUNTIME", build_script)
        self.assertIn('$ApplicationDataNames = @("Icons", "Other", "Templates", "Variables")', build_script)

    def test_single_executable_build_embeds_qt_runtime(self):
        """The distributable EXE must not depend on replaceable external Qt DLLs."""
        project = Path(__file__).resolve().parents[1]
        onefile_spec = (project / "ActGeneratorPC-onefile.spec").read_text(
            encoding="utf-8"
        )
        build_script = (project / "build.ps1").read_text(encoding="utf-8")

        self.assertIn("a.binaries", onefile_spec)
        self.assertIn("a.datas", onefile_spec)
        self.assertIn("runtime_dll_search.py", onefile_spec)
        self.assertIn("[switch]$SingleExeRuntime", build_script)
        self.assertIn("ActGeneratorPC-onefile.spec", build_script)


if __name__ == "__main__":
    unittest.main()
