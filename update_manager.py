from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import urllib.error
import urllib.request
import uuid
import zipfile
from pathlib import Path
from typing import Any

from PySide6.QtCore import QObject, Signal
from PySide6.QtWidgets import QApplication, QMessageBox

try:
    from version import __version__
except ModuleNotFoundError:
    # Allows the source version to start even if version.py was not copied.
    # GitHub Actions still generates version.py for every release build.
    __version__ = "1.0.0"


APP_NAME = "ActGeneratorPC"
CONFIG_NAME = "update_config.json"
USER_AGENT = f"{APP_NAME}/{__version__}"
API_ACCEPT = "application/vnd.github+json"
VERSION_RE = re.compile(r"\d+")


def portable_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def version_key(value: str) -> tuple[int, ...]:
    numbers = tuple(int(part) for part in VERSION_RE.findall(value or ""))
    return numbers or (0,)


def read_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8-sig") as stream:
        data = json.load(stream)
    if not isinstance(data, dict):
        raise ValueError(f"{path.name} must contain a JSON object")
    return data


def http_get_json(url: str) -> dict[str, Any]:
    request = urllib.request.Request(
        url,
        headers={
            "Accept": API_ACCEPT,
            "User-Agent": USER_AGENT,
            "X-GitHub-Api-Version": "2022-11-28",
        },
    )
    with urllib.request.urlopen(request, timeout=20) as response:
        return json.load(response)


def download_file(url: str, destination: Path) -> None:
    request = urllib.request.Request(
        url,
        headers={"Accept": "application/octet-stream", "User-Agent": USER_AGENT},
    )
    with urllib.request.urlopen(request, timeout=90) as response:
        with destination.open("wb") as stream:
            shutil.copyfileobj(response, stream, length=1024 * 1024)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def safe_extract(archive: Path, destination: Path) -> None:
    destination = destination.resolve()
    with zipfile.ZipFile(archive) as bundle:
        for member in bundle.infolist():
            target = (destination / member.filename).resolve()
            if target != destination and destination not in target.parents:
                raise ValueError(f"Unsafe path in update archive: {member.filename}")
        bundle.extractall(destination)


def payload_root(stage: Path) -> Path:
    children = [item for item in stage.iterdir() if item.name != "__MACOSX"]
    if len(children) == 1 and children[0].is_dir():
        return children[0]
    return stage


class UpdateManager(QObject):
    release_ready = Signal(object)
    update_downloaded = Signal(object)
    update_failed = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.window = parent
        self.root = portable_root()
        self.config_path = self.root / CONFIG_NAME
        self._busy = False
        self._download_prompted = False

        self.release_ready.connect(self._offer_update)
        self.update_downloaded.connect(self._apply_downloaded_update)
        self.update_failed.connect(self._show_download_error)

    def check_async(self) -> None:
        if self._busy:
            return

        try:
            config = read_json(self.config_path)
        except (OSError, ValueError, json.JSONDecodeError):
            return

        repository = str(config.get("repository", "")).strip()
        if (
            not config.get("check_on_start", True)
            or repository in {"", "OWNER/REPOSITORY"}
            or repository.count("/") != 1
        ):
            return

        self._busy = True
        threading.Thread(
            target=self._check_worker,
            args=(config,),
            name="github-update-check",
            daemon=True,
        ).start()

    def _check_worker(self, config: dict[str, Any]) -> None:
        try:
            repository = str(config["repository"]).strip()
            release = http_get_json(
                f"https://api.github.com/repos/{repository}/releases/latest"
            )
            tag = str(release.get("tag_name", "")).strip()
            if version_key(tag) <= version_key(__version__):
                return

            asset_name = str(
                config.get("asset_name", "ActGeneratorPC-portable.zip")
            ).strip()
            assets = {
                str(item.get("name", "")): item
                for item in release.get("assets", [])
                if isinstance(item, dict)
            }
            asset = assets.get(asset_name)
            checksum = assets.get(asset_name + ".sha256")
            if not asset:
                return

            self.release_ready.emit(
                {
                    "tag": tag,
                    "name": str(release.get("name") or tag),
                    "notes": str(release.get("body") or "").strip(),
                    "asset_name": asset_name,
                    "asset_url": str(asset.get("browser_download_url", "")),
                    "checksum_url": (
                        str(checksum.get("browser_download_url", ""))
                        if checksum
                        else ""
                    ),
                    "require_checksum": bool(
                        config.get("require_checksum", True)
                    ),
                }
            )
        except Exception:
            # A background check must never interrupt normal application startup.
            pass
        finally:
            self._busy = False

    def _offer_update(self, release: dict[str, Any]) -> None:
        notes = release["notes"]
        if len(notes) > 1200:
            notes = notes[:1200].rstrip() + "\n…"

        message = QMessageBox(self.window)
        message.setIcon(QMessageBox.Information)
        message.setWindowTitle("Доступно обновление")
        message.setText(
            f"Доступна версия {release['tag']} "
            f"(установлена {__version__})."
        )
        message.setInformativeText(
            (notes + "\n\n" if notes else "")
            + "Обновить portable-приложение сейчас?"
        )
        install_button = message.addButton(
            "Скачать и обновить", QMessageBox.AcceptRole
        )
        message.addButton("Позже", QMessageBox.RejectRole)
        message.exec()
        if message.clickedButton() != install_button:
            return

        self._download_prompted = True
        threading.Thread(
            target=self._download_worker,
            args=(release,),
            name="github-update-download",
            daemon=True,
        ).start()

    def _download_worker(self, release: dict[str, Any]) -> None:
        try:
            if release["require_checksum"] and not release["checksum_url"]:
                raise RuntimeError(
                    "В релизе отсутствует обязательный файл SHA-256."
                )

            work_dir = Path(
                tempfile.mkdtemp(prefix="actgenerator-update-")
            )
            archive = work_dir / release["asset_name"]
            download_file(release["asset_url"], archive)

            if release["checksum_url"]:
                checksum_file = work_dir / (release["asset_name"] + ".sha256")
                download_file(release["checksum_url"], checksum_file)
                expected = checksum_file.read_text(
                    encoding="utf-8-sig"
                ).strip().split()[0].lower()
                actual = sha256_file(archive)
                if not re.fullmatch(r"[0-9a-f]{64}", expected):
                    raise RuntimeError("Некорректный файл SHA-256 в релизе.")
                if actual != expected:
                    raise RuntimeError(
                        "Контрольная сумма обновления не совпала."
                    )

            stage = work_dir / "payload"
            stage.mkdir()
            safe_extract(archive, stage)
            self.update_downloaded.emit(
                {
                    "work_dir": str(work_dir),
                    "payload": str(payload_root(stage)),
                }
            )
        except Exception as error:
            self.update_failed.emit(str(error))

    def _show_download_error(self, details: str) -> None:
        if self._download_prompted:
            QMessageBox.critical(
                self.window,
                "Ошибка обновления",
                "Не удалось установить обновление:\n" + details,
            )
        self._download_prompted = False

    def _apply_downloaded_update(self, package: dict[str, str]) -> None:
        try:
            script = self._write_apply_script()
            executable = (
                Path(sys.executable).name
                if getattr(sys, "frozen", False)
                else ""
            )
            creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            subprocess.Popen(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(script),
                    "-ProcessId",
                    str(os.getpid()),
                    "-Payload",
                    package["payload"],
                    "-Destination",
                    str(self.root),
                    "-Executable",
                    executable,
                    "-WorkDirectory",
                    package["work_dir"],
                ],
                close_fds=True,
                creationflags=creation_flags,
            )
        except Exception as error:
            self._show_download_error(str(error))
            return

        QApplication.quit()

    def _write_apply_script(self) -> Path:
        script = self.root / f".actgenerator-apply-{uuid.uuid4().hex}.ps1"
        script.write_text(
            r'''param(
    [Parameter(Mandatory=$true)][int]$ProcessId,
    [Parameter(Mandatory=$true)][string]$Payload,
    [Parameter(Mandatory=$true)][string]$Destination,
    [string]$Executable,
    [Parameter(Mandatory=$true)][string]$WorkDirectory
)
$ErrorActionPreference = "Stop"
try {
    Wait-Process -Id $ProcessId -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 700

    Get-ChildItem -LiteralPath $Payload -Force | ForEach-Object {
        if ($_.Name -ne "Act_Ready") {
            Copy-Item -LiteralPath $_.FullName -Destination $Destination -Recurse -Force
        }
    }

    if ($Executable) {
        $target = Join-Path $Destination $Executable
        if (Test-Path -LiteralPath $target) {
            Start-Process -FilePath $target -WorkingDirectory $Destination
        }
    }
}
finally {
    Remove-Item -LiteralPath $WorkDirectory -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
}
''',
            encoding="utf-8-sig",
        )
        return script
