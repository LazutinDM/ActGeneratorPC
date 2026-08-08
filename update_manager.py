from __future__ import annotations

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

from PySide6.QtCore import QObject, Signal, Qt
from PySide6.QtWidgets import QApplication, QMessageBox, QProgressDialog

try:
    from version import __version__
except ModuleNotFoundError:
    # Allows the source version to start even if version.py was not copied.
    # GitHub Actions still generates version.py for every release build.
    __version__ = "1.0.0"


APP_NAME = "ActGeneratorPC"
CONFIG_NAME = "update_config.json"
CONFIG_RELATIVE_PATH = (
    Path("Data") / "Other" / "Configuration" / CONFIG_NAME
)
USER_AGENT = f"{APP_NAME}/{__version__}"
API_ACCEPT = "application/vnd.github+json"
VERSION_RE = re.compile(r"\d+")


def portable_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def confirm_healthy_startup() -> None:
    """Confirm a post-update launch to the waiting updater process."""
    try:
        token_index = sys.argv.index("--update-health-token")
        file_index = sys.argv.index("--update-health-file")
        token = sys.argv[token_index + 1].strip()
        health_file = Path(sys.argv[file_index + 1]).resolve()
        temp_root = Path(tempfile.gettempdir()).resolve()
        if (
            not re.fullmatch(r"[0-9a-fA-F]{32}", token)
            or temp_root not in health_file.parents
            or health_file.name != "startup-healthy"
        ):
            return
        temporary = health_file.with_suffix(".tmp")
        temporary.write_text(token, encoding="ascii")
        os.replace(temporary, health_file)
    except (ValueError, IndexError, OSError):
        pass


def version_key(value: str) -> tuple[int, ...]:
    numbers = tuple(int(part) for part in VERSION_RE.findall(value or ""))
    return numbers or (0,)


def read_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8-sig") as stream:
        data = json.load(stream)
    if not isinstance(data, dict):
        raise ValueError(f"{path.name} must contain a JSON object")
    return data


def is_certificate_error(error: BaseException) -> bool:
    details = f"{error!r} {error}".lower()
    return (
        "certificate_verify_failed" in details
        or "certificate verify failed" in details
        or "unable to get local issuer certificate" in details
    )


def download_with_windows_trust(
    url: str,
    destination: Path,
    accept: str,
) -> None:
    """Download through Windows TLS when Python cannot use the local CA."""
    script = Path(tempfile.gettempdir()) / (
        f"actgenerator-download-{uuid.uuid4().hex}.ps1"
    )
    script.write_text(
        r'''param(
    [Parameter(Mandatory=$true)][string]$Url,
    [Parameter(Mandatory=$true)][string]$Destination,
    [Parameter(Mandatory=$true)][string]$Accept,
    [Parameter(Mandatory=$true)][string]$UserAgent
)
$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Headers = @{
    "Accept" = $Accept
    "X-GitHub-Api-Version" = "2022-11-28"
}
Invoke-WebRequest `
    -UseBasicParsing `
    -Uri $Url `
    -Headers $Headers `
    -UserAgent $UserAgent `
    -OutFile $Destination
''',
        encoding="utf-8-sig",
    )
    try:
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        result = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script),
                "-Url",
                url,
                "-Destination",
                str(destination),
                "-Accept",
                accept,
                "-UserAgent",
                USER_AGENT,
            ],
            capture_output=True,
            creationflags=creation_flags,
            check=False,
        )
        if result.returncode != 0:
            details = (result.stderr or result.stdout).decode(
                errors="replace"
            ).strip()
            raise RuntimeError(
                "Системная проверка сертификата Windows также завершилась "
                f"ошибкой:\n{details or 'неизвестная ошибка PowerShell'}"
            )
    finally:
        script.unlink(missing_ok=True)


def http_get_json(url: str) -> dict[str, Any]:
    request = urllib.request.Request(
        url,
        headers={
            "Accept": API_ACCEPT,
            "User-Agent": USER_AGENT,
            "X-GitHub-Api-Version": "2022-11-28",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            return json.load(response)
    except urllib.error.URLError as error:
        if not is_certificate_error(error):
            raise

    temporary = Path(tempfile.gettempdir()) / (
        f"actgenerator-release-{uuid.uuid4().hex}.json"
    )
    try:
        download_with_windows_trust(url, temporary, API_ACCEPT)
        return read_json(temporary)
    finally:
        temporary.unlink(missing_ok=True)


def download_file(url: str, destination: Path) -> None:
    request = urllib.request.Request(
        url,
        headers={"Accept": "application/octet-stream", "User-Agent": USER_AGENT},
    )
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            with destination.open("wb") as stream:
                shutil.copyfileobj(response, stream, length=1024 * 1024)
        return
    except urllib.error.URLError as error:
        if not is_certificate_error(error):
            raise

    download_with_windows_trust(
        url,
        destination,
        "application/octet-stream",
    )


def safe_extract(archive: Path, destination: Path) -> None:
    destination = destination.resolve()
    with zipfile.ZipFile(archive) as bundle:
        members = bundle.infolist()
        if not members or len(members) > 10000:
            raise ValueError("The update archive is empty or contains too many files.")
        total_size = sum(member.file_size for member in members)
        if total_size > 1024 * 1024 * 1024:
            raise ValueError("The unpacked update is larger than the safety limit.")
        for member in members:
            if member.flag_bits & 0x1:
                raise ValueError("Encrypted files are not allowed in an update.")
            unix_mode = member.external_attr >> 16
            if unix_mode and (unix_mode & 0o170000) == 0o120000:
                raise ValueError("Symbolic links are not allowed in an update.")
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
    manual_check_finished = Signal(object)
    manual_check_failed = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.window = parent
        self.root = portable_root()
        self.config_path = self.root / CONFIG_RELATIVE_PATH
        self._busy = False
        self._download_prompted = False
        self._check_progress = None
        self._download_progress = None

        self.release_ready.connect(self._handle_release_ready)
        self.update_downloaded.connect(self._apply_downloaded_update)
        self.update_failed.connect(self._show_download_error)
        self.manual_check_finished.connect(self._show_manual_check_result)
        self.manual_check_failed.connect(self._show_manual_check_error)

    def check_async(self, manual: bool = False) -> None:
        if self._busy:
            if manual:
                QMessageBox.information(
                    self.window,
                    "Проверка обновлений",
                    "Проверка обновлений уже выполняется.",
                )
            return

        try:
            config = read_json(self.config_path)
        except (OSError, ValueError, json.JSONDecodeError) as error:
            if manual:
                QMessageBox.warning(
                    self.window,
                    "Проверка обновлений",
                    f"Не удалось прочитать {CONFIG_NAME}:\n{error}",
                )
            return

        repository = str(config.get("repository", "")).strip()
        if (
            (not manual and not config.get("check_on_start", True))
            or repository in {"", "OWNER/REPOSITORY"}
            or repository.count("/") != 1
        ):
            if manual:
                QMessageBox.warning(
                    self.window,
                    "Проверка обновлений",
                    f"В файле {CONFIG_NAME} неверно указан репозиторий.",
                )
            return

        self._busy = True
        if manual:
            self._check_progress = QProgressDialog(
                "Проверка обновлений…",
                "",
                0,
                0,
                self.window,
            )
            self._check_progress.setWindowTitle("Обновление ActGeneratorPC")
            self._check_progress.setCancelButton(None)
            self._check_progress.setWindowModality(Qt.WindowModal)
            self._check_progress.setMinimumDuration(0)
            self._check_progress.show()
        threading.Thread(
            target=self._check_worker,
            args=(config, manual),
            name="github-update-check",
            daemon=True,
        ).start()

    def _check_worker(self, config: dict[str, Any], manual: bool = False) -> None:
        try:
            repository = str(config["repository"]).strip()
            release = http_get_json(
                f"https://api.github.com/repos/{repository}/releases/latest"
            )
            tag = str(release.get("tag_name", "")).strip()
            if version_key(tag) <= version_key(__version__):
                if manual:
                    self.manual_check_finished.emit(
                        {"current": __version__, "latest": tag or __version__}
                    )
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
            if not asset:
                if manual:
                    self.manual_check_failed.emit(
                        f"В Release {tag} отсутствует файл {asset_name}."
                    )
                return
            self.release_ready.emit(
                {
                    "tag": tag,
                    "name": str(release.get("name") or tag),
                    "notes": str(release.get("body") or "").strip(),
                    "asset_name": asset_name,
                    "asset_url": str(asset.get("browser_download_url", "")),
                }
            )
        except Exception as error:
            # A background check must never interrupt normal application startup.
            if manual:
                self.manual_check_failed.emit(str(error))
        finally:
            self._busy = False

    def _close_check_progress(self) -> None:
        if self._check_progress is not None:
            self._check_progress.close()
            self._check_progress.deleteLater()
            self._check_progress = None

    def _handle_release_ready(self, release: dict[str, Any]) -> None:
        self._close_check_progress()
        self._offer_update(release)

    def _show_manual_check_result(self, result: dict[str, str]) -> None:
        self._close_check_progress()
        QMessageBox.information(
            self.window,
            "Проверка обновлений",
            "Установлена последняя версия "
            f"{result.get('current', __version__)}.\n"
            f"Последняя доступная версия: {result.get('latest', __version__)}.",
        )

    def _show_manual_check_error(self, details: str) -> None:
        self._close_check_progress()
        QMessageBox.critical(
            self.window,
            "Ошибка проверки обновлений",
            "Не удалось проверить обновления:\n" + details,
        )

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
        self._download_progress = QProgressDialog(
            "Скачивание и проверка обновления…",
            "",
            0,
            0,
            self.window,
        )
        self._download_progress.setWindowTitle("Обновление ActGeneratorPC")
        self._download_progress.setCancelButton(None)
        self._download_progress.setWindowModality(Qt.WindowModal)
        self._download_progress.setMinimumDuration(0)
        self._download_progress.show()
        threading.Thread(
            target=self._download_worker,
            args=(release,),
            name="github-update-download",
            daemon=True,
        ).start()

    def _download_worker(self, release: dict[str, Any]) -> None:
        try:
            work_dir = Path(
                tempfile.mkdtemp(prefix="actgenerator-update-")
            )
            archive = work_dir / release["asset_name"]
            download_file(release["asset_url"], archive)
            if not archive.is_file() or archive.stat().st_size < 1024:
                raise ValueError("Downloaded update archive is missing or too small.")

            stage = work_dir / "payload"
            stage.mkdir()
            safe_extract(archive, stage)
            payload = payload_root(stage)
            executable = payload / "ActGeneratorPC.exe"
            if not executable.is_file() or executable.stat().st_size < 1024 * 1024:
                raise ValueError(
                    "The verified package does not contain a valid ActGeneratorPC.exe."
                )
            self.update_downloaded.emit(
                {
                    "work_dir": str(work_dir),
                    "payload": str(payload),
                }
            )
        except Exception as error:
            self.update_failed.emit(str(error))

    def _show_download_error(self, details: str) -> None:
        if self._download_progress is not None:
            self._download_progress.close()
            self._download_progress.deleteLater()
            self._download_progress = None
        if self._download_prompted:
            QMessageBox.critical(
                self.window,
                "Ошибка обновления",
                "Не удалось установить обновление:\n" + details,
            )
        self._download_prompted = False

    def _apply_downloaded_update(self, package: dict[str, str]) -> None:
        if self._download_progress is not None:
            self._download_progress.close()
            self._download_progress.deleteLater()
            self._download_progress = None
        if not getattr(sys, "frozen", False):
            shutil.rmtree(package["work_dir"], ignore_errors=True)
            QMessageBox.information(
                self.window,
                "Обновление ActGeneratorPC",
                "Автообновление применяется только к portable-сборке. "
                "Исходный проект не изменён.",
            )
            return
        try:
            script = self._write_apply_script(Path(package["work_dir"]))
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
                cwd=package["work_dir"],
            )
        except Exception as error:
            self._show_download_error(str(error))
            return

        QApplication.quit()

    def _write_apply_script(self, work_directory: Path) -> Path:
        script = work_directory / f"actgenerator-apply-{uuid.uuid4().hex}.ps1"
        script.write_text(
            r'''param(
    [Parameter(Mandatory=$true)][int]$ProcessId,
    [Parameter(Mandatory=$true)][string]$Payload,
    [Parameter(Mandatory=$true)][string]$Destination,
    [string]$Executable,
    [Parameter(Mandatory=$true)][string]$WorkDirectory
)
$ErrorActionPreference = "Stop"
$Token = [guid]::NewGuid().ToString("N")
$Parent = Split-Path -Parent ([IO.Path]::GetFullPath($Destination))
$Leaf = Split-Path -Leaf ([IO.Path]::GetFullPath($Destination))
$Stage = Join-Path $Parent (".$Leaf-stage-$Token")
$Backup = Join-Path $Parent ("$Leaf.backup")
$Failed = Join-Path $Parent (".$Leaf-failed-$Token")
$HealthFile = Join-Path $WorkDirectory "startup-healthy"
$Swapped = $false
try {
    Wait-Process -Id $ProcessId -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 700

    $Payload = [IO.Path]::GetFullPath($Payload)
    $Destination = [IO.Path]::GetFullPath($Destination)
    $DriveRoot = [IO.Path]::GetPathRoot($Destination)
    if (
        -not (Test-Path -LiteralPath $Payload -PathType Container) -or
        -not (Test-Path -LiteralPath (Join-Path $Payload "ActGeneratorPC.exe") -PathType Leaf)
    ) {
        throw "The update package does not contain ActGeneratorPC.exe."
    }
    if ($Destination.TrimEnd("\") -eq $DriveRoot.TrimEnd("\")) {
        throw "Refusing to update a drive root."
    }

    Set-Location -LiteralPath $Parent
    New-Item -ItemType Directory -Path $Stage -Force | Out-Null
    Get-ChildItem -LiteralPath $Payload -Force | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $Stage -Recurse -Force
    }

    # These paths contain user documents, local lists, and settings.
    # The release must never overwrite them.
    $Preserve = @(
        "Acts",
        "Act_Ready",
        "Data\Variables",
        "Data\Other\Configuration",
        "Data\Other\Excel"
    )
    foreach ($Relative in $Preserve) {
        $Current = Join-Path $Destination $Relative
        $Staged = Join-Path $Stage $Relative
        if (Test-Path -LiteralPath $Current) {
            if (Test-Path -LiteralPath $Staged) {
                Remove-Item -LiteralPath $Staged -Recurse -Force
            }
            New-Item -ItemType Directory -Path (Split-Path -Parent $Staged) -Force | Out-Null
            Copy-Item -LiteralPath $Current -Destination $Staged -Recurse -Force
        }
    }

    if (Test-Path -LiteralPath $Backup) {
        Remove-Item -LiteralPath $Backup -Recurse -Force
    }
    Move-Item -LiteralPath $Destination -Destination $Backup
    try {
        Move-Item -LiteralPath $Stage -Destination $Destination
        $Swapped = $true
    }
    catch {
        Move-Item -LiteralPath $Backup -Destination $Destination
        throw
    }

    if (-not $Executable) {
        exit 0
    }
    $Target = Join-Path $Destination $Executable
    $Process = Start-Process -FilePath $Target -WorkingDirectory $Destination -PassThru -ArgumentList @(
        "--update-health-token", $Token,
        "--update-health-file", $HealthFile
    )
    $Healthy = $false
    for ($Attempt = 0; $Attempt -lt 60; $Attempt++) {
        Start-Sleep -Milliseconds 500
        if (Test-Path -LiteralPath $HealthFile -PathType Leaf) {
            $ReportedToken = (Get-Content -LiteralPath $HealthFile -Raw).Trim()
            if ($ReportedToken -eq $Token) {
                $Healthy = $true
                break
            }
        }
        if ($Process.HasExited) {
            break
        }
        $Process.Refresh()
    }
    if (-not $Healthy) {
        if (-not $Process.HasExited) {
            Stop-Process -Id $Process.Id -Force -ErrorAction SilentlyContinue
        }
        Move-Item -LiteralPath $Destination -Destination $Failed
        Move-Item -LiteralPath $Backup -Destination $Destination
        $OldTarget = Join-Path $Destination $Executable
        if (Test-Path -LiteralPath $OldTarget) {
            Start-Process -FilePath $OldTarget -WorkingDirectory $Destination
        }
        Remove-Item -LiteralPath $Failed -Recurse -Force -ErrorAction SilentlyContinue
        throw "The updated application did not confirm a healthy startup; rollback completed."
    }
}
catch {
    if ($Swapped -and (Test-Path -LiteralPath $Backup)) {
        if (Test-Path -LiteralPath $Destination) {
            Move-Item -LiteralPath $Destination -Destination $Failed -ErrorAction SilentlyContinue
        }
        Move-Item -LiteralPath $Backup -Destination $Destination -ErrorAction SilentlyContinue
        if ($Executable) {
            $OldTarget = Join-Path $Destination $Executable
            if (Test-Path -LiteralPath $OldTarget) {
                Start-Process -FilePath $OldTarget -WorkingDirectory $Destination
            }
        }
        Remove-Item -LiteralPath $Failed -Recurse -Force -ErrorAction SilentlyContinue
    }
    elseif (-not (Test-Path -LiteralPath $Destination) -and (Test-Path -LiteralPath $Backup)) {
        Move-Item -LiteralPath $Backup -Destination $Destination -ErrorAction SilentlyContinue
    }
    throw
}
finally {
    Remove-Item -LiteralPath $Stage -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $WorkDirectory -Recurse -Force -ErrorAction SilentlyContinue
}
''',
            encoding="utf-8-sig",
        )
        return script
