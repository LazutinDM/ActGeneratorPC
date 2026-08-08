"""Create a PDF preview of a generated DOCX without changing the document."""

from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path


class PreviewRenderError(RuntimeError):
    """Raised when no supported DOCX renderer can create the preview."""


_WORD_EXPORT_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$inputPath = [Environment]::GetEnvironmentVariable('ACTGEN_PREVIEW_INPUT')
$outputPath = [Environment]::GetEnvironmentVariable('ACTGEN_PREVIEW_OUTPUT')
$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($inputPath, $false, $true)
    $document.ExportAsFixedFormat($outputPath, 17)
}
finally {
    if ($document -ne $null) { $document.Close($false) }
    if ($word -ne $null) { $word.Quit() }
    if ($document -ne $null) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
    }
    if ($word -ne $null) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
"""


def _subprocess_flags() -> int:
    return getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _render_with_word(docx_path: Path, pdf_path: Path) -> str | None:
    if os.name != "nt":
        return "Microsoft Word COM доступен только в Windows."

    powershell = shutil.which("powershell.exe") or shutil.which("powershell")
    if not powershell:
        return "PowerShell не найден."

    env = os.environ.copy()
    env["ACTGEN_PREVIEW_INPUT"] = str(docx_path)
    env["ACTGEN_PREVIEW_OUTPUT"] = str(pdf_path)
    try:
        result = subprocess.run(
            [
                powershell,
                "-NoLogo",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                _WORD_EXPORT_SCRIPT,
            ],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            creationflags=_subprocess_flags(),
            timeout=60,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        return f"Word: {exc}"

    if result.returncode == 0 and pdf_path.is_file() and pdf_path.stat().st_size:
        return None
    details = (result.stderr or result.stdout or "неизвестная ошибка").strip()
    return f"Word: {details}"


def _libreoffice_candidates() -> list[str]:
    candidates = [shutil.which("soffice"), shutil.which("soffice.exe")]
    if os.name == "nt":
        for variable in ("ProgramFiles", "ProgramFiles(x86)"):
            root = os.environ.get(variable)
            if root:
                candidates.append(str(Path(root) / "LibreOffice" / "program" / "soffice.exe"))
    return [candidate for candidate in dict.fromkeys(candidates) if candidate and Path(candidate).is_file()]


def _render_with_libreoffice(docx_path: Path, pdf_path: Path) -> str | None:
    candidates = _libreoffice_candidates()
    if not candidates:
        return "LibreOffice не найден."

    generated_path = pdf_path.parent / f"{docx_path.stem}.pdf"
    try:
        result = subprocess.run(
            [
                candidates[0],
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(pdf_path.parent),
                str(docx_path),
            ],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=_subprocess_flags(),
            timeout=60,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        return f"LibreOffice: {exc}"

    if result.returncode == 0 and generated_path.is_file() and generated_path.stat().st_size:
        if generated_path != pdf_path:
            os.replace(generated_path, pdf_path)
        return None
    details = (result.stderr or result.stdout or "неизвестная ошибка").strip()
    return f"LibreOffice: {details}"


def render_docx_to_pdf(docx_path: str, pdf_path: str | None = None) -> str:
    """Render *docx_path* to PDF using Word, then LibreOffice as a fallback."""
    source = Path(docx_path).resolve()
    if not source.is_file():
        raise PreviewRenderError(f"Временный DOCX не найден: {source}")

    destination = Path(pdf_path).resolve() if pdf_path else source.with_suffix(".pdf")
    destination.parent.mkdir(parents=True, exist_ok=True)
    try:
        destination.unlink(missing_ok=True)
    except OSError as exc:
        raise PreviewRenderError(f"Не удалось подготовить PDF предпросмотра: {exc}") from exc

    errors = []
    for renderer in (_render_with_word, _render_with_libreoffice):
        error = renderer(source, destination)
        if error is None:
            return str(destination)
        errors.append(error)

    raise PreviewRenderError(
        "Не удалось создать визуальный предпросмотр. Нужен Microsoft Word "
        "или LibreOffice.\n\n" + "\n".join(errors)
    )
