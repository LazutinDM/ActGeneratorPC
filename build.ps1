param(
    [switch]$SkipInstall,
    [string]$CompatibilityRuntimeDir = $env:ACTGENERATOR_COMPAT_RUNTIME
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectDir

# COLLEAGUE EDIT POINT: keep the Windows portable runtime reproducible.
$BuildPython = Join-Path $ProjectDir ".python-3.12.10\python.exe"
if (-not (Test-Path -LiteralPath $BuildPython -PathType Leaf)) {
    $BuildPython = (Get-Command python -ErrorAction Stop).Source
}
$BuildPythonVersion = (& $BuildPython -c "import platform; print(platform.python_version())").Trim()
if ($BuildPythonVersion -ne "3.12.10") {
    throw "Portable build requires official Windows Python 3.12.10, found $BuildPythonVersion"
}

if (-not $SkipInstall) {
    & $BuildPython -m pip install --upgrade pip
    & $BuildPython -m pip install -r requirements-build.txt
}

& $BuildPython -m PyInstaller --clean --noconfirm ActGeneratorPC.spec
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed with exit code $LASTEXITCODE"
}

# COLLEAGUE EDIT POINT: portable folder, EXE and release archive names.
$PortableDir = Join-Path $ProjectDir "dist\ActGeneratorPC"
$BuiltExe = Join-Path $PortableDir "ActGeneratorPC.exe"
$PortableDataDir = Join-Path $PortableDir "Data"
$PortableConfigDir = Join-Path $PortableDataDir "Other\Configuration"
$PortableHistoryDir = Join-Path $PortableDataDir "Other\History"
$PortableExcelDir = Join-Path $PortableDataDir "Other\Excel"
$PortableExcelWorkbook = Join-Path $PortableExcelDir "Tables.xlsx"
$PortableExcelTemplate = Join-Path $PortableDataDir "Templates\Tables.xlsx"

# COLLEAGUE EDIT POINT: on the main workstation, PyInstaller's freshly collected
# runtime is replaced with the runtime from a portable build that passed a real
# GUI startup test. GitHub Actions has no X: drive, so it keeps the clean runtime
# produced from the pinned Python and PySide versions instead.
if ([string]::IsNullOrWhiteSpace($CompatibilityRuntimeDir)) {
    $LocalVerifiedRuntime = "X:\MySoftware\ActGeneratorPC\Portable\Data"
    if (Test-Path -LiteralPath $LocalVerifiedRuntime -PathType Container) {
        $CompatibilityRuntimeDir = $LocalVerifiedRuntime
    }
}

if (-not (Test-Path -LiteralPath $BuiltExe -PathType Leaf)) {
    throw "PyInstaller did not create $BuiltExe"
}
New-Item -ItemType Directory -Path $PortableDataDir -Force | Out-Null

$ApplicationDataNames = @("Icons", "Other", "Templates", "Variables")
if (-not [string]::IsNullOrWhiteSpace($CompatibilityRuntimeDir)) {
    $CompatibilityRuntimeDir = [System.IO.Path]::GetFullPath($CompatibilityRuntimeDir)
    $CompatibilityPython = Join-Path $CompatibilityRuntimeDir "python312.dll"
    # PyInstaller keeps Qt runtime DLLs inside Data\PySide6. Older builds
    # checked the Data root and incorrectly rejected a complete portable.
    $CompatibilityQt = Join-Path $CompatibilityRuntimeDir "PySide6\Qt6Core.dll"
    if (-not (Test-Path -LiteralPath $CompatibilityPython -PathType Leaf) -or
        -not (Test-Path -LiteralPath $CompatibilityQt -PathType Leaf)) {
        throw "Compatible runtime is incomplete: $CompatibilityRuntimeDir"
    }
    Get-ChildItem -LiteralPath $PortableDataDir -Force | Where-Object {
        $ApplicationDataNames -notcontains $_.Name
    } | Remove-Item -Recurse -Force
    Get-ChildItem -LiteralPath $CompatibilityRuntimeDir -Force | Where-Object {
        $ApplicationDataNames -notcontains $_.Name
    } | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $PortableDataDir -Recurse -Force
    }
}

if (Test-Path -LiteralPath "Data") {
    Get-ChildItem -LiteralPath "Data" -Force | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $PortableDataDir -Recurse -Force
    }
}

# Never ship development/test history or data copied from another installation.
# The portable package starts with its own empty per-document history directory.
if (Test-Path -LiteralPath $PortableHistoryDir) {
    Remove-Item -LiteralPath $PortableHistoryDir -Recurse -Force
}
New-Item -ItemType Directory -Path $PortableHistoryDir -Force | Out-Null

New-Item -ItemType Directory -Path $PortableConfigDir -Force | Out-Null
Copy-Item -LiteralPath "update_config.json" -Destination $PortableConfigDir -Force

# Every portable package includes a ready-to-open workbook. Do not ship a
# developer-specific external path or a workbook containing local act data.
$PortableExcelConfig = Join-Path $PortableConfigDir "excel_export.json"
if (Test-Path -LiteralPath $PortableExcelConfig) {
    Remove-Item -LiteralPath $PortableExcelConfig -Force
}
if (-not (Test-Path -LiteralPath $PortableExcelTemplate -PathType Leaf)) {
    throw "Excel template was not copied to $PortableExcelTemplate"
}
New-Item -ItemType Directory -Path $PortableExcelDir -Force | Out-Null
Copy-Item -LiteralPath $PortableExcelTemplate -Destination $PortableExcelWorkbook -Force

New-Item -ItemType Directory -Path (Join-Path $PortableDir "Acts") -Force | Out-Null

$Archive = Join-Path $ProjectDir "dist\ActGeneratorPC-portable.zip"
if (Test-Path -LiteralPath $Archive) {
    Remove-Item -LiteralPath $Archive -Force
}
Compress-Archive -Path (Join-Path $PortableDir "*") -DestinationPath $Archive -CompressionLevel Optimal

Write-Host ""
Write-Host "Portable build:"
Write-Host "  $Archive"
