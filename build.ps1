param([switch]$SkipInstall)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectDir

if (-not $SkipInstall) {
    python -m pip install --upgrade pip
    python -m pip install -r requirements-build.txt
}

python -m PyInstaller --clean --noconfirm ActGeneratorPC.spec

$PortableDir = Join-Path $ProjectDir "dist\ActGeneratorPC"
$BuiltExe = Join-Path $PortableDir "ActGeneratorPC.exe"
$PortableDataDir = Join-Path $PortableDir "Data"
$PortableConfigDir = Join-Path $PortableDataDir "Other\Configuration"

if (-not (Test-Path -LiteralPath $BuiltExe -PathType Leaf)) {
    throw "PyInstaller did not create $BuiltExe"
}
New-Item -ItemType Directory -Path $PortableDataDir -Force | Out-Null

if (Test-Path -LiteralPath "Data") {
    Get-ChildItem -LiteralPath "Data" -Force | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $PortableDataDir -Recurse -Force
    }
}

New-Item -ItemType Directory -Path $PortableConfigDir -Force | Out-Null
Copy-Item -LiteralPath "update_config.json" -Destination $PortableConfigDir -Force

New-Item -ItemType Directory -Path (Join-Path $PortableDir "Acts") -Force | Out-Null

$Archive = Join-Path $ProjectDir "dist\ActGeneratorPC-portable.zip"
if (Test-Path -LiteralPath $Archive) {
    Remove-Item -LiteralPath $Archive -Force
}
Compress-Archive -Path (Join-Path $PortableDir "*") -DestinationPath $Archive -CompressionLevel Optimal

Write-Host ""
Write-Host "Portable build:"
Write-Host "  $Archive"
