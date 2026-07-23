param(
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectDir

if (-not $SkipInstall) {
    python -m pip install --upgrade pip
    python -m pip install -r requirements-build.txt
}

python -m PyInstaller --clean --noconfirm ActGeneratorPC.spec

$PortableDir = Join-Path $ProjectDir "dist\ActGeneratorPC"
Copy-Item -LiteralPath "update_config.json" -Destination $PortableDir -Force

foreach ($Folder in @("Data", "Templates")) {
    if (Test-Path -LiteralPath $Folder) {
        Copy-Item -LiteralPath $Folder -Destination $PortableDir -Recurse -Force
    }
    else {
        New-Item -ItemType Directory -Path (Join-Path $PortableDir $Folder) -Force | Out-Null
    }
}

New-Item -ItemType Directory -Path (Join-Path $PortableDir "Act_Ready") -Force | Out-Null

$Archive = Join-Path $ProjectDir "dist\ActGeneratorPC-portable.zip"
if (Test-Path -LiteralPath $Archive) {
    Remove-Item -LiteralPath $Archive -Force
}
Compress-Archive -Path (Join-Path $PortableDir "*") -DestinationPath $Archive -CompressionLevel Optimal

$Hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $Archive).Hash.ToLowerInvariant()
$ChecksumPath = $Archive + ".sha256"
Set-Content -LiteralPath $ChecksumPath -Value "$Hash  ActGeneratorPC-portable.zip" -Encoding ascii

Write-Host ""
Write-Host "Portable build:"
Write-Host "  $Archive"
Write-Host "  $ChecksumPath"
