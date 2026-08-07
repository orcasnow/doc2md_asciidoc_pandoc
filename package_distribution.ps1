param(
    [string]$PackageName = "doc2md-drop-distribution.zip",
    [string]$OutputDir = "dist",
    [switch]$NoInstallers
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $Root

$outputPath = Join-Path $Root $OutputDir
if (-not (Test-Path -LiteralPath $outputPath)) {
    New-Item -ItemType Directory -Path $outputPath | Out-Null
}

$files = @(
    (Join-Path $Root "dist\doc2md-drop.exe"),
    (Join-Path $Root "dist\doc2md-drop.ini"),
    (Join-Path $Root "はじめにお読みください.md"),
    (Join-Path $Root "DOC2MD_DROP_INI_GUIDE.md")
)

if (-not $NoInstallers) {
    $files += @(
        (Join-Path $Root "pandoc-3.10-windows-x86_64.msi"),
        (Join-Path $Root "tesseract-ocr-w64-setup-5.5.0.20241111.exe")
    )
}

$missing = @($files | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) })
if ($missing.Count -gt 0) {
    $missingText = $missing -join [Environment]::NewLine
    throw "Package input files are missing:$([Environment]::NewLine)$missingText"
}

$packagePath = Join-Path $outputPath $PackageName
Compress-Archive -LiteralPath $files -DestinationPath $packagePath -Force

$zip = Get-Item -LiteralPath $packagePath
Write-Host "Package created: $($zip.FullName)"
Write-Host "Size: $([Math]::Round($zip.Length / 1MB, 1)) MB"
Write-Host ""
Write-Host "Contents:"
foreach ($file in $files) {
    $item = Get-Item -LiteralPath $file
    Write-Host " - $($item.Name)"
}
