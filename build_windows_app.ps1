param(
    [switch]$OneFile,
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $Root

if (-not $SkipInstall) {
    python -m pip install -r requirements.txt
    python -m pip install -r requirements-build.txt
}

$mode = "--onedir"
if ($OneFile) {
    $mode = "--onefile"
}

$pyinstallerArgs = @(
    $mode,
    "--noconfirm",
    "--clean",
    "--name", "doc2md-drop",
    "--console",
    "--paths", $Root,
    "--hidden-import", "doc2md_asciidoc_pandoc",
    "--hidden-import", "fitz",
    "--hidden-import", "pymupdf",
    "--hidden-import", "pymupdf4llm",
    "--hidden-import", "pytesseract",
    "--hidden-import", "PIL",
    "--hidden-import", "PIL.Image",
    "--hidden-import", "PIL.ImageEnhance",
    "--hidden-import", "pandas",
    "--hidden-import", "openpyxl",
    "--hidden-import", "xlrd",
    "--hidden-import", "mammoth",
    "--hidden-import", "markdownify",
    "--hidden-import", "tabulate",
    "--exclude-module", "IPython",
    "--exclude-module", "jupyter",
    "--exclude-module", "notebook",
    "--exclude-module", "matplotlib",
    "--exclude-module", "PyQt5",
    "--exclude-module", "PyQt6",
    "--exclude-module", "PySide2",
    "--exclude-module", "PySide6",
    "--exclude-module", "scipy",
    "--exclude-module", "dask",
    "--exclude-module", "distributed",
    "--exclude-module", "pyarrow",
    "--exclude-module", "numba",
    "--exclude-module", "llvmlite",
    "--exclude-module", "pygame",
    "--exclude-module", "pytest",
    "--exclude-module", "pandas.tests",
    "--exclude-module", "pymupdf.layout",
    "--exclude-module", "torch",
    "--exclude-module", "tensorflow",
    "--exclude-module", "keras",
    "--exclude-module", "onnxruntime",
    "doc2md_windows_drop.py"
)

python -m PyInstaller @pyinstallerArgs
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed with exit code $LASTEXITCODE"
}

$configSource = Join-Path $Root "doc2md-drop.ini"
if (Test-Path -LiteralPath $configSource) {
    if ($OneFile) {
        $configTarget = Join-Path (Join-Path $Root "dist") "doc2md-drop.ini"
    } else {
        $configTarget = Join-Path (Join-Path (Join-Path $Root "dist") "doc2md-drop") "doc2md-drop.ini"
    }
    Copy-Item -LiteralPath $configSource -Destination $configTarget -Force
    Write-Host "Config: $configTarget"
}

Write-Host ""
Write-Host "Build complete."
if ($OneFile) {
    Write-Host "EXE: $Root\dist\doc2md-drop.exe"
} else {
    Write-Host "EXE: $Root\dist\doc2md-drop\doc2md-drop.exe"
}
