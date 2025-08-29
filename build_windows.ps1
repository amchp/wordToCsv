# Minimal Windows build (PowerShell)
# Usage:
#   ./build_windows.ps1

$ErrorActionPreference = 'Stop'

# 1) Find pandoc (fallback to the default install path)
$pandocPath = { $pandocPath = 'C:\Program Files\Pandoc\pandoc.exe' }
if (-not (Test-Path $pandocPath)) { throw "Pandoc not found at: $pandocPath" }

Write-Host "Using pandoc at: $pandocPath"

# 2) Build GUI app, bundling pandoc (note the ; separator on Windows)
pyinstaller .\main.py --name "wordToCsv" --windowed --add-binary "$pandocPath;pandoc"

Write-Host "Done. EXE at: dist\wordToCsv\wordToCsv.exe"

