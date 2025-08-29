@echo off
REM Minimal Windows build (CMD)
setlocal

for /f "usebackq delims=" %%i in (`powershell -NoProfile -Command "(Get-Command pandoc -ErrorAction SilentlyContinue).Source ?? 'C:\Program Files\Pandoc\pandoc.exe'"`) do set "PANDOC_PATH=%%i"

if not exist "%PANDOC_PATH%" (
  echo Pandoc not found at "%PANDOC_PATH%".
  exit /b 1
)

echo Using pandoc at: "%PANDOC_PATH%"
pyinstaller main.py --name "wordToCsv" --windowed --add-binary "%PANDOC_PATH%;pandoc"
echo Done. EXE at: dist\wordToCsv\wordToCsv.exe

