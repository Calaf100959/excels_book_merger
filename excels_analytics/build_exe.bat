@echo off
setlocal
cd /d %~dp0

where py >nul 2>&1
if %errorlevel% neq 0 (
  echo [ERROR] Python launcher 'py' が見つかりません。
  exit /b 1
)

if not exist .venv (
  py -m venv .venv || exit /b 1
)

call .venv\Scripts\activate || exit /b 1
py -m pip install --upgrade pip || exit /b 1
py -m pip install -r requirements.txt || exit /b 1
py -m pip install -r requirements-build.txt || exit /b 1

py -m PyInstaller --noconfirm --clean --windowed --onefile --name ExcelMerger --add-data "merge_excel_sheets.ps1;." excel_merger_gui.py || exit /b 1

if not exist dist\ExcelMerger.exe (
  echo [ERROR] EXEが生成されませんでした。
  exit /b 1
)

echo.
echo Build completed.
echo EXE: dist\ExcelMerger.exe
endlocal
