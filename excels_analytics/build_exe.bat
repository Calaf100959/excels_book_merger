@echo off
setlocal
cd /d %~dp0

if not exist .venv (
  python -m venv .venv
)

call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install -r requirements-build.txt

pyinstaller --noconfirm --clean --windowed --onefile --name ExcelMerger --add-data "merge_excel_sheets.ps1;." excel_merger_gui.py

echo.
echo Build completed.
echo EXE: dist\ExcelMerger.exe
endlocal
