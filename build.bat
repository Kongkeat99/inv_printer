@echo off
REM ============================================================
REM  Step 1 - build the app folder   ->  dist\InvoicePrinter\InvoicePrinter.exe
REM  (run build_installer.bat afterwards to make the Setup.exe)
REM ============================================================
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python is not installed or not in PATH. Install Python 3.10+ from python.org ^(tick "Add to PATH"^).
  pause & exit /b 1
)

echo [1/3] Installing build requirements...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt || (echo pip failed & pause & exit /b 1)

echo [2/3] Creating icon...
python make_icon.py

echo [3/3] Building InvoicePrinter ^(folder build - fast start, fewer antivirus warnings^)...
python -m PyInstaller --noconfirm --clean --onedir --windowed ^
  --name InvoicePrinter --icon app.ico ^
  --hidden-import win32timezone --hidden-import pythoncom --hidden-import pywintypes ^
  --hidden-import win32print --hidden-import win32com.client ^
  --collect-submodules pypdf --collect-submodules openpyxl ^
  invoice_printer.py
if errorlevel 1 (echo BUILD FAILED & pause & exit /b 1)

echo.
echo DONE.  App folder: %~dp0dist\InvoicePrinter\
echo Test it:  dist\InvoicePrinter\InvoicePrinter.exe
echo Next   :  build_installer.bat
pause
