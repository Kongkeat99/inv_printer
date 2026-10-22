@echo off
REM ============================================================
REM  Step 2 - make the installer   ->  installer_output\InvoicePrinter_Setup_x.y.z.exe
REM  Needs Inno Setup 6 (free): https://jrsoftware.org/isdl.php
REM ============================================================
cd /d "%~dp0"

if not exist dist\InvoicePrinter\InvoicePrinter.exe (
  echo dist\InvoicePrinter\InvoicePrinter.exe not found - run build.bat first.
  pause & exit /b 1
)

REM version comes from APP_VERSION in invoice_printer.py
for /f %%v in ('python get_version.py') do set APPVER=%%v
echo Version: %APPVER%

set ISCC=
for %%p in ("%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" "%ProgramFiles%\Inno Setup 6\ISCC.exe" "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe") do (
  if exist %%p set ISCC=%%~p
)
if "%ISCC%"=="" (
  echo Inno Setup 6 not found. Install it from https://jrsoftware.org/isdl.php and run this again.
  pause & exit /b 1
)

"%ISCC%" /DAppVersion=%APPVER% installer.iss
if errorlevel 1 (echo INSTALLER BUILD FAILED & pause & exit /b 1)

echo.
echo DONE.  Give users:  installer_output\InvoicePrinter_Setup_%APPVER%.exe
pause
