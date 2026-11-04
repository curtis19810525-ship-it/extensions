@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

set "PY=python"
%PY% -c "import sys" 1>nul 2>nul
if errorlevel 1 set "PY=py"
%PY% -c "import sys" 1>nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python and enable Add python.exe to PATH.
  pause
  exit /b 1
)

%PY% -c "import openpyxl" 1>nul 2>nul
if errorlevel 1 (
  echo Installing openpyxl...
  %PY% -m pip install openpyxl
  if errorlevel 1 (
    echo Could not install openpyxl.
    pause
    exit /b 1
  )
)

%PY% "%~dp0append_nba_journal.py" --dry-run %*
if errorlevel 1 (
  pause
  exit /b 1
)

echo.
echo The check above is not written yet.
echo Press any key to write into the Excel file. Close this window to cancel.
pause >nul

%PY% "%~dp0append_nba_journal.py" %*
echo.
pause
