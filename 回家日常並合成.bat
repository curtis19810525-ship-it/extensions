@echo off
cd /d "%~dp0"
echo ============================================================
echo NBA home, then merge. Do not close this window.
echo ============================================================
echo.
if not exist "%~dp0nba_home.py" (
echo Missing nba_home.py in this folder.
pause
exit /b 1
)
if not exist "%~dp0merge_daily_report.py" (
echo Missing merge_daily_report.py in this folder.
pause
exit /b 1
)
set NBA_NO_PAUSE=1
set NBA_SKIP_MERGE=1
call "%~dp0回家日常.bat"
set NBA_NO_PAUSE=
set NBA_SKIP_MERGE=
echo.
chcp 65001 >nul
python "%~dp0merge_daily_report.py"
echo.
pause
