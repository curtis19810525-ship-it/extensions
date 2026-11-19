@echo off
cd /d "%~dp0"
chcp 65001 >nul
python "%~dp0nba_home.py"
if not "%NBA_SKIP_MERGE%"=="1" python "%~dp0merge_daily_report.py"
if "%NBA_NO_PAUSE%"=="1" exit /b 0
echo.
pause
