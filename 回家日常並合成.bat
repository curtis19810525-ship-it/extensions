@echo off
cd /d "%~dp0"
echo ============================================================
echo 先跑回家日常，跑完會自動合成，請勿關掉視窗
echo ============================================================
echo.
if not exist "%~dp0回家日常.bat" (
echo 找不到回家日常.bat
pause
exit /b 1
)
if not exist "%~dp0merge_daily_report.py" (
echo 找不到 merge_daily_report.py
pause
exit /b 1
)
findstr /v /i /r /c:"^ *pause" "%~dp0回家日常.bat" > "%temp%\nba_home_nopause.bat"
cmd /c "%temp%\nba_home_nopause.bat"
echo.
echo 回家日常結束，開始合成。
chcp 65001 >nul
python "%~dp0merge_daily_report.py"
echo.
pause
