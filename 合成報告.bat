@echo off
cd /d "%~dp0"
if not exist "%~dp0merge_daily_report.py" (
echo 找不到 merge_daily_report.py
echo 請放到跟這個 bat 同一層，不要放進 exports。
pause
exit /b 1
)
echo 正在合成，請稍候。
chcp 65001 >nul
python "%~dp0merge_daily_report.py"
echo.
pause
