@echo off
cd /d "%~dp0"
echo ============================================================
echo 先跑回家日常，結束後自動合成報告
echo ============================================================
echo.
echo 請放在「抓取NBA網路資料」，跟 exports 同一層。
echo 不要改名、不要放進 exports。
echo.
if not exist "%~dp0回家日常.bat" (
echo 找不到回家日常.bat
pause
exit /b 1
)
if not exist "%~dp0merge_daily_report.py" (
echo 找不到 merge_daily_report.py，請放到同一個資料夾。
pause
exit /b 1
)
cmd /c "%~dp0回家日常.bat"
echo.
echo 回家日常結束，開始合成。
chcp 65001 >nul
python "%~dp0merge_daily_report.py"
echo.
pause
