@echo off
cd /d "%~dp0"
echo ============================================================
echo 先跑回家日常，跑完就自動合成。請不要關掉這個視窗
echo ============================================================
echo.
echo 請先關閉 Excel：
echo   NBA26-27爬蟲數據.xlsx
echo   玖九盤口變化.xlsx
echo   盤口觀察.xlsx
echo.
if not exist "%~dp0home_pipeline.py" (
echo 找不到 home_pipeline.py，請放在這個 bat 的同一層。
pause
exit /b 1
)
if not exist "%~dp0merge_daily_report.py" (
echo 找不到 merge_daily_report.py
pause
exit /b 1
)
chcp 65001 >nul
python "%~dp0home_pipeline.py"
python "%~dp0merge_daily_report.py"
echo.
pause
