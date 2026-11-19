@echo off
cd /d "%~dp0"
echo ============================================================
echo NBA 回家日常
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
chcp 65001 >nul
python "%~dp0home_pipeline.py"
echo.
pause
