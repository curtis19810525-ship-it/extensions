@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo 合成賽後與賽前（不重跑爬蟲）
echo ============================================================
echo.
echo 請放在「抓取NBA網路資料」，跟 exports 資料夾同一層。
echo 不要放進 exports 裡面。
echo.

python "%~dp0merge_daily_report.py"
echo.
echo 成功的話，請打開 exports，按 F5 重新整理。
echo 檔名像是 20261112賽後與1113賽前.txt
echo.
pause
