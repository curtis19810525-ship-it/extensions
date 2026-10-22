@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo NBA 回家日常（只串現有爬蟲）
echo ============================================================
echo.
echo 請放在「抓取NBA網路資料」，與 run_all.py 同一層。
echo 請先關閉 NBA26-27爬蟲數據.xlsx。
echo.
echo 賽果日跑步驟 1～3，關注日跑步驟 4～6。
echo 不歸檔、不填盤口、不匯出。
echo 手動選單仍用 爬蟲正負各隊.bat。
echo.

python "%~dp0nba_home.py"
echo.
pause
