# 各隊 total 與球員狀態

給本機「抓取NBA網路資料」用的 ESPN 腳本。寫入 `NBA26-27爬蟲數據.xlsx` 的「各隊total」「球員狀態」，傷兵頁成功時一併更新「球員傷兵」。

## 空數字的原因

ESPN 對 Chrome User-Agent 常回 **HTTP 202、內容是空的**。舊程式仍寫入隊名和 `Total`，後面 26 欄留空，合理讓分就變成 0。

這次改兩件事：

1. 改用拿得到表格的 User-Agent，202 或空頁就換下一個，不把空頁當成績。
2. 網址鎖定 `season/2026` + `seasontype/2`（**2025-26 例行賽**）。頁面賽季不符，或任一隊沒有 GP／PTS，**整次不覆寫 Excel**。

資料夾名 `26-27` 不是 ESPN 賽季代碼。目前 ESPN 球隊數據仍是 2025-26，還沒有 2026-27。

## 接到本機

先關閉 OneDrive 上的 `NBA26-27爬蟲數據.xlsx`，再把下面檔案蓋進「抓取NBA網路資料」：

- `espn_common.py`
- `爬取ESPN各隊Total.py`
- `爬取ESPN球員狀態.py`
- `espn_team_total.py`
- `espn_player_stats.py`

原本的 bat 照舊呼叫 `espn_team_total.py`、`espn_player_stats.py`。

路徑讀同資料夾 `config.py` 的 `NBA_STATS_FILE`。沒有 config 時預設：

`C:\Users\curti\OneDrive\NBA26-27\NBA26-27爬蟲數據.xlsx`

路徑若指向 MLB 活頁簿，程式直接結束，不寫入。

隊名沿用各隊 total 名稱切換後的寫法，賽爾提克（不是塞爾提克），30 隊順序與「各隊total」相同。

## 先試抓、不寫檔

```bat
python 爬取ESPN各隊Total.py --dry-run --only 賽爾提克
python 爬取ESPN球員狀態.py --dry-run --stats-only --only 賽爾提克
```

確認有 GP、PTS 後，再跑原本的 bat。寫入時必須抓滿 30 隊。
