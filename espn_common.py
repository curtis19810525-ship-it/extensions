# -*- coding: utf-8 -*-
"""ESPN NBA 共用：鎖定 2025-26 例行賽，並避開會回 HTTP 202 空頁的 User-Agent。"""

import os
import re

import requests

# ESPN：season 2026 = 2025-26；seasontype 2 = Regular Season。
# 資料夾名 26-27 不是 ESPN 賽季代碼。目前站上還沒有 2026-27 球隊數據。
SEASON_YEAR = 2026
SEASON_TYPE = 2
SEASON_LABEL = "2025-26"
SEASON_TYPE_NAME = "Regular Season"

# Chrome 這組在 ESPN 常得到 HTTP 202 且 body 為空；空頁仍寫入隊名就是「各隊total 只有隊名」。
USER_AGENTS = [
    "Mozilla/5.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:126.0) Gecko/20100101 Firefox/126.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
]

# 與 2026-03-31「各隊total 名稱切換」同一套隊名與順序（賽爾提克，不是塞爾提克）。
NBA_TEAMS = [
    ("賽爾提克", "bos", "boston-celtics"),
    ("籃網", "bkn", "brooklyn-nets"),
    ("尼克", "ny", "new-york-knicks"),
    ("76人", "phi", "philadelphia-76ers"),
    ("暴龍", "tor", "toronto-raptors"),
    ("公牛", "chi", "chicago-bulls"),
    ("騎士", "cle", "cleveland-cavaliers"),
    ("活塞", "det", "detroit-pistons"),
    ("溜馬", "ind", "indiana-pacers"),
    ("公鹿", "mil", "milwaukee-bucks"),
    ("老鷹", "atl", "atlanta-hawks"),
    ("黃蜂", "cha", "charlotte-hornets"),
    ("熱火", "mia", "miami-heat"),
    ("魔術", "orl", "orlando-magic"),
    ("巫師", "was", "washington-wizards"),
    ("勇士", "gs", "golden-state-warriors"),
    ("快艇", "lac", "los-angeles-clippers"),
    ("湖人", "lal", "los-angeles-lakers"),
    ("太陽", "phx", "phoenix-suns"),
    ("國王", "sac", "sacramento-kings"),
    ("金塊", "den", "denver-nuggets"),
    ("灰狼", "min", "minnesota-timberwolves"),
    ("雷霆", "okc", "oklahoma-city-thunder"),
    ("拓荒者", "por", "portland-trail-blazers"),
    ("爵士", "utah", "utah-jazz"),
    ("獨行俠", "dal", "dallas-mavericks"),
    ("火箭", "hou", "houston-rockets"),
    ("灰熊", "mem", "memphis-grizzlies"),
    ("鵜鶘", "no", "new-orleans-pelicans"),
    ("馬刺", "sa", "san-antonio-spurs"),
]

DEFAULT_EXCEL = r"C:\Users\curti\OneDrive\NBA26-27\NBA26-27爬蟲數據.xlsx"


def excel_path():
    """同資料夾 config.NBA_STATS_FILE 優先；明確拒絕寫進 MLB 活頁簿。"""
    env = os.environ.get("NBA_STATS_FILE")
    if env:
        path = env
    else:
        try:
            from config import NBA_STATS_FILE as path
        except ImportError:
            path = DEFAULT_EXCEL
    upper = str(path).upper().replace("/", "\\")
    if "\\MLB" in upper or "MLB" in os.path.basename(str(path)).upper():
        raise SystemExit(
            "拒絕寫入：路徑指向 MLB 活頁簿（%s）。"
            "請在「抓取NBA網路資料」的 config.py 或環境變數 NBA_STATS_FILE "
            "指定 NBA26-27爬蟲數據.xlsx。" % path
        )
    return path


def stats_url(abbr, slug):
    return (
        "https://www.espn.com/nba/team/stats/_/name/%s/%s/season/%s/seasontype/%s"
        % (abbr, slug, SEASON_YEAR, SEASON_TYPE)
    )


def injuries_url(abbr, slug):
    return "https://www.espn.com/nba/team/injuries/_/name/%s/%s" % (abbr, slug)


def season_error(html):
    """頁面不是鎖定的 2025-26 例行賽時回傳原因，否則 None。"""
    if not html or not str(html).strip():
        return "空頁"
    match = re.search(
        r'"requestedSeason"\s*:\s*\{[^{}]*"year"\s*:\s*(\d+)'
        r'[^{}]*"type"\s*:\s*(\d+)'
        r'[^{}]*"displayName"\s*:\s*"([^"]+)"',
        html,
    )
    if match:
        year = int(match.group(1))
        season_type = int(match.group(2))
        display = match.group(3)
        if year != SEASON_YEAR or season_type != SEASON_TYPE or display != SEASON_LABEL:
            return "頁面是 %s（season=%s type=%s），不是 %s type=%s" % (
                display, year, season_type, SEASON_LABEL, SEASON_TYPE,
            )
        return None
    title_match = re.search(r"<title[^>]*>(.*?)</title>", html, re.I | re.S)
    title = re.sub(r"\s+", " ", title_match.group(1)).strip() if title_match else ""
    if SEASON_LABEL in title and SEASON_TYPE_NAME in title:
        return None
    return "無法確認頁面是 %s %s（title=%s）" % (SEASON_LABEL, SEASON_TYPE_NAME, title or "無")


def fetch_html(url, check_season=False):
    """換 User-Agent 重試。202 或空 body 視為失敗，不把空頁交給解析。"""
    last_error = None
    for ua in USER_AGENTS:
        headers = {"User-Agent": ua}
        for _ in range(2):
            try:
                response = requests.get(url, headers=headers, timeout=30)
            except requests.RequestException as exc:
                last_error = exc
                continue
            body = response.text or ""
            if response.status_code == 202 or not body.strip():
                last_error = RuntimeError("HTTP %s / 空內容" % response.status_code)
                continue
            if response.status_code >= 400:
                last_error = RuntimeError("HTTP %s" % response.status_code)
                continue
            if check_season:
                reason = season_error(body)
                if reason:
                    last_error = RuntimeError(reason)
                    continue
            return body
    raise RuntimeError(str(last_error))


def cell_value(val):
    """統計格轉成數字。ESPN 在 TO=0 時顯示 INF，改寫 0。隊名不要用這個。"""
    if val is None:
        return ""
    text = str(val).strip()
    if text == "":
        return ""
    if text.upper() in ("INF", "-INF", "INFINITY", "-INFINITY") or text == "∞":
        return 0
    numeric = text.replace(",", "")
    try:
        if "." in numeric:
            return float(numeric)
        return int(numeric)
    except ValueError:
        return text
