# -*- coding: utf-8 -*-
"""把 NotebookLM 產出的一天，寫進本機的量化日記。

雙擊「寫入量化日記.bat」。Excel 要先關掉。
預設讀同一個資料夾的「今日日記.json」。
日記檔會先找「NBA26-27_Quant_Betting_Journal (1).xlsx」。
氣候只用：氣候A (放養紅利期)、氣候B (設局誘盤期)、氣候C (大收割日)、氣候D (冷卻重置期)。
"""

from __future__ import annotations

import json
import os
import re
import shutil
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

WORKBOOK_NAMES = (
    "NBA26-27_Quant_Betting_Journal (1).xlsx",
    "NBA26-27_Quant_Betting_Journal.xlsx",
)
JSON_NAME = "今日日記.json"
PATH_FILE_NAME = "日記路徑.txt"
MASTER_LAST_ROW = 400

NAVY = "1F4E78"
ICE = "DCE6F1"
ICE_ALT = "EDF3F9"
GREEN = "E2EFDA"
GREEN_ALT = "F4FBF2"
WHITE = "FFFFFF"
LINE = "B0B7C3"
FONT_NAME = "Calibri"

OFFICIAL_CLIMATES = (
    "氣候A (放養紅利期)",
    "氣候B (設局誘盤期)",
    "氣候C (大收割日)",
    "氣候D (冷卻重置期)",
)
OFFICIAL_SIGNALS = (
    "RLM逆向推盤",
    "低溫盲區受讓",
    "KeyPoint鎖水",
    "高溫熔斷受讓",
    "風控鎖倉",
)
CLIMATES = {"".join(name.casefold().split()): name for name in OFFICIAL_CLIMATES}
SIGNALS = {
    "rlm逆向推盤": "RLM逆向推盤",
    "rlm": "RLM逆向推盤",
    "低溫盲區受讓": "低溫盲區受讓",
    "低溫盲區": "低溫盲區受讓",
    "低溫受讓": "低溫盲區受讓",
    "高溫送水盲區": "低溫盲區受讓",
    "盲區": "低溫盲區受讓",
    "keypoint鎖水": "KeyPoint鎖水",
    "keypoint": "KeyPoint鎖水",
    "高溫熔斷受讓": "高溫熔斷受讓",
    "高溫熔斷": "高溫熔斷受讓",
    "極端高溫熔斷": "高溫熔斷受讓",
    "風控鎖倉": "風控鎖倉",
    "風控": "風控鎖倉",
}
OUTCOMES = {
    "全贏": "全贏",
    "贏": "全贏",
    "全輸": "全輸",
    "輸": "全輸",
    "平手": "平手",
    "走盤": "平手",
    "走水": "平手",
    "推": "平手",
    "鎖倉避險": "鎖倉避險",
    "鎖倉": "鎖倉避險",
    "不進場": "鎖倉避險",
}

THIN = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
FILL_NAVY = PatternFill("solid", fgColor=NAVY)
FILL_ICE = PatternFill("solid", fgColor=ICE)
FILL_ICE_ALT = PatternFill("solid", fgColor=ICE_ALT)
FILL_GREEN = PatternFill("solid", fgColor=GREEN)
FILL_GREEN_ALT = PatternFill("solid", fgColor=GREEN_ALT)
FILL_WHITE = PatternFill("solid", fgColor=WHITE)
FILL_SECTION = PatternFill("solid", fgColor="2F5496")
FONT_BODY = Font(name=FONT_NAME, size=11, color="1A1A1A")
FONT_SECTION = Font(name=FONT_NAME, size=11, bold=True, color=WHITE)
FONT_HEAD = Font(name=FONT_NAME, size=12, bold=True, color=WHITE)
FONT_LINK = Font(name=FONT_NAME, size=11, bold=True, color=NAVY, underline="single")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)


class JournalError(Exception):
    pass


def die(message: str, code: int = 1) -> None:
    print(message)
    raise SystemExit(code)


def script_dir() -> Path:
    return Path(__file__).resolve().parent


def search_roots() -> list[Path]:
    home = Path.home()
    roots = [
        script_dir(),
        Path.cwd(),
        home / "Downloads",
        home / "下載",
        home / "OneDrive" / "NBA26-27",
        home / "OneDrive" / "Downloads",
        home / "Desktop",
        home / "OneDrive" / "桌面",
        home / "OneDrive" / "Desktop",
    ]
    unique: list[Path] = []
    seen = set()
    for root in roots:
        try:
            resolved = root.resolve()
        except OSError:
            continue
        if resolved in seen:
            continue
        seen.add(resolved)
        unique.append(root)
    return unique


def reject_foreign_workbook(path: Path) -> None:
    name = path.name.casefold()
    if "mlb" in name or "棒球" in path.name:
        raise JournalError(f"這個路徑像 MLB 檔，不寫入：{path}")
    if "爬蟲" in path.name:
        raise JournalError(f"這是爬蟲數據，不是量化日記，不寫入：{path}")


def find_workbooks(roots: list[Path] | None = None) -> list[Path]:
    found: list[Path] = []
    seen = set()
    for name in WORKBOOK_NAMES:
        for root in roots if roots is not None else search_roots():
            path = root / name
            if not path.is_file():
                continue
            try:
                resolved = path.resolve()
            except OSError:
                continue
            if resolved in seen:
                continue
            seen.add(resolved)
            found.append(path)
    return found


def choose_workbook(
    explicit: Path | None = None,
    roots: list[Path] | None = None,
    override_file: Path | None = None,
) -> tuple[Path, list[Path]]:
    if explicit is not None:
        if not explicit.is_file():
            raise JournalError(f"找不到日記檔：{explicit}")
        reject_foreign_workbook(explicit)
        return explicit, []
    if override_file is None:
        override_file = script_dir() / PATH_FILE_NAME
    override = None
    if override_file.is_file():
        folder = override_file.parent
        for line in override_file.read_text(encoding="utf-8-sig").splitlines():
            text = line.strip()
            if not text or text.startswith("#"):
                continue
            override = Path(text)
            if not override.is_absolute():
                override = folder / override
            break
    if override is not None:
        if not override.is_file():
            raise JournalError(f"日記路徑.txt 指到的檔不存在：{override}")
        reject_foreign_workbook(override)
        return override, []
    found = find_workbooks(roots)
    if not found:
        places = "\n".join(f"  {root}" for root in (roots or search_roots()))
        raise JournalError(
            "找不到量化日記。請把 NBA26-27_Quant_Betting_Journal (1).xlsx 放到這支 bat 旁邊，\n"
            "或把完整路徑寫進同一個資料夾的 日記路徑.txt。\n"
            f"有找過：\n{places}"
        )
    reject_foreign_workbook(found[0])
    return found[0], found[1:]


def choose_json(explicit: Path | None = None) -> Path:
    if explicit is not None:
        if not explicit.is_file():
            raise JournalError(f"找不到 JSON：{explicit}")
        return explicit
    path = script_dir() / JSON_NAME
    if not path.is_file():
        raise JournalError(
            f"找不到 {JSON_NAME}。請把 NotebookLM 產出的一天存成這個檔，放在 bat 旁邊。"
        )
    return path


def load_package(path: Path) -> dict:
    raw = path.read_text(encoding="utf-8-sig").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*", "", raw, flags=re.IGNORECASE)
        raw = re.sub(r"\s*```$", "", raw)
    start = raw.find("{")
    end = raw.rfind("}")
    if start < 0 or end <= start:
        raise JournalError(f"{path.name} 裡沒有 JSON。請貼一天的資料，最外層是大括號。")
    try:
        data = json.loads(raw[start:end + 1])
    except json.JSONDecodeError as exc:
        raise JournalError(
            f"{path.name} 不是合法的 JSON。請檢查引號和逗號。第 {exc.lineno} 行附近：{exc.msg}"
        ) from exc
    if not isinstance(data, dict):
        raise JournalError("請用一個日期的 JSON。不要用陣列一次放很多天。")
    if data.get("_範例") is True:
        raise JournalError(
            "這份還是範例。請把 NotebookLM 的日期、比賽和文字貼進 今日日記.json，再刪掉 _範例 那一行。"
        )
    return data


def pick(data: dict, *names: str):
    for name in names:
        if name in data and data[name] not in (None, ""):
            return data[name]
    return None


def squash(value) -> str:
    text = str(value or "").strip()
    text = text.replace("（", "(").replace("）", ")").replace("　", "")
    return re.sub(r"\s+", "", text)


def parse_date(value) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    matched = re.fullmatch(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})\D*", text)
    if matched:
        year, month, day = (int(part) for part in matched.groups())
        try:
            return date(year, month, day)
        except ValueError as exc:
            raise JournalError(f"日期不正確：{value}") from exc
    raise JournalError(f"日期請寫成 2026-11-04。現在讀到：{value}")


def canonical_climate(value) -> str:
    key = squash(value).casefold()
    if key not in CLIMATES:
        names = "、".join(OFFICIAL_CLIMATES)
        raise JournalError(f"氣候要是這四個之一：{names}。現在讀到：{value}")
    return CLIMATES[key]


def canonical_signal(value) -> str:
    text = str(value or "").strip()
    key = squash(text).casefold()
    if key not in SIGNALS:
        names = "、".join(OFFICIAL_SIGNALS)
        raise JournalError(f"訊號類型要是：{names}。現在讀到：{value}")
    return SIGNALS[key]


def canonical_outcome(value) -> str:
    key = str(value or "").strip()
    if key not in OUTCOMES:
        raise JournalError(f"盤口結果要是全贏、全輸、平手、鎖倉避險。現在讀到：{value}")
    return OUTCOMES[key]


def parse_stake(value) -> float:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)
    else:
        text = str(value or "").strip()
        if text in ("不進場", "鎖倉", "鎖倉避險"):
            return 0.0
        text = text.replace("Ｕ", "U").replace("ｕ", "U")
        text = re.sub(r"[Uu]$", "", text).strip()
        try:
            number = float(text)
        except ValueError as exc:
            raise JournalError(f"注碼要是數字，例如 0.5 或 1。現在讀到：{value}") from exc
    if number < 0:
        raise JournalError(f"注碼不能是負數：{value}")
    return number


def parse_cf(value):
    if value is None or str(value).strip() in ("", "—", "-"):
        return None
    raw = str(value).strip()
    percent = raw.endswith("%") or raw.endswith("％")
    text = raw.replace("%", "").replace("％", "").strip()
    try:
        number = float(text)
    except ValueError as exc:
        raise JournalError(f"尾盤金流要是數字，例如 0.55 或 55%。現在讀到：{value}") from exc
    if percent or number > 1:
        number = number / 100
    return number


def parse_margin(value):
    if value is None or str(value).strip() in ("", "—", "-"):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip().replace("＋", "+").replace("－", "-")
    try:
        return float(text)
    except ValueError as exc:
        raise JournalError(f"讓分差要是數字。推薦的那一隊贏寫正數，輸寫負數。現在讀到：{value}") from exc


def parse_line(value, stake: float) -> str:
    if value is None or str(value).strip() in ("", "—", "-"):
        return "—" if stake == 0 else ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f"{float(value):.3f}"
    return str(value).strip()


def cell_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def pl_formula(row: int) -> str:
    odds = f'IFERROR(VALUE(TRIM(RIGHT(SUBSTITUTE(F{row}," ",REPT(" ",99)),99))),0)'
    return (
        f'=IF(A{row}="","",'
        f'IF(L{row}="全贏",N(I{row})*{odds},'
        f'IF(L{row}="全輸",-N(I{row}),0)))'
    )


def normalize_package(data: dict) -> dict:
    day = parse_date(pick(data, "date", "日期"))
    games_raw = data.get("games", data.get("比賽", []))
    if games_raw is None:
        games_raw = []
    if not isinstance(games_raw, list):
        raise JournalError("games 要是陣列。每一場一個大括號。")
    day_climate = pick(data, "climate", "氣候", "大盤氣候")
    games = []
    for index, item in enumerate(games_raw, start=1):
        if not isinstance(item, dict):
            raise JournalError(f"第 {index} 場不是 JSON 物件。")
        climate_value = pick(item, "climate", "氣候", "大盤氣候") or day_climate
        if climate_value is None:
            raise JournalError(f"第 {index} 場沒有氣候。")
        stake_value = pick(item, "stake_u", "注碼", "建議注碼U", "建議注碼")
        if stake_value is None:
            raise JournalError(f"第 {index} 場沒有注碼。鎖倉請寫 0。")
        stake = parse_stake(stake_value)
        if stake not in (0.0, 0.5, 1.0):
            raise JournalError(f"第 {index} 場注碼只能是 0、0.5、1。現在讀到：{stake_value}")
        matchup = pick(item, "matchup", "對戰組合", "對戰")
        if not matchup:
            raise JournalError(f"第 {index} 場沒有對戰組合。")
        game_id = str(pick(item, "game_id", "場次", "id") or f"G{index}").strip()
        raw_signal = pick(item, "signal_type", "訊號類型", "訊號")
        raw_outcome = pick(item, "outcome", "盤口結果", "結果")
        if stake == 0 and not raw_signal:
            raw_signal = "風控鎖倉"
        if stake == 0 and not raw_outcome:
            raw_outcome = "鎖倉避險"
        signal = canonical_signal(raw_signal)
        outcome = canonical_outcome(raw_outcome)
        pick_text = str(pick(item, "pick", "推薦標的", "推薦") or "").strip()
        line_odds = parse_line(pick(item, "line_odds", "盤口", "盤口規格與水位", "水位"), stake)
        if stake == 0 or outcome == "鎖倉避險" or signal == "風控鎖倉":
            stake = 0.0
            outcome = "鎖倉避險"
            signal = "風控鎖倉"
            line_odds = "—"
            pick_text = "【$0 風控鎖倉】"
        climate = canonical_climate(climate_value)
        games.append({
            "game_id": game_id,
            "matchup": str(matchup).strip(),
            "climate": climate,
            "pick": pick_text,
            "line_odds": line_odds,
            "cf_pct": parse_cf(pick(item, "cf_pct", "金流", "尾盤金流", "cf")),
            "signal_type": signal,
            "stake_u": stake,
            "score": "" if pick(item, "score", "終場比分", "比分") is None else str(pick(item, "score", "終場比分", "比分")).strip(),
            "margin": parse_margin(pick(item, "margin", "讓分差")),
            "outcome": outcome,
            "review_notes": str(pick(item, "review_notes", "檢討", "賽後邏輯", "備註") or "").strip(),
            "signal_alias": squash(str(raw_signal or "")).casefold() not in {squash(name).casefold() for name in OFFICIAL_SIGNALS},
        })
    if day_climate is None and games:
        day_climate = games[0]["climate"]
    if day_climate is None:
        raise JournalError("這一天沒有氣候。")
    climate = canonical_climate(day_climate)
    for game in games:
        if game["climate"] != climate:
            raise JournalError(f"{game['game_id']} 的氣候是 {game['climate']}，當天是 {climate}。請改成同一天同一個氣候。")
    pregame = pick(data, "pre_game_text", "賽前", "pregame")
    postgame = pick(data, "post_game_text", "賽後", "postgame")
    if not pregame or not postgame:
        raise JournalError("請同時給賽前文字和賽後文字。")
    warnings = []
    for game in games:
        if game.pop("signal_alias"):
            warnings.append(f"{game['game_id']} 的訊號已改寫成 {game['signal_type']}。")
        if game["outcome"] == "全贏" and not game["line_odds"]:
            warnings.append(f"{game['game_id']} 是全贏，但水位是空的，損益會變成 0。")
    return {
        "date": day,
        "climate": climate,
        "pre_game_text": str(pregame).strip(),
        "post_game_text": str(postgame).strip(),
        "games": games,
        "warnings": warnings,
    }


def sheet_or_stop(wb, name: str):
    if name not in wb.sheetnames:
        raise JournalError(f"這本日記沒有分頁 {name}。請確認是量化日記，不是別的 Excel。")
    return wb[name]


def last_value_row(ws, column: int, start: int) -> int:
    last = start - 1
    for row in range(start, ws.max_row + 1):
        if ws.cell(row, column).value not in (None, ""):
            last = row
    return last


def date_exists(ws, column: int, start: int, day: date) -> bool:
    for row in range(start, ws.max_row + 1):
        if cell_date(ws.cell(row, column).value) == day:
            return True
    return False


def ensure_closed(path: Path) -> None:
    try:
        handle = open(path, "r+b")
    except PermissionError as exc:
        raise JournalError(
            "這本 Excel 正在打開，沒有寫入。請先關閉日記檔，再按一次 bat。"
        ) from exc
    handle.close()


def put(cell, value, font, fill, align, fmt=None) -> None:
    cell.value = value
    cell.font = font
    cell.fill = fill
    cell.alignment = align
    cell.border = THIN
    if fmt:
        cell.number_format = fmt


def text_height(text: str) -> int:
    lines = text.count("\n") + 1
    return min(320, max(96, lines * 18 + 12))


def extend_filter(ws, last_row: int) -> None:
    ref = ws.auto_filter.ref or "A2:N2"
    match = re.search(r"(\d+)$", ref.split(":")[-1])
    current = int(match.group(1)) if match else last_row
    ws.auto_filter.ref = f"A2:N{max(current, last_row)}"


def extend_chart(ws, last_row: int) -> None:
    if not ws._charts:
        return
    chart = ws._charts[0]
    series = chart.series[0]
    title = ws.title.replace("'", "''")
    if series.val is not None and series.val.numRef is not None:
        series.val.numRef.f = f"'{title}'!$G$3:$G${last_row}"
    if series.cat is not None and series.cat.numRef is not None:
        series.cat.numRef.f = f"'{title}'!$A$3:$A${last_row}"
    anchor = chart.anchor
    marker = getattr(anchor, "_from", None)
    if marker is not None:
        marker.row = last_row + 1
        marker.col = 0


def append_master(ws, package: dict) -> list[int]:
    last = last_value_row(ws, 1, 3)
    start = last + 1
    games = package["games"]
    if start + len(games) - 1 > MASTER_LAST_ROW:
        raise JournalError(f"流水帳只預留到第 {MASTER_LAST_ROW} 列，這一天放不下。")
    rows = []
    for offset, game in enumerate(games):
        row = start + offset
        alt = (row - 3) % 2 == 1
        ice = FILL_ICE_ALT if alt else FILL_ICE
        green = FILL_GREEN_ALT if alt else FILL_GREEN
        values = [
            package["date"],
            game["game_id"],
            game["matchup"],
            game["climate"],
            game["pick"],
            game["line_odds"],
            game["cf_pct"],
            game["signal_type"],
            game["stake_u"],
            game["score"],
            game["margin"],
            game["outcome"],
            pl_formula(row),
            game["review_notes"],
        ]
        for column, value in enumerate(values, start=1):
            fill = green if column == 13 else ice
            align = LEFT if column in (3, 5, 6, 14) else CENTER
            font = Font(name=FONT_NAME, size=11, bold=True) if column == 13 else FONT_BODY
            put(ws.cell(row, column), value, font, fill, align)
        ws.cell(row, 1).number_format = "yyyy-mm-dd"
        ws.cell(row, 7).number_format = "0.0%"
        ws.cell(row, 9).number_format = '0.0"U"'
        ws.cell(row, 11).number_format = "0.0"
        ws.cell(row, 13).number_format = '+0.000"U";-0.000"U";0.000"U"'
        ws.row_dimensions[row].height = 22 if game["stake_u"] == 0 else 30
        rows.append(row)
    if rows:
        extend_filter(ws, rows[-1])
    return rows


def append_equity(ws, package: dict, has_games: bool = True) -> int:
    del has_games
    last = last_value_row(ws, 1, 3)
    row = last + 1
    alt = (row - 3) % 2 == 1
    ice = FILL_ICE_ALT if alt else FILL_ICE
    green = FILL_GREEN_ALT if alt else FILL_GREEN
    put(ws.cell(row, 1), package["date"], FONT_BODY, ice, CENTER, "yyyy-mm-dd")
    put(ws.cell(row, 2), package["climate"], FONT_BODY, green, LEFT)
    put(ws.cell(row, 3), f'=COUNTIFS(\'2_Master_Log\'!A:A,A{row},\'2_Master_Log\'!I:I,">0")', FONT_BODY, green, CENTER, "0")
    put(ws.cell(row, 4), f'=COUNTIFS(\'2_Master_Log\'!A:A,A{row},\'2_Master_Log\'!I:I,0)', FONT_BODY, green, CENTER, "0")
    put(ws.cell(row, 5), f'=SUMIF(\'2_Master_Log\'!A:A,A{row},\'2_Master_Log\'!I:I)', FONT_BODY, green, CENTER, '0.0"U"')
    put(
        ws.cell(row, 6),
        f'=SUMIF(\'2_Master_Log\'!A:A,A{row},\'2_Master_Log\'!M:M)',
        FONT_BODY, green, CENTER, '+0.000"U";-0.000"U";0.000"U"',
    )
    put(
        ws.cell(row, 7),
        f"=SUM(F$3:F{row})",
        Font(name=FONT_NAME, size=11, bold=True, color="1A1A1A"),
        green, CENTER, '+0.000"U";-0.000"U";0.000"U"',
    )
    put(
        ws.cell(row, 8),
        f'=IFERROR(COUNTIFS(\'2_Master_Log\'!A:A,A{row},\'2_Master_Log\'!L:L,"全贏")/C{row},0)',
        FONT_BODY, green, CENTER, "0.0%",
    )
    put(
        ws.cell(row, 9),
        f'=HYPERLINK("#\'4_Daily_Journal\'!A"&MATCH(A{row},\'4_Daily_Journal\'!A:A,0),"📖 查看 "&TEXT(A{row},"yyyy-mm-dd")&" 詳細賽前分析與復盤日誌")',
        FONT_LINK, green, CENTER,
    )
    ws.row_dimensions[row].height = 24
    extend_chart(ws, row)
    return row


def journal_start(ws) -> int:
    last = 2
    for row in range(3, ws.max_row + 1):
        if ws.cell(row, 1).value not in (None, "") or ws.cell(row, 2).value not in (None, ""):
            last = row
    start = last + 2
    while ws.cell(start, 1).value not in (None, "") or ws.cell(start, 2).value not in (None, ""):
        start += 1
        if start > 5000:
            raise JournalError("日誌頁太長，找不到可以接上的空白列。")
    return start


def _merge_text(ws, row: int, value, font, fill, align, height: int) -> None:
    put(ws.cell(row, 2), value, font, fill, align)
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
    ws.row_dimensions[row].height = height


def append_journal(ws, package: dict) -> int:
    row = journal_start(ws)
    day = package["date"]
    banner = f"{day.isoformat()} | {package['climate']}"
    blocks = [
        (banner, FONT_HEAD, FILL_NAVY, Alignment(horizontal="left", vertical="center", wrap_text=True), 28),
        ("【賽前量化分析與對話推論】", FONT_SECTION, FILL_SECTION, Alignment(horizontal="left", vertical="center"), 22),
        (package["pre_game_text"], FONT_BODY, FILL_ICE, TOP, text_height(package["pre_game_text"])),
        ("【賽後對戰結算與復盤日誌】", FONT_SECTION, FILL_SECTION, Alignment(horizontal="left", vertical="center"), 22),
        (package["post_game_text"], FONT_BODY, FILL_GREEN, TOP, text_height(package["post_game_text"])),
    ]
    for offset, (value, font, fill, align, height) in enumerate(blocks):
        put(ws.cell(row + offset, 1), day, FONT_HEAD if offset == 0 else FONT_BODY, fill, CENTER, "yyyy-mm-dd")
        _merge_text(ws, row + offset, value, font, fill, align, height)
    link_row = row + 5
    put(ws.cell(link_row, 1), day, FONT_BODY, FILL_WHITE, CENTER, "yyyy-mm-dd")
    put(
        ws.cell(link_row, 2),
        '=HYPERLINK("#\'3_Equity_Timeline\'!A1","⬆️ 返回時間軸總覽")',
        FONT_LINK, FILL_WHITE,
        Alignment(horizontal="left", vertical="center"),
    )
    ws.row_dimensions[link_row].height = 22
    ws.row_dimensions[link_row + 1].height = 16
    return row


def locate_rows(wb, package: dict) -> dict:
    master = sheet_or_stop(wb, "2_Master_Log")
    equity = sheet_or_stop(wb, "3_Equity_Timeline")
    journal = sheet_or_stop(wb, "4_Daily_Journal")
    sheet_or_stop(wb, "1_Dashboard")
    day = package["date"]
    if date_exists(master, 1, 3, day) or date_exists(equity, 1, 3, day) or date_exists(journal, 1, 3, day):
        raise JournalError(f"{day.isoformat()} 已經在日記裡。這支不會把同一天再寫一次。")
    master_last = last_value_row(master, 1, 3)
    equity_last = last_value_row(equity, 1, 3)
    game_count = len(package["games"])
    master_rows = list(range(master_last + 1, master_last + 1 + game_count))
    return {
        "master_rows": master_rows,
        "equity_row": equity_last + 1,
        "journal_row": journal_start(journal),
    }


def apply_package(wb, package: dict) -> dict:
    plan = locate_rows(wb, package)
    master_rows = append_master(sheet_or_stop(wb, "2_Master_Log"), package)
    equity_row = append_equity(sheet_or_stop(wb, "3_Equity_Timeline"), package, bool(package["games"]))
    journal_row = append_journal(sheet_or_stop(wb, "4_Daily_Journal"), package)
    if master_rows != plan["master_rows"] or equity_row != plan["equity_row"] or journal_row != plan["journal_row"]:
        raise JournalError("寫入位置和檢查時不一樣，已停止，沒有存檔。")
    return plan


def describe(workbook: Path, json_path: Path, package: dict, plan: dict, others: list[Path]) -> str:
    games = package["games"]
    if plan["master_rows"]:
        first = plan["master_rows"][0]
        last = plan["master_rows"][-1]
        master_text = f"第 {first} 列到第 {last} 列，共 {len(games)} 場"
    else:
        master_text = "這一天沒有比賽列"
    lines = [
        f"日記檔：{workbook}",
        f"JSON：{json_path}",
        f"日期：{package['date'].isoformat()}",
        f"氣候：{package['climate']}",
        f"流水帳：{master_text}",
        f"時間軸：第 {plan['equity_row']} 列",
        f"日誌：第 {plan['journal_row']} 列",
    ]
    for other in others:
        lines.append(f"另外也找到這本，這次不會寫它：{other}")
    for warning in package.get("warnings", []):
        lines.append(f"請注意：{warning}")
    return "\n".join(lines)


def save_workbook(wb, path: Path) -> Path:
    backup_dir = path.parent / "還原點"
    backup_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = backup_dir / f"{path.stem}_{stamp}.xlsx"
    shutil.copy2(path, backup)
    temporary = path.with_name(f".{path.stem}.writing.xlsx")
    try:
        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True
        wb.save(temporary)
        os.replace(temporary, path)
    except PermissionError as exc:
        raise JournalError("存檔時 Excel 把日記檔鎖住了。請先關閉日記檔，再按一次。原本的檔沒有改。") from exc
    finally:
        if temporary.exists():
            temporary.unlink()
    return backup


def other_workbooks(selected: Path, roots: list[Path] | None = None) -> list[Path]:
    chosen = selected.resolve()
    others = []
    for path in find_workbooks(roots):
        try:
            if path.resolve() == chosen:
                continue
        except OSError:
            continue
        others.append(path)
    return others


def run(json_path: Path, workbook: Path, dry_run: bool = False, roots: list[Path] | None = None) -> str:
    package = normalize_package(load_package(json_path))
    selected, _ignored = choose_workbook(workbook, roots)
    others = other_workbooks(selected, roots)
    ensure_closed(selected)
    wb = load_workbook(selected)
    plan = locate_rows(wb, package) if dry_run else apply_package(wb, package)
    text = describe(selected, json_path, package, plan, others)
    if dry_run:
        text += "\n上面只是檢查，還沒寫入。"
        text += "\n確認日期和日記檔路徑後，回到這個視窗按一個鍵就會寫入。要取消就關掉視窗。"
        print(text)
        return text
    backup = save_workbook(wb, selected)
    text += f"\n已寫入。還原點：{backup}"
    text += "\n請重新打開這本 Excel。第 1 頁的數字會自己重算。"
    print(text)
    return text


def main(argv: list[str] | None = None) -> None:
    args = list(sys.argv[1:] if argv is None else argv)
    dry_run = False
    json_arg = None
    workbook_arg = None
    for arg in args:
        if arg == "--dry-run":
            dry_run = True
            continue
        path = Path(arg)
        suffix = path.suffix.casefold()
        if suffix == ".json":
            json_arg = path
        elif suffix == ".xlsx":
            workbook_arg = path
        else:
            die(f"看不懂這個參數：{arg}")
    try:
        json_path = choose_json(json_arg)
        workbook, _others = choose_workbook(workbook_arg)
        run(json_path, workbook, dry_run=dry_run)
    except JournalError as exc:
        die(str(exc))


if __name__ == "__main__":
    main()
