# -*- coding: utf-8 -*-
"""賽前預測：下拉選讓分邊，賽後在同一列寫出贏輸與顏色。"""

from __future__ import annotations

import re

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

PICK_LABEL = "賽前預測"
NOTE_LABELS = ("賽前想法", "比賽結果", "賽後想法")
SUCCESS = "預測成功"
FAIL = "預測失敗"
PUSH = "預測走盤"

LABEL_FONT = Font(bold=True, size=14, color="1F4E79")
VALUE_FONT = Font(size=16, color="000000")
LABEL_FILL = PatternFill(fill_type="solid", fgColor="DDEBF7")
SUCCESS_FILL = PatternFill(fill_type="solid", fgColor="C6EFCE")
SUCCESS_FONT = Font(bold=True, size=16, color="006100")
FAIL_FILL = PatternFill(fill_type="solid", fgColor="FFC7CE")
FAIL_FONT = Font(bold=True, size=16, color="9C0006")
PUSH_FILL = PatternFill(fill_type="solid", fgColor="D9D9D9")
PUSH_FONT = Font(bold=True, size=16, color="595959")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
STATUS_STYLE = {
    SUCCESS: (SUCCESS_FILL, SUCCESS_FONT),
    FAIL: (FAIL_FILL, FAIL_FONT),
    PUSH: (PUSH_FILL, PUSH_FONT),
}
PICK_RE = re.compile(
    r"^(?P<team>.+?)(?P<side>受讓|讓)(?P<pts>\d+(?:\.\d+)?)分(?P<mark>贏|輸|走)?$"
)


def format_points(value: float) -> str:
    number = abs(float(value))
    if abs(number - round(number)) < 1e-6:
        return str(int(round(number)))
    return f"{number:.2f}".rstrip("0").rstrip(".")


def parse_spread_token(token):
    """從儲存格文字取出讓分。正數代表主隊讓分。(7+50) 視為 7.5。"""
    if isinstance(token, bool) or token is None:
        return None
    if isinstance(token, (int, float)):
        return float(token)
    text = str(token).strip()
    if not text:
        return None
    matched = re.search(r"\(([+-]?\d+(?:\.\d+)?)(\+50|-50)?\)", text)
    if matched:
        base = float(matched.group(1))
        if matched.group(2) == "+50":
            base += 0.5 if base >= 0 else -0.5
        return base
    matched = re.search(r"([+-]?\d+(?:\.\d+)?)\+50", text)
    if matched:
        base = float(matched.group(1))
        return base + (0.5 if base >= 0 else -0.5)
    matched = re.search(r"[+-]?\d+(?:\.\d+)?", text)
    if matched:
        return float(matched.group(0))
    return None


def parse_pick(value):
    if value is None:
        return None
    matched = PICK_RE.match(str(value).strip())
    if not matched:
        return None
    return matched.groupdict()


def base_pick(value) -> str:
    parsed = parse_pick(value)
    if not parsed:
        return ""
    return f"{parsed['team']}{parsed['side']}{parsed['pts']}分"


def strip_role(value) -> str:
    text = str(value or "").strip()
    return text.replace("（主）", "").replace("(主)", "").strip()


def menu_options(home: str, away: str, signed_home):
    if signed_home is None or not home or not away:
        return []
    points = format_points(signed_home)
    if float(signed_home) > 0:
        return [f"{home}讓{points}分", f"{away}受讓{points}分"]
    if float(signed_home) < 0:
        return [f"{home}受讓{points}分", f"{away}讓{points}分"]
    return [f"{home}讓0分", f"{away}受讓0分"]


def judge(margin: float, side: str, points: float) -> str:
    diff = margin - points if side == "讓" else margin + points
    if diff > 0.01:
        return "贏"
    if diff < -0.01:
        return "輸"
    return "走"


def margin_text(team: str, margin: float) -> str:
    if abs(margin) < 0.01:
        return f"{team}平手"
    number = format_points(margin)
    if margin > 0:
        return f"{team}贏{number}分"
    return f"{team}輸{number}分"


def score_after(name: str, text: str):
    if not name:
        return None
    matched = re.search(re.escape(name) + r"\s*(\d{2,3})(?!\d)", text)
    if not matched:
        return None
    return int(matched.group(1))


def resolve_home_margin(result, home: str, away: str):
    if result is None:
        return None
    text = str(result).strip()
    if not text:
        return None
    home_score = score_after(home, text)
    away_score = score_after(away, text)
    if home_score is not None and away_score is not None:
        return float(home_score - away_score)
    matched = re.search(r"客\s*(\d{2,3})\s*[-:：]\s*主\s*(\d{2,3})", text)
    if matched:
        return float(int(matched.group(2)) - int(matched.group(1)))
    matched = re.search(r"主\s*(\d{2,3})\s*[-:：]\s*客\s*(\d{2,3})", text)
    if matched:
        return float(int(matched.group(1)) - int(matched.group(2)))
    matched = re.search(r"主隊贏\s*(\d+(?:\.\d+)?)", text)
    if matched:
        return float(matched.group(1))
    matched = re.search(r"客隊贏\s*(\d+(?:\.\d+)?)", text)
    if matched:
        return -float(matched.group(1))
    if "平手" in text or "打平" in text:
        return 0.0
    return None


def grade_pick(pick, home: str, away: str, home_margin: float):
    parsed = parse_pick(pick)
    if not parsed or home_margin is None:
        return None
    team = parsed["team"]
    if team == home:
        margin = float(home_margin)
    elif team == away:
        margin = -float(home_margin)
    else:
        return None
    mark = judge(margin, parsed["side"], float(parsed["pts"]))
    status = { "贏": SUCCESS, "輸": FAIL, "走": PUSH }[mark]
    return {
        "b": f"{team}{parsed['side']}{parsed['pts']}分{mark}",
        "c": margin_text(team, margin),
        "d": status,
    }


def lookup_pick(game_notes, title: str) -> str:
    if not isinstance(game_notes, dict):
        return ""
    title = str(title or "").strip()
    exact = game_notes.get(title)
    if isinstance(exact, dict) and "賽前預測" in exact:
        return base_pick(exact.get("賽前預測"))
    for key, slot in game_notes.items():
        if not isinstance(slot, dict) or "賽前預測" not in slot:
            continue
        key_text = str(key or "").strip()
        if key_text and (key_text == title or key_text in title or title in key_text):
            return base_pick(slot.get("賽前預測"))
    return ""


def remember_pick(game_notes, title: str, pick: str):
    if not isinstance(game_notes, dict):
        return
    slot = game_notes.get(title)
    if not isinstance(slot, dict):
        slot = {}
        game_notes[title] = slot
    slot["賽前預測"] = pick


def find_header_row(ws, pick_row: int):
    for row in range(pick_row - 1, max(0, pick_row - 40), -1):
        marker = str(ws.cell(row, 1).value or "").strip()
        home = str(ws.cell(row, 2).value or "")
        if marker == "時間" or "（主）" in home or "(主)" in home:
            return row
        if "@" in marker:
            break
    return None


def find_spread_column(ws, header_row: int):
    columns = []
    if header_row:
        columns.append(header_row)
    columns.extend(range(1, 6))
    seen = set()
    for row in columns:
        if row in seen or row < 1:
            continue
        seen.add(row)
        for col in range(15, 81):
            if ws.cell(row, col).value == "讓分A":
                return col
    return None


def signed_from_display(token, home: str, header_d: str):
    number = parse_spread_token(token)
    if number is None:
        return None
    text = str(token)
    if number < 0 or re.search(r"\(\s*-", text):
        return -abs(number)
    header = str(header_d or "")
    if "受讓" in header and (not home or home in header):
        return -abs(number)
    return abs(number)


def read_latest_line(ws, pick_row: int):
    header = find_header_row(ws, pick_row)
    if header is None:
        return "", "", None
    home = strip_role(ws.cell(header, 2).value)
    away = strip_role(ws.cell(header, 3).value)
    spread_col = find_spread_column(ws, header)
    header_d = ws.cell(header, 4).value
    for row in range(pick_row - 1, header, -1):
        marker = str(ws.cell(row, 1).value or "").strip()
        if marker in NOTE_LABELS or marker == PICK_LABEL:
            continue
        if spread_col:
            raw = ws.cell(row, spread_col).value
            if isinstance(raw, (int, float)) and not isinstance(raw, bool):
                return home, away, float(raw)
        signed = signed_from_display(ws.cell(row, 4).value, home, header_d)
        if signed is not None:
            return home, away, signed
    return home, away, None


def result_below(ws, pick_row: int):
    for row in range(pick_row + 1, pick_row + 8):
        if ws.cell(row, 1).value == "比賽結果":
            return ws.cell(row, 2).value
    return None


def _apply_border(cell):
    cell.border = THIN
    cell.alignment = CENTER


def _clear_status(cell):
    cell.value = None
    cell.fill = PatternFill(fill_type=None)
    cell.font = VALUE_FONT
    _apply_border(cell)


def _add_menu(ws, cell, options):
    cleaned = []
    for item in options:
        text = str(item or "").strip()
        if text and text not in cleaned and "," not in text and '"' not in text:
            cleaned.append(text)
    if len(cleaned) < 2:
        return
    check = DataValidation(
        type="list",
        formula1='"' + ",".join(cleaned) + '"',
        allow_blank=True,
    )
    check.showErrorMessage = False
    check.showInputMessage = False
    ws.add_data_validation(check)
    check.add(cell)


def render_pick(ws, title, row, game_notes, last_col=None):
    saved = lookup_pick(game_notes, title)
    home, away, signed = read_latest_line(ws, row)
    options = menu_options(home, away, signed)
    result = result_below(ws, row)
    if not result and isinstance(game_notes, dict):
        slot = game_notes.get(title)
        if isinstance(slot, dict):
            result = slot.get("比賽結果")
    graded = None
    if saved:
        margin = resolve_home_margin(result, home, away)
        graded = grade_pick(saved, home, away, margin) if margin is not None else None

    ws.row_dimensions[row].height = 48
    label = ws.cell(row, 1, PICK_LABEL)
    label.font = LABEL_FONT
    label.fill = LABEL_FILL
    _apply_border(label)

    pick_cell = ws.cell(row, 2, graded["b"] if graded else (saved or None))
    pick_cell.font = VALUE_FONT
    pick_cell.fill = PatternFill(fill_type=None)
    _apply_border(pick_cell)

    margin_cell = ws.cell(row, 3)
    status_cell = ws.cell(row, 4)
    if graded:
        margin_cell.value = graded["c"]
        margin_cell.font = VALUE_FONT
        margin_cell.fill = PatternFill(fill_type=None)
        _apply_border(margin_cell)
        fill, font = STATUS_STYLE[graded["d"]]
        status_cell.value = graded["d"]
        status_cell.fill = fill
        status_cell.font = font
        _apply_border(status_cell)
    else:
        _clear_status(margin_cell)
        _clear_status(status_cell)
        menu = list(options)
        if saved and saved not in menu:
            menu.insert(0, saved)
        _add_menu(ws, pick_cell, menu)

    if saved:
        remember_pick(game_notes, title, saved)
    return row + 1


def _title_above(ws, row: int):
    for current in range(row, 0, -1):
        text = str(ws.cell(current, 1).value or "").strip()
        if "@" in text and not text.startswith("賽"):
            return text
    return ""


def merge_picks(wb, notes):
    if not isinstance(notes, dict) or "時段觀察" not in wb.sheetnames:
        archive_into(notes, wb, seen=set())
        return notes
    ws = wb["時段觀察"]
    seen = set()
    current = ""
    for row in range(1, (ws.max_row or 0) + 1):
        marker = ws.cell(row, 1).value
        text = str(marker or "").strip()
        if "@" in text and not text.startswith("賽"):
            current = text
        if marker == PICK_LABEL and current:
            seen.add(current)
            raw = ws.cell(row, 2).value
            slot = notes.setdefault(current, {})
            if isinstance(slot, dict):
                slot["賽前預測"] = base_pick(raw) if raw else ""
    archive_into(notes, wb, seen)
    return notes


def _archive_layout(ws):
    header_row = None
    title_col = None
    for row in range(1, 8):
        found = False
        for col in range(1, (ws.max_column or 1) + 1):
            value = ws.cell(row, col).value
            if value in NOTE_LABELS:
                header_row = row
                found = True
            if value in ("場次", "比賽", "標題"):
                title_col = col
        if found:
            break
    if header_row and title_col is None:
        title_col = 1
    return header_row, title_col


def _pick_column(ws, header_row: int, create: bool):
    last = ws.max_column or 1
    for col in range(1, last + 1):
        if ws.cell(header_row, col).value == PICK_LABEL:
            return col
    if not create:
        return None
    col = last + 1
    cell = ws.cell(header_row, col, PICK_LABEL)
    cell.font = LABEL_FONT
    return col


def archive_into(notes, wb, seen):
    if not isinstance(notes, dict) or "賽事筆記" not in wb.sheetnames:
        return
    ws = wb["賽事筆記"]
    header_row, title_col = _archive_layout(ws)
    pick_col = _pick_column(ws, header_row, create=False) if header_row else None
    if not header_row or not pick_col:
        return
    for row in range(header_row + 1, (ws.max_row or header_row) + 1):
        title = str(ws.cell(row, title_col).value or "").strip()
        if not title or title in seen:
            continue
        pick = base_pick(ws.cell(row, pick_col).value)
        if not pick:
            continue
        slot = notes.setdefault(title, {})
        if isinstance(slot, dict) and not slot.get("賽前預測"):
            slot["賽前預測"] = pick


def archive_picks(wb, game_notes):
    if not isinstance(game_notes, dict) or "賽事筆記" not in getattr(wb, "sheetnames", []):
        return
    ws = wb["賽事筆記"]
    header_row, title_col = _archive_layout(ws)
    if not header_row:
        return
    pick_col = _pick_column(ws, header_row, create=True)
    rows = {}
    for row in range(header_row + 1, (ws.max_row or header_row) + 1):
        title = str(ws.cell(row, title_col).value or "").strip()
        if title:
            rows[title] = row
    for title, slot in game_notes.items():
        if not isinstance(slot, dict) or "賽前預測" not in slot:
            continue
        pick = base_pick(slot.get("賽前預測"))
        title_text = str(title).strip()
        target = rows.get(title_text)
        if target is None:
            for key, row in rows.items():
                if key and (key in title_text or title_text in key):
                    target = row
                    break
        if target is None:
            if not pick:
                continue
            target = (ws.max_row or header_row) + 1
            ws.cell(target, title_col, title_text)
            rows[title_text] = target
        ws.cell(target, pick_col, pick or None)


def selftest():
    assert format_points(2) == "2"
    assert format_points(14.5) == "14.5"
    assert parse_spread_token("(7+50)") == 7.5
    assert parse_spread_token("(-2)0.9") == -2
    assert base_pick("活塞讓2分贏") == "活塞讓2分"
    assert menu_options("老鷹", "活塞", -2) == ["老鷹受讓2分", "活塞讓2分"]

    win = grade_pick("活塞讓2分", "老鷹", "活塞", -5)
    assert win["b"] == "活塞讓2分贏" and win["c"] == "活塞贏5分" and win["d"] == SUCCESS
    lose = grade_pick("活塞讓2分", "老鷹", "活塞", -1)
    assert lose["b"] == "活塞讓2分輸" and lose["c"] == "活塞贏1分" and lose["d"] == FAIL
    push = grade_pick("活塞讓2分", "老鷹", "活塞", -2)
    assert push["b"] == "活塞讓2分走" and push["c"] == "活塞贏2分" and push["d"] == PUSH
    dog = grade_pick("老鷹受讓2分", "老鷹", "活塞", -1)
    assert dog["d"] == SUCCESS and dog["c"] == "老鷹輸1分"
    dog_lose = grade_pick("老鷹受讓2分", "老鷹", "活塞", -5)
    assert dog_lose["d"] == FAIL
    assert grade_pick("活塞讓2.5分", "老鷹", "活塞", -2)["d"] == FAIL
    assert grade_pick("活塞讓2.5分", "老鷹", "活塞", -3)["d"] == SUCCESS
    assert resolve_home_margin("客116-主111 客隊贏5", "老鷹", "活塞") == -5
    assert resolve_home_margin("活塞116:老鷹111", "老鷹", "活塞") == -5
    print("nba_pick selftest ok")


if __name__ == "__main__":
    selftest()
