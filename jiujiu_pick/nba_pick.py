# -*- coding: utf-8 -*-
"""賽前預測：AI 與我的推薦各一組選單，賽後分別寫出分差與顏色。"""

from __future__ import annotations

import math
import re

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

PICK_LABEL = "賽前預測"
AI_KEY = "AI推薦"
MINE_KEY = "我的推薦"
LEGACY_KEY = "賽前預測"
NOTE_LABELS = ("賽前想法", "比賽結果", "賽後想法")
SUCCESS = "預測正確"
FAIL = "預測失敗"
PUSH = "預測走盤"
HEADER_LABELS = (
    (2, "AI推薦"),
    (3, "賽後結果"),
    (4, "預測結果"),
    (5, "我的推薦"),
    (6, "賽後結果"),
    (7, "預測結果"),
)

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
NOTE_ALIGN = Alignment(horizontal="left", vertical="top", wrap_text=True)
THICK_BOTTOM = Side(style="thick", color="000000")
NOTE_HEADERS = ("賽前想法", "比賽結果")
DATE_AT_START = re.compile(r"^\s*(\d{1,2})-(\d{1,2})")
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
    r"^(?P<team>.+?)(?P<side>受讓|讓)(?P<pts>\d+(?:\.\d+)?)分"
    r"(?P<tail>(?:[+-](?:25|50|75)|贏(?:25|50|75)|輸(?:25|50|75)))?"
    r"(?P<mark>贏|輸|走)?$"
)


def format_points(value: float) -> str:
    number = abs(float(value))
    if abs(number - round(number)) < 1e-6:
        return str(int(round(number)))
    return f"{number:.2f}".rstrip("0").rstrip(".")


def _is_quarter_line(number: float) -> bool:
    quarters = abs(float(number)) * 4
    return abs(quarters - round(quarters)) < 1e-6


def parse_spread_token(token):
    """取出球頭。括號裡的 +50、-25 是剛好贏到該分時的輸贏成數，不是再加減分數。

    (2-25)0.95 的球頭是 2，0.95 是水位。(2.5) 的球頭是 2.5。
    單獨的 0.95 是水位，不當成讓分。
    """
    if isinstance(token, bool) or token is None:
        return None
    if isinstance(token, (int, float)):
        number = float(token)
        if _is_quarter_line(number):
            return number
        return None
    text = str(token).strip()
    matched = re.search(r"\(([+-]?\d+(?:\.\d+)?)(?:[+-]\d+)?\)", text)
    if not matched:
        return None
    return float(matched.group(1))


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
    return f"{parsed['team']}{parsed['side']}{parsed['pts']}分{parsed.get('tail') or ''}"


def settlement_suffix(token) -> str:
    matched = re.search(r"\([+-]?\d+(?:\.\d+)?([+-])(25|50|75)\)", str(token or ""))
    if not matched:
        return ""
    return matched.group(1) + matched.group(2)


def flip_suffix(suffix: str) -> str:
    if not suffix:
        return ""
    number = suffix[1:]
    if suffix[0] == "-":
        return f"贏{number}"
    if suffix[0] == "+":
        return f"輸{number}"
    return ""


def strip_role(value) -> str:
    text = str(value or "").strip()
    return text.replace("（主）", "").replace("(主)", "").strip()


def menu_options(home: str, away: str, signed_home, suffix="", suffix_on_home=True):
    if signed_home is None or not home or not away:
        return []
    points = format_points(signed_home)
    home_lays = float(signed_home) > 0 or float(signed_home) == 0
    if home_lays:
        lay_team, dog_team = home, away
        raw_on_lay = bool(suffix_on_home)
    else:
        lay_team, dog_team = away, home
        raw_on_lay = not bool(suffix_on_home)
    lay_tail = suffix if raw_on_lay else flip_suffix(suffix)
    dog_tail = flip_suffix(suffix) if raw_on_lay else suffix
    return [f"{lay_team}讓{points}分{lay_tail}", f"{dog_team}受讓{points}分{dog_tail}"]


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


def _slot_for(game_notes, title: str):
    if not isinstance(game_notes, dict):
        return None
    title = str(title or "").strip()
    exact = game_notes.get(title)
    if isinstance(exact, dict):
        return exact
    for key, slot in game_notes.items():
        key_text = str(key or "").strip()
        if isinstance(slot, dict) and key_text and (key_text == title or key_text in title or title in key_text):
            return slot
    return None


def lookup_pick(game_notes, title: str, key: str = MINE_KEY) -> str:
    slot = _slot_for(game_notes, title)
    if not isinstance(slot, dict):
        return ""
    if key in slot and slot.get(key):
        return base_pick(slot.get(key))
    if key == MINE_KEY:
        return base_pick(slot.get(LEGACY_KEY))
    return ""


def remember_pick(game_notes, title: str, pick: str, key: str = MINE_KEY):
    if not isinstance(game_notes, dict):
        return
    slot = game_notes.get(title)
    if not isinstance(slot, dict):
        slot = {}
        game_notes[title] = slot
    slot[key] = pick
    if key == MINE_KEY:
        slot[LEGACY_KEY] = pick


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


def _line_from_token(ws, header, col, raw, home, away):
    header_text = ws.cell(header, col).value
    suffix = settlement_suffix(raw)
    if col == 4 or (home and home in str(header_text or "")):
        return signed_from_display(raw, home, header_text), suffix, True
    away_signed = signed_from_display(raw, away, header_text)
    signed = None if away_signed is None else -float(away_signed)
    return signed, suffix, False


def read_latest_line(ws, pick_row: int):
    header = find_header_row(ws, pick_row)
    if header is None:
        return "", "", None, "", True
    home = strip_role(ws.cell(header, 2).value)
    away = strip_role(ws.cell(header, 3).value)
    spread_col = find_spread_column(ws, header)
    for row in range(pick_row - 1, header, -1):
        marker = str(ws.cell(row, 1).value or "").strip()
        second = str(ws.cell(row, 2).value or "").strip()
        if marker in NOTE_LABELS or marker == PICK_LABEL or second == AI_KEY:
            continue
        found = None
        for col in range(4, 16):
            raw = ws.cell(row, col).value
            points = parse_spread_token(raw)
            if points is not None and "(" in str(raw or ""):
                found = (col, raw)
                break
        if found is None and spread_col:
            raw = ws.cell(row, spread_col).value
            if isinstance(raw, (int, float)) and not isinstance(raw, bool) and _is_quarter_line(float(raw)):
                return home, away, float(raw), "", True
        if found is not None:
            signed, suffix, on_home = _line_from_token(ws, header, found[0], found[1], home, away)
            return home, away, signed, suffix, on_home
    return home, away, None, "", True


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


def _write_side(ws, row, pick_col, saved, graded, options):
    pick_cell = ws.cell(row, pick_col, saved or None)
    pick_cell.font = VALUE_FONT
    pick_cell.fill = PatternFill(fill_type=None)
    _apply_border(pick_cell)
    margin_cell = ws.cell(row, pick_col + 1)
    status_cell = ws.cell(row, pick_col + 2)
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
        return
    _clear_status(margin_cell)
    _clear_status(status_cell)
    menu = list(options)
    if saved and saved not in menu:
        menu.insert(0, saved)
    _add_menu(ws, pick_cell, menu)


def _generic_spread_header(value) -> bool:
    text = str(value or "").strip()
    if not text:
        return False
    if text in {"讓分", "受讓", "讓分A"}:
        return True
    return text.startswith(("主讓", "客讓", "主受", "客受"))


def _spread_cell(value) -> bool:
    return parse_spread_token(value) is not None and "(" in str(value or "")


def label_spread_headers(ws, pick_row: int):
    """主隊讓分時，D/E 仍可能停在「主讓分／客讓分」。用最後一筆有球頭的列改成隊名。"""
    header = find_header_row(ws, pick_row)
    if header is None:
        return
    home = strip_role(ws.cell(header, 2).value)
    away = strip_role(ws.cell(header, 3).value)
    if not home or not away:
        return
    home_header = ws.cell(header, 4).value
    away_header = ws.cell(header, 5).value
    if not (_generic_spread_header(home_header) or _generic_spread_header(away_header)):
        return
    latest = None
    for row in range(pick_row - 1, header, -1):
        marker = str(ws.cell(row, 1).value or "").strip()
        second = str(ws.cell(row, 2).value or "").strip()
        if marker in NOTE_LABELS or marker == PICK_LABEL or second == AI_KEY:
            continue
        left = ws.cell(row, 4).value
        right = ws.cell(row, 5).value
        if _spread_cell(left) or _spread_cell(right):
            latest = (left, right)
            break
    if latest is None:
        return
    if _spread_cell(latest[0]) and not _spread_cell(latest[1]):
        ws.cell(header, 4).value = f"{home}讓分"
        ws.cell(header, 5).value = f"{away}受讓"
    elif _spread_cell(latest[1]) and not _spread_cell(latest[0]):
        ws.cell(header, 4).value = f"{home}受讓"
        ws.cell(header, 5).value = f"{away}讓分"


def render_pick(ws, title, row, game_notes, last_col=None):
    label_spread_headers(ws, row)
    home, away, signed, suffix, on_home = read_latest_line(ws, row)
    options = menu_options(home, away, signed, suffix, on_home)
    result = result_below(ws, row)
    if not result and isinstance(game_notes, dict):
        slot = game_notes.get(title)
        if isinstance(slot, dict):
            result = slot.get("比賽結果")
    margin = resolve_home_margin(result, home, away)

    ws.row_dimensions[row].height = 32
    label = ws.cell(row, 1, PICK_LABEL)
    label.font = LABEL_FONT
    label.fill = LABEL_FILL
    _apply_border(label)
    for col, text in HEADER_LABELS:
        cell = ws.cell(row, col, text)
        cell.font = LABEL_FONT
        cell.fill = LABEL_FILL
        _apply_border(cell)

    value_row = row + 1
    ws.row_dimensions[value_row].height = 48
    ws.cell(value_row, 1).value = None
    for key, col in ((AI_KEY, 2), (MINE_KEY, 5)):
        saved = lookup_pick(game_notes, title, key)
        graded = grade_pick(saved, home, away, margin) if saved and margin is not None else None
        _write_side(ws, value_row, col, saved, graded, options)
        if saved:
            remember_pick(game_notes, title, saved, key)
    return value_row + 1


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
            slot = notes.setdefault(current, {})
            if not isinstance(slot, dict):
                continue
            if str(ws.cell(row, 2).value or "") == AI_KEY:
                slot[AI_KEY] = base_pick(ws.cell(row + 1, 2).value)
                slot[MINE_KEY] = base_pick(ws.cell(row + 1, 5).value)
                slot[LEGACY_KEY] = slot[MINE_KEY]
            else:
                slot[MINE_KEY] = base_pick(ws.cell(row, 2).value)
                slot[LEGACY_KEY] = slot[MINE_KEY]
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


def _named_column(ws, header_row: int, name: str, create: bool):
    last = ws.max_column or 1
    for col in range(1, last + 1):
        if ws.cell(header_row, col).value == name:
            return col
    if not create:
        return None
    col = last + 1
    cell = ws.cell(header_row, col, name)
    cell.font = LABEL_FONT
    return col


def _row_date(title):
    match = DATE_AT_START.match(str(title or ""))
    if not match:
        return None
    return (int(match.group(1)), int(match.group(2)))


def _text_units(text) -> int:
    units = 0
    for char in str(text or ""):
        units += 2 if ord(char) > 0x2E80 else 1
    return units


def _wrapped_lines(text, width: float) -> int:
    usable = max(4.0, float(width) - 1.0)
    raw = "" if text is None else str(text).replace("\r\n", "\n")
    if raw.strip() == "":
        return 1
    lines = 0
    for para in raw.split("\n"):
        units = _text_units(para)
        lines += 1 if units == 0 else max(1, math.ceil(units / usable))
    return lines


def _column_width(ws, col: int) -> float:
    letter = ws.cell(1, col).column_letter
    width = ws.column_dimensions[letter].width
    if not isinstance(width, (int, float)) or width < 16:
        width = 32
        ws.column_dimensions[letter].width = width
    return float(width)


def _with_bottom(cell, side):
    current = cell.border
    cell.border = Border(
        left=current.left,
        right=current.right,
        top=current.top,
        bottom=side,
    )


def _needed_height(cell, width: float) -> float:
    size = cell.font.size if isinstance(getattr(cell.font, "size", None), (int, float)) else 11
    lines = _wrapped_lines(cell.value, width)
    # 手機 Excel 會把最後一行裁掉，多留一行的高度。
    if str(cell.value or "").strip() and lines >= 2:
        lines += 1
    return min(409, lines * (float(size) + 4) + 4)


def fit_note_archive(ws):
    """賽前想法、比賽結果依文字換行並調整列高，每天最後一列畫粗底線。"""
    header_row, title_col = _archive_layout(ws)
    if not header_row or not title_col:
        return
    note_cols = []
    last_col = 1
    for col in range(1, (ws.max_column or 1) + 1):
        value = ws.cell(header_row, col).value
        if value not in (None, ""):
            last_col = col
        if value in NOTE_HEADERS:
            note_cols.append(col)
    if not note_cols:
        note_cols = [2, 3]
        last_col = max(last_col, 3)
    widths = {col: _column_width(ws, col) for col in note_cols}
    dated = []
    for row in range(header_row + 1, (ws.max_row or header_row) + 1):
        title = str(ws.cell(row, title_col).value or "").strip()
        if not title:
            continue
        day = _row_date(title)
        if day is not None:
            dated.append((row, day))
        height = 18.0
        for col in note_cols:
            cell = ws.cell(row, col)
            cell.alignment = NOTE_ALIGN
            height = max(height, _needed_height(cell, widths[col]))
        ws.row_dimensions[row].height = height
    boundaries = set()
    for index, (row, day) in enumerate(dated):
        nxt = dated[index + 1][1] if index + 1 < len(dated) else None
        if day != nxt:
            boundaries.add(row)
    for row in range(header_row + 1, (ws.max_row or header_row) + 1):
        title = str(ws.cell(row, title_col).value or "").strip()
        if not title:
            continue
        side = THICK_BOTTOM if row in boundaries else None
        for col in range(1, last_col + 1):
            cell = ws.cell(row, col)
            current = cell.border.bottom.style if cell.border.bottom is not None else None
            if side is None and current != "thick":
                continue
            _with_bottom(cell, side if side is not None else Side())


def archive_into(notes, wb, seen):
    if not isinstance(notes, dict) or "賽事筆記" not in wb.sheetnames:
        return
    ws = wb["賽事筆記"]
    header_row, title_col = _archive_layout(ws)
    if not header_row:
        return
    ai_col = _named_column(ws, header_row, AI_KEY, create=False)
    mine_col = _named_column(ws, header_row, MINE_KEY, create=False)
    legacy_col = _named_column(ws, header_row, LEGACY_KEY, create=False)
    if not ai_col and not mine_col and not legacy_col:
        return
    for row in range(header_row + 1, (ws.max_row or header_row) + 1):
        title = str(ws.cell(row, title_col).value or "").strip()
        if not title or title in seen:
            continue
        slot = notes.setdefault(title, {})
        if not isinstance(slot, dict):
            continue
        if ai_col and not slot.get(AI_KEY):
            slot[AI_KEY] = base_pick(ws.cell(row, ai_col).value)
        mine = base_pick(ws.cell(row, mine_col).value) if mine_col else ""
        if not mine and legacy_col:
            mine = base_pick(ws.cell(row, legacy_col).value)
        if mine and not slot.get(MINE_KEY):
            slot[MINE_KEY] = mine
            slot[LEGACY_KEY] = mine


def archive_picks(wb, game_notes):
    if not isinstance(game_notes, dict) or "賽事筆記" not in getattr(wb, "sheetnames", []):
        return
    ws = wb["賽事筆記"]
    header_row, title_col = _archive_layout(ws)
    if not header_row:
        return
    columns = {
        AI_KEY: _named_column(ws, header_row, AI_KEY, create=True),
        MINE_KEY: _named_column(ws, header_row, MINE_KEY, create=True),
    }
    rows = {}
    for row in range(header_row + 1, (ws.max_row or header_row) + 1):
        title = str(ws.cell(row, title_col).value or "").strip()
        if title:
            rows[title] = row
    for title, slot in game_notes.items():
        if not isinstance(slot, dict):
            continue
        if AI_KEY not in slot and MINE_KEY not in slot and LEGACY_KEY not in slot:
            continue
        stored = {
            AI_KEY: base_pick(slot.get(AI_KEY)),
            MINE_KEY: base_pick(slot.get(MINE_KEY) or slot.get(LEGACY_KEY)),
        }
        title_text = str(title).strip()
        target = rows.get(title_text)
        if target is None:
            for key, row in rows.items():
                if key and (key in title_text or title_text in key):
                    target = row
                    break
        if target is None:
            if not stored[AI_KEY] and not stored[MINE_KEY]:
                continue
            target = (ws.max_row or header_row) + 1
            ws.cell(target, title_col, title_text)
            rows[title_text] = target
        for key, col in columns.items():
            ws.cell(target, col, stored[key] or None)
    fit_note_archive(ws)


def selftest():
    assert format_points(2) == "2"
    assert format_points(14.5) == "14.5"
    assert parse_spread_token("(7+50)") == 7
    assert parse_spread_token("(2-25)0.95") == 2
    assert parse_spread_token("(2.5)0.95") == 2.5
    assert parse_spread_token("0.95") is None
    assert parse_spread_token("(-2)0.9") == -2
    assert base_pick("活塞讓2分贏") == "活塞讓2分"
    assert base_pick("快艇受讓7分贏25") == "快艇受讓7分贏25"
    assert base_pick("灰熊讓7分-25") == "灰熊讓7分-25"
    assert menu_options("老鷹", "活塞", -2) == ["活塞讓2分", "老鷹受讓2分"]
    assert menu_options("灰熊", "快艇", 7, "-25", True) == ["灰熊讓7分-25", "快艇受讓7分贏25"]
    covered = grade_pick("灰熊讓7分-25", "灰熊", "快艇", 18)
    assert covered["c"] == "灰熊贏18分" and covered["d"] == SUCCESS
    missed = grade_pick("快艇受讓7分贏25", "灰熊", "快艇", 18)
    assert missed["c"] == "快艇輸18分" and missed["d"] == FAIL

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
