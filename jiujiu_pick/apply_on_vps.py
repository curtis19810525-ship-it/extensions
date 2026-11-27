# -*- coding: utf-8 -*-
"""把賽前預測掛上 VPS 的 /opt/jiujiu-scraper，不改原有三行筆記的寫法。"""

from __future__ import annotations

import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

MARKER = "nba_pick_installed"
WRAPPER = '''
# nba_pick_installed
# pick_two_rows
import nba_pick as _nba_pick

_nba_pick_orig_write = write_game_notes
_nba_pick_orig_merge = merged_game_notes
_nba_pick_orig_archive = write_game_note_archive
_NBA_PICK_KEYS = ("賽前預測", "AI推薦", "我的推薦")


def _nba_pick_pop(game_notes):
    saved = {}
    if not isinstance(game_notes, dict):
        return saved
    for key, slot in list(game_notes.items()):
        if not isinstance(slot, dict):
            continue
        kept = {name: slot.pop(name) for name in _NBA_PICK_KEYS if name in slot}
        if kept:
            saved[key] = kept
    return saved


def _nba_pick_put(game_notes, saved):
    if not isinstance(game_notes, dict):
        return
    for key, value in saved.items():
        slot = game_notes.get(key)
        if not isinstance(slot, dict):
            slot = {}
            game_notes[key] = slot
        if isinstance(value, dict):
            slot.update(value)


def write_game_notes(*args, **kwargs):
    ws = args[0] if args else kwargs.get("ws")
    game_notes = args[3] if len(args) > 3 else kwargs.get("game_notes")
    last_col = args[4] if len(args) > 4 else kwargs.get("last_col")
    second = args[1] if len(args) > 1 else None
    third = args[2] if len(args) > 2 else None
    row_index = None
    if isinstance(second, int):
        row, title = second, third
        row_index = 1
    elif isinstance(third, int):
        title, row = second, third
        row_index = 2
    else:
        row = kwargs.get("row")
        title = second if second is not None else kwargs.get("title")
    call_args = list(args)
    call_kwargs = dict(kwargs)
    if isinstance(row, int):
        if row_index is not None and len(call_args) > row_index:
            call_args[row_index] = row + 2
        elif "row" in call_kwargs:
            call_kwargs["row"] = row + 2
    saved = _nba_pick_pop(game_notes)
    try:
        written = _nba_pick_orig_write(*call_args, **call_kwargs)
    finally:
        _nba_pick_put(game_notes, saved)
    try:
        if isinstance(row, int):
            _nba_pick.render_pick(ws, title, row, game_notes, last_col)
    except Exception as exc:
        print(f"賽前預測略過：{exc}")
    return written


def merged_game_notes(*args, **kwargs):
    notes = _nba_pick_orig_merge(*args, **kwargs)
    try:
        wb = args[0] if args else kwargs.get("wb")
        if wb is not None and isinstance(notes, dict):
            _nba_pick.merge_picks(wb, notes)
    except Exception as exc:
        print(f"賽前預測讀取略過：{exc}")
    return notes


def write_game_note_archive(*args, **kwargs):
    wb = args[0] if args else kwargs.get("wb")
    game_notes = args[1] if len(args) > 1 else kwargs.get("game_notes")
    saved = _nba_pick_pop(game_notes)
    try:
        result = _nba_pick_orig_archive(*args, **kwargs)
    finally:
        _nba_pick_put(game_notes, saved)
    try:
        if wb is not None and isinstance(game_notes, dict):
            _nba_pick.archive_picks(wb, game_notes)
    except Exception as exc:
        print(f"賽前預測存檔略過：{exc}")
    return result
'''


def python_for(root: Path) -> str:
    venv = root / "venv" / "bin" / "python"
    if venv.exists():
        return str(venv)
    return sys.executable


def install(root: Path, module_source: Path) -> str:
    root = Path(root)
    notes = root / "nba_game_notes.py"
    if not notes.exists():
        raise SystemExit(f"找不到 {notes}")
    if not module_source.exists():
        raise SystemExit(f"找不到 {module_source}")
    text = notes.read_text(encoding="utf-8")
    for name in ("def merged_game_notes", "def write_game_notes", "def write_game_note_archive"):
        if name not in text:
            raise SystemExit(f"stop: {notes.name} 沒有 {name}")
    dest = root / "nba_pick.py"
    previous_module = dest.read_text(encoding="utf-8") if dest.exists() else None
    dest.write_text(module_source.read_text(encoding="utf-8"), encoding="utf-8")
    backup = None
    if MARKER not in text:
        stamp = datetime.now().strftime("%Y%m%d%H%M%S")
        backup = notes.with_name(f"nba_game_notes.py.bak-pick-{stamp}")
        shutil.copy2(notes, backup)
        if not text.endswith("\n"):
            text += "\n"
        notes.write_text(text + "\n" + WRAPPER.lstrip("\n"), encoding="utf-8")
        print(f"已掛上賽前預測，備份 {backup.name}")
    elif "pick_two_rows" not in text:
        stamp = datetime.now().strftime("%Y%m%d%H%M%S")
        backup = notes.with_name(f"nba_game_notes.py.bak-pick-{stamp}")
        shutil.copy2(notes, backup)
        head = text.split("# " + MARKER, 1)[0].rstrip()
        notes.write_text(head + "\n\n" + WRAPPER.lstrip("\n"), encoding="utf-8")
        print(f"已更新成 AI 與我的預測兩列，備份 {backup.name}")
    else:
        print("賽前預測已經掛過，只更新 nba_pick.py")
    _widen_period_block(root)
    try:
        py = python_for(root)
        subprocess.check_call([py, "-m", "py_compile", str(dest), str(notes)])
        subprocess.check_call([py, "-c", "import nba_pick; nba_pick.selftest()"], cwd=str(root))
    except Exception:
        if previous_module is None and dest.exists():
            dest.unlink()
        elif previous_module is not None:
            dest.write_text(previous_module, encoding="utf-8")
        if backup is not None:
            shutil.copy2(backup, notes)
        raise
    return py


def _widen_period_block(root: Path):
    path = root / "period_sheet.py"
    if not path.exists():
        return
    text = path.read_text(encoding="utf-8")
    if "MIN_BLOCK_ROWS) + 5" in text:
        return
    for old in ("MIN_BLOCK_ROWS) + 4", "MIN_BLOCK_ROWS) + 3"):
        if text.count(old) == 1:
            path.write_text(text.replace(old, "MIN_BLOCK_ROWS) + 5", 1), encoding="utf-8")
            print("已加高每場列數，容納 AI 與我的預測")
            return


def restart_service():
    subprocess.check_call(["systemctl", "restart", "jiujiu-trigger"])
    state = subprocess.check_output(["systemctl", "is-active", "jiujiu-trigger"], text=True).strip()
    print(state)
    if state != "active":
        raise SystemExit(f"jiujiu-trigger 狀態是 {state}")


def main():
    root = Path("/opt/jiujiu-scraper")
    source = Path(__file__).resolve().parent / "nba_pick.py"
    install(root, source)
    restart_service()
    print("賽前預測已裝上。下一輪寫入後，每一場會有 AI推薦 與 我的推薦 兩組選單。")


if __name__ == "__main__":
    main()
