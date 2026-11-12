# -*- coding: utf-8 -*-
"""把 exports 裡的賽後結果與賽前推薦合成一份 UTF-8 報告。"""

from __future__ import annotations

import sys
from datetime import datetime, timedelta
from pathlib import Path

POST_SUFFIX = "賽後結果.txt"
PRE_SUFFIXES = ("賽前推薦.txt", "賽前推票.txt")
SEPARATOR = "========================================="


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError, OSError):
            pass


def read_text(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "cp950"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def normalize_newlines(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").strip("\n")


def report_title(results_day: str, focus_day: str) -> str:
    results = datetime.strptime(results_day, "%Y%m%d")
    focus = datetime.strptime(focus_day, "%Y%m%d")
    return f"{results:%Y%m%d}賽後與{focus:%m%d}賽前"


def build_report(post_text: str, pre_text: str, title: str) -> str:
    body = "\n".join((
        title,
        normalize_newlines(post_text),
        SEPARATOR,
        normalize_newlines(pre_text),
    ))
    return body.replace("\n", "\r\n") + "\r\n"


def report_path(folder: Path, results_day: str, focus_day: str) -> Path:
    return folder / f"{report_title(results_day, focus_day)}.txt"


def _day_from_name(name: str, suffix: str) -> str | None:
    if not name.endswith(suffix):
        return None
    day = name[: -len(suffix)]
    if len(day) != 8 or not day.isdigit():
        return None
    try:
        datetime.strptime(day, "%Y%m%d")
    except ValueError:
        return None
    return day


def _search_folders(base: Path) -> list[Path]:
    folders = []
    exports = base / "exports"
    if exports.is_dir():
        folders.append(exports)
    if base.is_dir():
        folders.append(base)
    return folders


def locate_sources(base: Path, results_day: str | None, focus_day: str | None):
    """回傳 (賽後檔, 賽前檔, 賽果日, 關注日)。找不到時前兩個是 None。"""
    folders = _search_folders(Path(base))
    if results_day:
        datetime.strptime(results_day, "%Y%m%d")
    if focus_day:
        datetime.strptime(focus_day, "%Y%m%d")
    if not results_day:
        newest = None
        for folder in folders:
            for path in folder.glob(f"????????{POST_SUFFIX}"):
                day = _day_from_name(path.name, POST_SUFFIX)
                if day is None:
                    continue
                stamp = path.stat().st_mtime
                if newest is None or stamp >= newest[0]:
                    newest = (stamp, day)
        if newest is None:
            return None, None, None, None
        results_day = newest[1]
    if not focus_day:
        focus_day = (datetime.strptime(results_day, "%Y%m%d") + timedelta(days=1)).strftime("%Y%m%d")
    post_name = f"{results_day}{POST_SUFFIX}"
    for folder in folders:
        post_path = folder / post_name
        if not post_path.is_file():
            continue
        for suffix in PRE_SUFFIXES:
            pre_path = folder / f"{focus_day}{suffix}"
            if pre_path.is_file():
                return post_path, pre_path, results_day, focus_day
    return None, None, results_day, focus_day


def write_report(folder: Path, results_day: str | None = None, focus_day: str | None = None) -> Path | None:
    post_path, pre_path, results_day, focus_day = locate_sources(folder, results_day, focus_day)
    if post_path is None or pre_path is None:
        if results_day and focus_day:
            print(f"在 exports 找不到 {results_day}{POST_SUFFIX} 或 {focus_day}{PRE_SUFFIXES[0]}，未產出綜合報告。")
        else:
            print("在 exports 找不到賽後結果與隔日賽前推薦，未產出綜合報告。")
        return None
    destination = report_path(post_path.parent, results_day, focus_day)
    title = report_title(results_day, focus_day)
    text = build_report(read_text(post_path), read_text(pre_path), title)
    destination.write_bytes(text.encode("utf-8-sig"))
    print(f"綜合報告：{destination}")
    return destination


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    argv = list(sys.argv[1:] if argv is None else argv)
    if len(argv) > 2:
        print("用法：python merge_daily_report.py [賽果日 關注日]")
        return 1
    results_day = argv[0] if argv else None
    focus_day = argv[1] if len(argv) == 2 else None
    folder = Path(__file__).resolve().parent
    try:
        write_report(folder, results_day, focus_day)
    except ValueError as exc:
        print(f"日期請寫成 YYYYMMDD。{exc}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
