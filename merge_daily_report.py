# -*- coding: utf-8 -*-
"""把今日賽後.txt 與明日賽前.txt 合成一份 UTF-8 綜合報告。"""

from __future__ import annotations

import sys
from datetime import datetime, timedelta
from pathlib import Path

POST_NAME = "今日賽後.txt"
PRE_NAME = "明日賽前.txt"
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


def build_report(post_text: str, pre_text: str) -> str:
    body = "\n".join((
        normalize_newlines(post_text),
        SEPARATOR,
        normalize_newlines(pre_text),
    ))
    return body.replace("\n", "\r\n") + "\r\n"


def report_path(folder: Path, results_day: str, focus_day: str) -> Path:
    results = datetime.strptime(results_day, "%Y%m%d")
    focus = datetime.strptime(focus_day, "%Y%m%d")
    return folder / f"{results:%Y%m%d}賽後與{focus:%m%d}賽前.txt"


def write_report(folder: Path, results_day: str | None = None, focus_day: str | None = None) -> Path | None:
    folder = Path(folder)
    post_path = folder / POST_NAME
    pre_path = folder / PRE_NAME
    if not post_path.is_file() or not pre_path.is_file():
        missing = [name for name, path in ((POST_NAME, post_path), (PRE_NAME, pre_path)) if not path.is_file()]
        print("找不到" + "、".join(missing) + "，未產出綜合報告。")
        return None
    if not results_day:
        results_day = datetime.now().strftime("%Y%m%d")
    if not focus_day:
        focus_day = (datetime.strptime(results_day, "%Y%m%d") + timedelta(days=1)).strftime("%Y%m%d")
    destination = report_path(folder, results_day, focus_day)
    text = build_report(read_text(post_path), read_text(pre_path))
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
