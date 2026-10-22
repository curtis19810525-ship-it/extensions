# -*- coding: utf-8 -*-
"""
NBA 回家日常：只串現有爬蟲。

賽果日跑步驟 1～3，並把該日期傳進既有腳本。
關注日跑步驟 4～6。這三步若本來不吃日期，命令列不附加日期，只在畫面上顯示。

賽果日爬完後歸檔，關注日爬完後把總表日期改過去。
對不上總表版面就不寫。不填盤口、不匯出。
某一支失敗仍繼續，最後列成功／失敗。
同資料夾若有 run_all._build_command，命令與手動選單相同。
"""

from __future__ import annotations

import importlib
import subprocess
import sys
from datetime import datetime, timedelta


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr, sys.stdin):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError, OSError):
            pass


def parse_yyyymmdd(value: str) -> str:
    text = (value or "").strip()
    datetime.strptime(text, "%Y%m%d")
    return text


def default_command(step: int, day: str) -> list[str] | None:
    """沒有 run_all.py 時的 NBA 命令。步驟 4～6 不附日期。"""
    if step == 1:
        return [sys.executable, "執行範圍爬蟲.py", day, day]
    if step == 2:
        return [sys.executable, "批次轉換正負盤.py", f"{day}~{day}"]
    if step == 3:
        return [sys.executable, "批次轉換各隊正負.py", f"{day}~{day}"]
    if step == 4:
        return [sys.executable, "espn_player_stats.py"]
    if step == 5:
        return [sys.executable, "espn_team_total.py"]
    if step == 6:
        return [sys.executable, "nba_lineups.py"]
    return None


def load_run_all():
    try:
        return importlib.import_module("run_all")
    except ModuleNotFoundError:
        return None
    except Exception as exc:
        raise SystemExit(f"無法讀取同資料夾 run_all.py：{exc}")


def is_mlb_menu(module) -> bool:
    names = getattr(module, "STEP_NAMES", {}) or {}
    blob = " ".join(str(value) for value in names.values()).upper()
    return "MLB" in blob


def command_for(step: int, day: str, run_all=None) -> list[str] | None:
    """回傳要執行的命令。選單沒有該步時回傳 None。"""
    if run_all is not None and hasattr(run_all, "_build_command"):
        try:
            return list(run_all._build_command(step, day, day))
        except ValueError:
            return None
    return default_command(step, day)


def validate_dates(results_day: str, focus_day: str) -> tuple[str, str, str | None]:
    """回傳 (賽果日, 關注日, 警告)。關注日早於賽果日時丟出 ValueError。"""
    results_day = parse_yyyymmdd(results_day)
    focus_day = parse_yyyymmdd(focus_day)
    results_dt = datetime.strptime(results_day, "%Y%m%d")
    focus_dt = datetime.strptime(focus_day, "%Y%m%d")
    if focus_dt < results_dt:
        raise ValueError("關注日不可早於賽果日")
    warning = None
    gap = (focus_dt - results_dt).days
    if gap != 1:
        warning = f"關注日不是賽果日的隔天（間隔 {gap} 天）"
    return results_day, focus_day, warning


def _keeps_filled_season(cmd) -> bool:
    """各隊 total、球員狀態已用 2025-26 補好，一鍵不重抓。"""
    text = " ".join(str(part) for part in cmd)
    names = (
        "espn_player_stats.py",
        "espn_team_total.py",
        "爬取ESPN球員狀態.py",
        "爬取ESPN各隊Total.py",
    )
    return any(name in text for name in names)


def run_plan(results_day: str, focus_day: str, runner, run_all=None, groups=None):
    """依序執行。runner(cmd) 回傳 exit code。回傳 (步驟, OK/FAILED/SKIP, 說明)。"""
    log = []
    if groups is None:
        groups = (
            ("賽果日", results_day, (1, 2, 3)),
            ("關注日", focus_day, (4, 5, 6)),
        )
    for label, day, steps in groups:
        print()
        print("=" * 60)
        print(f"{label} {day}")
        print("=" * 60)
        for step in steps:
            cmd = command_for(step, day, run_all)
            if cmd is None:
                print(f"步驟 {step}：選單沒有這一步，略過")
                log.append((step, "SKIP", "選單沒有這一步"))
                continue
            shown = " ".join(cmd[1:])
            if _keeps_filled_season(cmd):
                print(f"步驟 {step}：保留已補好的 2025-26，不重抓")
                log.append((step, "SKIP", "保留已補好的 2025-26"))
                continue
            print(f"步驟 {step}：{shown}")
            try:
                code = runner(cmd)
            except OSError as exc:
                print(f"[失敗] {exc}")
                log.append((step, "FAILED", str(exc)))
                continue
            if code == 0:
                log.append((step, "OK", shown))
            else:
                print(f"[失敗] 步驟 {step} 結束碼 {code}，繼續下一步")
                log.append((step, "FAILED", f"結束碼 {code}"))
    return log


def _local_day(offset: int) -> str:
    return (datetime.now() + timedelta(days=offset)).strftime("%Y%m%d")


def _prompt_dates() -> tuple[str, str]:
    default_results = _local_day(0)
    default_focus = _local_day(1)
    print("請輸入賽果日、關注日（YYYYMMDD）。")
    print("  日期請用「每日數據」的開賽日。")
    print(f"  直接按 Enter 可用預設：賽果日 {default_results}、關注日 {default_focus}")
    while True:
        try:
            results_raw = input(f"賽果日 [{default_results}]: ").strip() or default_results
            focus_raw = input(f"關注日 [{default_focus}]: ").strip() or default_focus
        except (EOFError, KeyboardInterrupt):
            raise SystemExit(1)
        try:
            results_day, focus_day, warning = validate_dates(results_raw, focus_raw)
        except ValueError as exc:
            print(f"錯誤：{exc}，請重新輸入。")
            continue
        if warning:
            print(f"警告：{warning}")
            try:
                answer = input("按 Enter 繼續，或輸入 Q 重輸：").strip().upper()
            except (EOFError, KeyboardInterrupt):
                raise SystemExit(1)
            if answer == "Q":
                continue
        return results_day, focus_day


def _print_summary(log: list[tuple[int, str, str]]) -> int:
    print()
    print("=" * 60)
    print("回家日常 Summary")
    print("=" * 60)
    failed = False
    skipped = False
    for step, status, detail in log:
        if status == "FAILED":
            failed = True
        if status == "SKIP":
            skipped = True
        extra = f"  {detail}" if detail else ""
        name = step if isinstance(step, str) else f"步驟 {step}"
        print(f"{status:<6} {name}{extra}")
    print("=" * 60)
    if failed:
        print("有步驟失敗。未填盤口、未匯出記事本。")
        return 1
    if skipped:
        print("有步驟略過。尚未填盤口、尚未匯出記事本。")
        return 0
    print("爬蟲與總表日期跑完。尚未填盤口、尚未匯出記事本。")
    return 0


def subprocess_runner(cmd: list[str]) -> int:
    completed = subprocess.run(cmd)
    return int(completed.returncode)


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    argv = list(sys.argv[1:] if argv is None else argv)
    dry_run = "--dry-run" in argv
    positionals = [item for item in argv if not item.startswith("--")]

    print("=" * 60)
    print("NBA 回家日常（只串現有爬蟲）")
    print("=" * 60)
    print("賽果日：步驟 1～3，然後歸檔。關注日：步驟 4～6，然後改總表日期。")
    print("請先關閉 NBA26-27爬蟲數據.xlsx。")
    print("不填盤口、不匯出。總表版面對不上時不寫格子。")
    print()

    if len(positionals) == 2:
        try:
            results_day, focus_day, warning = validate_dates(positionals[0], positionals[1])
        except ValueError as exc:
            print(f"錯誤：{exc}")
            return 1
        if warning:
            print(f"警告：{warning}")
    elif len(positionals) == 0:
        results_day, focus_day = _prompt_dates()
    else:
        print("用法：python nba_home.py [--dry-run] [賽果日 關注日]")
        return 1

    run_all = load_run_all()
    if run_all is not None and is_mlb_menu(run_all):
        print("這支是 NBA 一鍵。同資料夾的 run_all 是 MLB 選單，已停止。")
        return 1
    if run_all is None:
        print("同資料夾沒有 run_all.py，改用 NBA 預設步驟命令。")
    else:
        print("步驟命令跟同資料夾的 run_all 選單相同。")
    print(f"賽果日 {results_day}　關注日 {focus_day}")

    def runner(cmd: list[str]) -> int:
        if dry_run:
            print("dry-run，未執行")
            return 0
        return subprocess_runner(cmd)

    from nba_archive import excel_path, run_archive

    path = excel_path()
    print(f"Excel: {path}")
    log = []
    log.extend(run_plan(
        results_day, focus_day, runner, run_all,
        groups=(("賽果日", results_day, (1, 2, 3)),),
    ))
    log.append(run_archive(path, results_day, "results", dry_run))
    log.extend(run_plan(
        results_day, focus_day, runner, run_all,
        groups=(("關注日", focus_day, (4, 5, 6)),),
    ))
    log.append(run_archive(path, focus_day, "focus", dry_run))
    return _print_summary(log)


if __name__ == "__main__":
    sys.exit(main())
