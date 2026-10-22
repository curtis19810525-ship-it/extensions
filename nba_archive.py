# -*- coding: utf-8 -*-
"""
NBA 總表日期與紀錄歸檔。

只在認得出版面時寫入：
  - 「紀錄」A 欄 1～33 列有「第N場」，且「總表」A1 是空的或日期
    → 總表 A1 設成指定日並重算；賽果日再把紀錄 A1:I33 貼上值到 K1:S33，K1 改為賽果日
  - 「總表」A3 是「天數」、A4 是「日期」（一天一欄）
    → 只重算，不搬任何格子
  - 其他版面 → 不寫

不碰「各隊total」「球員狀態」。
"""

from __future__ import annotations

import os
import re
import sys

DEFAULT_EXCEL = r"C:\Users\curti\OneDrive\NBA26-27\NBA26-27爬蟲數據.xlsx"
SHEET_ZONG = "總表"
SHEET_RECORD = "紀錄"
SRC_RANGE = "A1:I33"
DST_RANGE = "K1:S33"
DATE_CELL = "A1"
GAME_LABEL = re.compile(r"^第\d+場")


def excel_path() -> str:
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
            "拒絕寫入：路徑指向 MLB 活頁簿（%s）。" % path
        )
    return path


def looks_like_date(value) -> bool:
    if value is None:
        return False
    if hasattr(value, "year") and hasattr(value, "month"):
        return True
    text = str(value).strip().replace("-", "").replace("/", "")
    return len(text) == 8 and text.isdigit()


def a1_can_hold_date(value) -> bool:
    if value is None or str(value).strip() == "":
        return True
    return looks_like_date(value)


def _game_label_in_column_a(ws) -> bool:
    for row in range(1, 34):
        value = ws.cell(row, 1).value
        if isinstance(value, str) and GAME_LABEL.match(value.strip()):
            return True
    return False


def classify_workbook(wb) -> dict:
    names = list(wb.sheetnames)
    record_ok = (
        SHEET_RECORD in names and _game_label_in_column_a(wb[SHEET_RECORD])
    )
    if record_ok and SHEET_ZONG in names and a1_can_hold_date(wb[SHEET_ZONG]["A1"].value):
        return {
            "kind": "record",
            "detail": "紀錄直欄。總表 A1 寫日期；賽果日把 A1:I33 貼上值到 K1:S33",
            "sheets": names,
        }
    if record_ok:
        return {
            "kind": "record_keep_a1",
            "detail": "紀錄有第N場，但總表 A1 不是日期。只歸檔紀錄，不改總表 A1",
            "sheets": names,
        }
    if SHEET_ZONG in names:
        master = wb[SHEET_ZONG]
        if master["A3"].value == "天數" and master["A4"].value == "日期":
            return {
                "kind": "calendar",
                "detail": "總表日期在第 4 列橫排，每天一欄。只重算，不搬格子",
                "sheets": names,
            }
    sample = ""
    if SHEET_ZONG in names:
        master = wb[SHEET_ZONG]
        sample = "總表 A1=%r A3=%r A4=%r" % (
            master["A1"].value, master["A3"].value, master["A4"].value,
        )
    return {
        "kind": "unknown",
        "detail": "對不上已知版面，未寫入。%s 分頁：%s" % (sample, "、".join(names)),
        "sheets": names,
    }


def classify_path(path: str) -> dict:
    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        return classify_workbook(wb)
    finally:
        wb.close()


def actions_for(kind: str, day: str, phase: str) -> list[str]:
    """phase 為 results（賽果日爬完）或 focus（關注日爬完）。"""
    if kind == "unknown":
        return []
    if kind == "calendar":
        return ["重算整本"]
    if kind == "record_keep_a1":
        if phase == "results":
            return [
                "重算整本",
                "紀錄 %s 貼上值到 %s" % (SRC_RANGE, DST_RANGE),
                "紀錄 K1=%s" % day,
            ]
        return ["重算整本"]
    if phase == "results":
        return [
            "總表 A1=%s" % day,
            "重算整本",
            "紀錄 %s 貼上值到 %s" % (SRC_RANGE, DST_RANGE),
            "紀錄 K1=%s" % day,
        ]
    return ["總表 A1=%s" % day, "重算整本"]


def _require_windows_excel():
    if sys.platform != "win32":
        raise RuntimeError("Excel 重算需在 Windows 本機執行")
    try:
        import win32com.client  # noqa: F401
    except ImportError as exc:
        raise RuntimeError("缺少 pywin32，無法用 Excel 重算") from exc


def _apply_com(path: str, day: str, phase: str, kind: str) -> None:
    import time
    import win32com.client

    _require_windows_excel()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    wb = excel.Workbooks.Open(os.path.abspath(path), UpdateLinks=0, ReadOnly=False)
    try:
        names = [sh.Name for sh in wb.Worksheets]
        if kind in ("record", "record_keep_a1"):
            if SHEET_RECORD not in names:
                raise RuntimeError("分頁在寫入前消失")
            if kind == "record":
                if SHEET_ZONG not in names:
                    raise RuntimeError("分頁在寫入前消失")
                master = wb.Worksheets(SHEET_ZONG)
                master.Range(DATE_CELL).NumberFormat = "@"
                master.Range(DATE_CELL).Value = day
            _recalc(excel)
            if phase == "results":
                record = wb.Worksheets(SHEET_RECORD)
                record.Range(DST_RANGE).ClearContents()
                record.Range(SRC_RANGE).Copy()
                record.Range(DST_RANGE).PasteSpecial(Paste=-4163)
                try:
                    excel.CutCopyMode = False
                except Exception:
                    pass
                record.Range("K1").NumberFormat = "@"
                record.Range("K1").Value = day
        elif kind == "calendar":
            _recalc(excel)
        else:
            raise RuntimeError("未知版面不寫入")
        wb.Save()
    finally:
        try:
            wb.Close(SaveChanges=0)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
        time.sleep(0.8)


def _recalc(excel) -> None:
    try:
        excel.CalculateFullRebuild()
    except Exception:
        try:
            excel.CalculateFull()
        except Exception:
            excel.Calculate()


def run_archive(path: str, day: str, phase: str, dry_run: bool) -> tuple[str, str, str]:
    """回傳 (名稱, OK/FAILED/SKIP, 說明)。dry_run 不開 Excel。"""
    label = "歸檔賽果日" if phase == "results" else "總表改關注日"
    if not os.path.exists(path):
        status = "SKIP" if dry_run else "FAILED"
        return label, status, "找不到活頁簿 %s" % path
    try:
        layout = classify_path(path)
    except Exception as exc:
        return label, "FAILED", "讀不到活頁簿：%s" % exc

    kind = layout["kind"]
    planned = actions_for(kind, day, phase)
    print("[%s] %s" % (label, layout["detail"]))
    if kind == "unknown":
        return label, "FAILED", layout["detail"]
    for step in planned:
        print("  - %s" % step)
    if dry_run:
        return label, "OK", "dry-run " + layout["detail"]
    try:
        _apply_com(path, day, phase, kind)
    except Exception as exc:
        return label, "FAILED", str(exc)
    return label, "OK", layout["detail"]
