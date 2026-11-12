# -*- coding: utf-8 -*-
import tempfile
import unittest
from pathlib import Path

import merge_daily_report


class MergeDailyReport(unittest.TestCase):
    def test_joins_post_and_pre_with_separator_and_bom(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            exports = root / "exports"
            exports.mkdir()
            (exports / "20261112賽後結果.txt").write_text("今日賽後比分與復盤數據\n熱火 117", encoding="utf-8")
            (exports / "20261113賽前推票.txt").write_bytes("明日賽前盤口走勢與量化分析\n魔術 +2".encode("cp950"))
            destination = merge_daily_report.write_report(root, "20261112", "20261113")
            self.assertEqual(destination.parent, exports)
            self.assertEqual(destination.name, "20261112賽後與1113賽前.txt")
            raw = destination.read_bytes()
            self.assertTrue(raw.startswith(b"\xef\xbb\xbf"))
            text = raw.decode("utf-8-sig")
            self.assertEqual(
                text,
                "今日賽後比分與復盤數據\r\n熱火 117\r\n"
                "=========================================\r\n"
                "明日賽前盤口走勢與量化分析\r\n魔術 +2\r\n",
            )

    def test_without_dates_uses_the_newest_postgame_file(self):
        import os
        import time
        with tempfile.TemporaryDirectory() as folder:
            exports = Path(folder) / "exports"
            exports.mkdir()
            older = exports / "20261111賽後結果.txt"
            newer = exports / "20261112賽後結果.txt"
            older.write_text("舊", encoding="utf-8")
            newer.write_text("新賽後", encoding="utf-8")
            (exports / "20261113賽前推票.txt").write_text("新賽前", encoding="utf-8")
            later = time.time() + 5
            os.utime(newer, (later, later))
            destination = merge_daily_report.write_report(Path(folder))
            self.assertEqual(destination.name, "20261112賽後與1113賽前.txt")

    def test_focus_month_and_day_stay_four_digits(self):
        path = merge_daily_report.report_path(Path("."), "20260105", "20260106")
        self.assertEqual(path.name, "20260105賽後與0106賽前.txt")

    def test_missing_source_does_not_write(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            exports = root / "exports"
            exports.mkdir()
            (exports / "20261112賽後結果.txt").write_text("只有賽後", encoding="utf-8")
            self.assertIsNone(merge_daily_report.write_report(root, "20261112", "20261113"))
            self.assertFalse((exports / "20261112賽後與1113賽前.txt").exists())


if __name__ == "__main__":
    unittest.main()
