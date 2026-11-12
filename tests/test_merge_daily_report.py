# -*- coding: utf-8 -*-
import tempfile
import unittest
from pathlib import Path

import merge_daily_report


class MergeDailyReport(unittest.TestCase):
    def test_joins_post_and_pre_with_separator_and_bom(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "今日賽後.txt").write_text("今日賽後比分與復盤數據\n熱火 117", encoding="utf-8")
            (root / "明日賽前.txt").write_bytes("明日賽前盤口走勢與量化分析\n魔術 +2".encode("cp950"))
            destination = merge_daily_report.write_report(root, "20261112")
            self.assertEqual(destination.name, "20261112_日常綜合報告.txt")
            raw = destination.read_bytes()
            self.assertTrue(raw.startswith(b"\xef\xbb\xbf"))
            text = raw.decode("utf-8-sig")
            self.assertEqual(
                text,
                "今日賽後比分與復盤數據\r\n熱火 117\r\n"
                "=========================================\r\n"
                "明日賽前盤口走勢與量化分析\r\n魔術 +2\r\n",
            )

    def test_missing_source_does_not_write(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "今日賽後.txt").write_text("只有賽後", encoding="utf-8")
            self.assertIsNone(merge_daily_report.write_report(root, "20261112"))
            self.assertEqual(list(root.glob("*綜合報告.txt")), [])


if __name__ == "__main__":
    unittest.main()
