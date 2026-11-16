# -*- coding: utf-8 -*-
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BAT = ROOT / "回家日常並合成.bat"


class AutoMergeBat(unittest.TestCase):
    def test_wrapper_runs_home_then_merges(self):
        text = BAT.read_bytes().decode("cp950")
        self.assertIn('cmd /c "%~dp0回家日常.bat"', text)
        self.assertIn('python "%~dp0merge_daily_report.py"', text)
        home_at = text.index("回家日常.bat")
        merge_at = text.index("merge_daily_report.py")
        self.assertLess(home_at, merge_at)


if __name__ == "__main__":
    unittest.main()
