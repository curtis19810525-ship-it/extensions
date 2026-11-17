# -*- coding: utf-8 -*-
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BAT = ROOT / "回家日常並合成.bat"


class AutoMergeBat(unittest.TestCase):
    def test_wrapper_runs_home_then_merges(self):
        text = BAT.read_bytes().decode("cp950")
        self.assertIn('findstr /v /i /r /c:"^ *pause"', text)
        self.assertIn('cmd /c "%temp%\\nba_home_nopause.bat"', text)
        self.assertNotIn('cmd /c "%~dp0回家日常.bat"', text)
        self.assertLess(text.index("nba_home_nopause.bat"), text.index('python "%~dp0merge_daily_report.py"'))


if __name__ == "__main__":
    unittest.main()
