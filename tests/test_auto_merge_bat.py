# -*- coding: utf-8 -*-
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BAT = ROOT / "回家日常並合成.bat"
HOME = ROOT / "回家日常.bat"


class AutoMergeBat(unittest.TestCase):
    def test_wrapper_runs_home_then_merges(self):
        text = BAT.read_bytes().decode("cp950")
        self.assertIn('call "%~dp0回家日常.bat"', text)
        self.assertIn("NBA_NO_PAUSE=1", text)
        self.assertIn("NBA_SKIP_MERGE=1", text)
        self.assertNotIn("findstr", text)
        self.assertNotIn("home_pipeline", text)
        self.assertLess(text.index('call "%~dp0回家日常.bat"'), text.index('python "%~dp0merge_daily_report.py"'))

    def test_home_bat_is_ascii_and_runs_nba_home(self):
        raw = HOME.read_bytes()
        text = raw.decode("ascii")
        self.assertIn('python "%~dp0nba_home.py"', text)
        self.assertIn('if "%NBA_NO_PAUSE%"=="1" exit /b 0', text)
        self.assertNotIn("home_pipeline", text)
        self.assertNotIn(".xlsx", text)


if __name__ == "__main__":
    unittest.main()
