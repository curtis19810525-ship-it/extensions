# -*- coding: utf-8 -*-
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BAT = ROOT / "回家日常並合成.bat"
HOME = ROOT / "回家日常.bat"


class AutoMergeBat(unittest.TestCase):
    def test_wrapper_runs_home_pipeline_then_merges(self):
        text = BAT.read_bytes().decode("cp950")
        self.assertIn('python "%~dp0home_pipeline.py"', text)
        self.assertIn('python "%~dp0merge_daily_report.py"', text)
        self.assertIn("NBA26-27爬蟲數據.xlsx", text)
        self.assertIn("玖九盤口變化.xlsx", text)
        self.assertIn("盤口觀察.xlsx", text)
        self.assertNotIn("nba_home.py", text)
        self.assertNotIn("findstr", text)
        self.assertNotIn("Temp", text)
        self.assertLess(
            text.index('python "%~dp0home_pipeline.py"'),
            text.index('python "%~dp0merge_daily_report.py"'),
        )

    def test_home_bat_runs_home_pipeline(self):
        text = HOME.read_bytes().decode("cp950")
        self.assertIn('python "%~dp0home_pipeline.py"', text)
        self.assertNotIn("nba_home.py", text)
        self.assertNotIn("merge_daily_report.py", text)
        self.assertNotIn("Temp", text)


if __name__ == "__main__":
    unittest.main()
