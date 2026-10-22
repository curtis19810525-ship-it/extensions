# -*- coding: utf-8 -*-
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import nba_home


class Dates(unittest.TestCase):
    def test_focus_before_results_is_rejected(self):
        with self.assertRaises(ValueError):
            nba_home.validate_dates("20261022", "20261021")

    def test_next_day_has_no_warning(self):
        results, focus, warning = nba_home.validate_dates("20261021", "20261022")
        self.assertEqual(results, "20261021")
        self.assertEqual(focus, "20261022")
        self.assertIsNone(warning)

    def test_bad_calendar_day(self):
        with self.assertRaises(ValueError):
            nba_home.parse_yyyymmdd("20260231")


class Commands(unittest.TestCase):
    def test_results_steps_receive_the_date(self):
        crawler = nba_home.default_command(1, "20261021")
        ox = nba_home.default_command(2, "20261021")
        teams = nba_home.default_command(3, "20261021")
        self.assertEqual(crawler[1:], ["執行範圍爬蟲.py", "20261021", "20261021"])
        self.assertEqual(ox[1:], ["批次轉換正負盤.py", "20261021~20261021"])
        self.assertEqual(teams[1:], ["批次轉換各隊正負.py", "20261021~20261021"])

    def test_focus_steps_do_not_receive_the_date(self):
        for step in (4, 5, 6):
            cmd = nba_home.default_command(step, "20261022")
            self.assertNotIn("20261022", cmd[1:])

    def test_menu_command_is_used_when_run_all_exists(self):
        class Menu:
            def _build_command(self, step, start, end):
                if step == 6:
                    raise ValueError("no step")
                return ["python", f"step{step}.py", start]

        self.assertEqual(
            nba_home.command_for(1, "20261021", Menu()),
            ["python", "step1.py", "20261021"],
        )
        self.assertIsNone(nba_home.command_for(6, "20261022", Menu()))

    def test_mlb_menu_is_recognized(self):
        class Mlb:
            STEP_NAMES = {4: "ESPN MLB Batting"}

        class Nba:
            STEP_NAMES = {4: "ESPN Player Stats/Injuries"}

        self.assertTrue(nba_home.is_mlb_menu(Mlb()))
        self.assertFalse(nba_home.is_mlb_menu(Nba()))


class Plan(unittest.TestCase):
    def test_failure_does_not_stop_later_steps(self):
        calls = []

        def runner(cmd):
            calls.append(cmd[1])
            return 1 if cmd[1] == "批次轉換正負盤.py" else 0

        log = nba_home.run_plan("20261021", "20261022", runner, run_all=None)
        self.assertEqual(
            [name for name, _status, _detail in ((item[0], item[1], item[2]) for item in log)],
            [1, 2, 3, 4, 5, 6],
        )
        self.assertIn("批次轉換正負盤.py", calls)
        self.assertIn("espn_player_stats.py", calls)
        self.assertEqual(log[1][1], "FAILED")
        self.assertEqual(log[2][1], "OK")

    def test_missing_step_is_skip_not_failure(self):
        class Menu:
            def _build_command(self, step, start, end):
                if step == 6:
                    raise ValueError("missing")
                return [sys.executable, f"s{step}.py"]

        log = nba_home.run_plan("20261021", "20261022", lambda cmd: 0, Menu())
        self.assertEqual(log[-1], (6, "SKIP", "選單沒有這一步"))

    def test_dry_run_main_does_not_launch_scripts(self):
        launched = []
        original = nba_home.subprocess_runner
        nba_home.subprocess_runner = lambda cmd: launched.append(cmd) or 0
        try:
            code = nba_home.main(["--dry-run", "20261021", "20261022"])
        finally:
            nba_home.subprocess_runner = original
        self.assertEqual(code, 0)
        self.assertEqual(launched, [])


if __name__ == "__main__":
    unittest.main()
