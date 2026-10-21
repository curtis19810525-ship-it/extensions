# -*- coding: utf-8 -*-
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from espn_common import cell_value, season_error
from 爬取ESPN各隊Total import parse_team_total
from 爬取ESPN球員狀態 import parse_injuries, parse_player_stats

STATS_HTML = """
<title>Boston Celtics 2025-26 NBA Regular Season Stats - ESPN</title>
<script>{"requestedSeason":{"year":2026,"type":2,"name":"Regular Season","displayName":"2025-26"}}</script>
<table>
  <tr><th>Name</th></tr>
  <tr><td>Jaylen Brown G</td></tr>
  <tr><td>Jayson Tatum F</td></tr>
  <tr><td>Total</td></tr>
</table>
<table>
  <tr><th>GP</th><th>GS</th><th>MIN</th><th>PTS</th><th>OR</th><th>DR</th><th>REB</th><th>AST</th><th>STL</th><th>BLK</th><th>TO</th><th>PF</th><th>AST/TO</th></tr>
  <tr><td>71</td><td>71</td><td>34.4</td><td>28.7</td><td>1.1</td><td>5.8</td><td>6.9</td><td>5.1</td><td>1.0</td><td>0.4</td><td>3.6</td><td>2.7</td><td>1.4</td></tr>
  <tr><td>16</td><td>16</td><td>32.6</td><td>21.8</td><td>0.5</td><td>9.5</td><td>10.0</td><td>5.3</td><td>1.4</td><td>0.2</td><td>0</td><td>1.6</td><td>INF</td></tr>
  <tr><td>82</td><td></td><td></td><td>114.9</td><td>12.5</td><td>33.9</td><td>46.4</td><td>24.6</td><td>7.1</td><td>5.0</td><td>11.5</td><td>19.0</td><td>2.1</td></tr>
</table>
<table>
  <tr><th>Name</th></tr>
  <tr><td>Jaylen Brown G</td></tr>
  <tr><td>Jayson Tatum F</td></tr>
  <tr><td>Total</td></tr>
</table>
<table>
  <tr><th>FGM</th><th>FGA</th><th>FG%</th><th>3PM</th><th>3PA</th><th>3P%</th><th>FTM</th><th>FTA</th><th>FT%</th><th>2PM</th><th>2PA</th><th>2P%</th><th>SC-EFF</th><th>SH-EFF</th></tr>
  <tr><td>10.4</td><td>21.7</td><td>47.7</td><td>2.0</td><td>5.7</td><td>34.7</td><td>6.0</td><td>7.5</td><td>79.5</td><td>8.4</td><td>16.0</td><td>52.3</td><td>1.321</td><td>0.52</td></tr>
  <tr><td>7.4</td><td>17.9</td><td>41.1</td><td>2.9</td><td>8.9</td><td>32.9</td><td>4.1</td><td>4.9</td><td>82.3</td><td>4.4</td><td>9.0</td><td>49.3</td><td>1.213</td><td>0.49</td></tr>
  <tr><td>42.1</td><td>90.2</td><td>46.7</td><td>15.5</td><td>42.1</td><td>36.7</td><td>15.1</td><td>18.7</td><td>80.7</td><td>26.7</td><td>48.1</td><td>55.5</td><td>1.273</td><td>0.55</td></tr>
</table>
"""

WRONG_SEASON_HTML = """
<title>Boston Celtics 2024-25 NBA Regular Season Stats - ESPN</title>
<script>{"requestedSeason":{"year":2025,"type":2,"name":"Regular Season","displayName":"2024-25"}}</script>
"""

INJURY_HTML = """
<div class="Injuries__groupDate">Oct 18</div>
<div class="ContentList__Item">
  <div class="Athlete__PlayerWrapper">
    <span class="Athlete__PlayerName">Devin Carter</span>
    <span>G</span>
    <span>Status</span>
    <span class="TextStatus">OUT</span>
    <span>Carter has been diagnosed with mononucleosis.</span>
  </div>
</div>
"""


class ParseTests(unittest.TestCase):
    def test_locked_season_accepts_2026_regular(self):
        self.assertIsNone(season_error(STATS_HTML))

    def test_locked_season_rejects_other_year(self):
        reason = season_error(WRONG_SEASON_HTML)
        self.assertIn("2024-25", reason)

    def test_empty_page_is_rejected(self):
        self.assertEqual(season_error("   "), "空頁")

    def test_team_total_uses_total_row_not_a_player(self):
        stats = parse_team_total(STATS_HTML)
        self.assertEqual(stats["GP"], "82")
        self.assertEqual(stats["PTS"], "114.9")
        self.assertEqual(stats["REB"], "46.4")
        self.assertEqual(stats["FG%"], "46.7")
        self.assertEqual(stats["DR"], "33.9")
        from 爬取ESPN各隊Total import STAT_FIELDS
        self.assertNotIn("DR", STAT_FIELDS)

    def test_player_rows_skip_total_and_dr(self):
        rows = parse_player_stats(STATS_HTML, "賽爾提克")
        self.assertEqual(len(rows), 2)
        brown, tatum = rows
        self.assertEqual(brown[0], "賽爾提克")
        self.assertEqual(brown[1], "Jaylen Brown G")
        self.assertEqual(brown[2], 71)   # GP
        self.assertEqual(brown[5], 28.7)  # PTS
        self.assertEqual(brown[7], 6.9)   # REB，不是 DR 5.8
        self.assertEqual(brown[13], 1.4)  # AST/TO
        self.assertEqual(tatum[13], 0)    # INF -> 0
        self.assertEqual(brown[14], 10.4)  # FGM

    def test_cell_value_inf(self):
        self.assertEqual(cell_value("INF"), 0)
        self.assertEqual(cell_value(""), "")

    def test_injury_card(self):
        rows = parse_injuries(INJURY_HTML, "賽爾提克")
        self.assertEqual(rows, [[
            "賽爾提克",
            "Devin Carter",
            "Carter has been diagnosed with mononucleosis.",
            "Out",
            "Oct 18",
        ]])


if __name__ == "__main__":
    unittest.main()
