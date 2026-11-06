# -*- coding: utf-8 -*-
import hashlib
import importlib.util
import json
import os
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "日記範例" / "NBA26-27_Quant_Betting_Journal.xlsx"
SCRIPT = ROOT / "日記範例" / "append_nba_journal.py"
BAT = ROOT / "日記範例" / "寫入量化日記.bat"


def load_module():
    spec = importlib.util.spec_from_file_location("append_nba_journal", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


journal = load_module()


def file_hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def sample_package():
    return {
        "date": "2026-11-04",
        "climate": "氣候D (冷卻重置期)",
        "pre_game_text": "【2026-11-04 賽前分析】\n大盤氣候 D。",
        "post_game_text": "【2026-11-04 賽後覆盤】\n今日兩場。",
        "games": [
            {
                "game_id": "G1",
                "matchup": "湖人客 @ 勇士主",
                "climate": "氣候D (冷卻重置期)",
                "pick": "湖人 -3",
                "line_odds": "B盤 0.914",
                "cf_pct": "55%",
                "signal_type": "高溫送水盲區",
                "stake_u": "1U",
                "score": "95-112",
                "margin": -17,
                "outcome": "輸",
                "review_notes": "湖人進攻冰凍",
            },
            {
                "場次": "G2",
                "對戰組合": "當日鎖倉",
                "氣候": "氣候D (冷卻重置期)",
                "推薦": "鎖倉不進場",
                "注碼": 0,
                "結果": "鎖倉",
                "檢討": "這張不進場",
            },
        ],
    }


class AppendJournalTests(unittest.TestCase):
    def test_bat_is_ascii(self):
        BAT.read_bytes().decode("ascii")

    def test_template_is_example_and_refused(self):
        with self.assertRaises(journal.JournalError):
            journal.load_package(ROOT / "日記範例" / "今日日記.json")

    def test_normalize_aliases_and_markdown_fence(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "day.json"
            path.write_text("說明\n```json\n" + json.dumps(sample_package(), ensure_ascii=False) + "\n```\n", encoding="utf-8")
            package = journal.normalize_package(journal.load_package(path))
        self.assertEqual(package["date"].isoformat(), "2026-11-04")
        self.assertEqual(package["climate"], "氣候D (冷卻重置期)")
        self.assertEqual(package["games"][0]["signal_type"], "低溫盲區受讓")
        self.assertIn("低溫盲區受讓", package["warnings"][0])
        self.assertEqual(package["games"][0]["outcome"], "全輸")
        self.assertEqual(package["games"][0]["cf_pct"], 0.55)
        self.assertEqual(package["games"][0]["stake_u"], 1.0)
        self.assertEqual(package["games"][1]["outcome"], "鎖倉避險")
        self.assertEqual(package["games"][1]["line_odds"], "—")
        self.assertEqual(package["games"][1]["stake_u"], 0.0)
        self.assertEqual(package["games"][1]["pick"], "【$0 風控鎖倉】")
        with self.assertRaises(journal.JournalError):
            journal.normalize_package({"date": "2026-11-04", "climate": "氣候A (冷卻重置期)", "pre_game_text": "a", "post_game_text": "b", "games": []})
        with self.assertRaises(journal.JournalError):
            journal.normalize_package({
                "date": "2026-11-04",
                "climate": "氣候A (放養紅利期)",
                "pre_game_text": "a",
                "post_game_text": "b",
                "games": [{
                    "game_id": "G1", "matchup": "客 @ 主", "climate": "氣候A (放養紅利期)",
                    "pick": "客 +1", "line_odds": "C盤 0.900", "cf_pct": 0.4,
                    "signal_type": "低溫盲區受讓", "stake_u": 2, "score": "", "margin": "主+7",
                    "outcome": "全贏", "review_notes": "",
                }],
            })

    def test_prefers_downloaded_copy_name(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            plain = root / "NBA26-27_Quant_Betting_Journal.xlsx"
            named = root / "Downloads" / "NBA26-27_Quant_Betting_Journal (1).xlsx"
            named.parent.mkdir()
            plain.write_bytes(b"a")
            named.write_bytes(b"b")
            chosen, others = journal.choose_workbook(roots=[root, named.parent])
            self.assertEqual(chosen, named)
            self.assertEqual(others, [plain])

    def test_path_file_and_mlb_guard(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            target = root / "NBA26-27_Quant_Betting_Journal (1).xlsx"
            target.write_bytes(b"a")
            override = root / "日記路徑.txt"
            override.write_text(f"# 註解\n{target}\n", encoding="utf-8")
            chosen, _others = journal.choose_workbook(roots=[], override_file=override)
            self.assertEqual(chosen, target)
            mlb = root / "MLB.xlsx"
            mlb.write_bytes(b"a")
            with self.assertRaises(journal.JournalError):
                journal.choose_workbook(explicit=mlb)

    def test_append_one_day_keeps_existing_rows(self):
        before = file_hash(SOURCE)
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            workbook = root / "NBA26-27_Quant_Betting_Journal (1).xlsx"
            shutil.copy2(SOURCE, workbook)
            payload = root / "今日日記.json"
            payload.write_text(json.dumps(sample_package(), ensure_ascii=False), encoding="utf-8")
            original = load_workbook(workbook)
            dashboard = [(cell.value) for row in original["1_Dashboard"].iter_rows() for cell in row]
            master_tail = [original["2_Master_Log"].cell(78, column).value for column in range(1, 15)]
            journal_banner = original["4_Daily_Journal"]["B48"].value
            original.close()

            text = journal.run(payload, workbook, dry_run=True, roots=[root])
            self.assertIn("還沒寫入", text)
            self.assertEqual(file_hash(workbook), file_hash(SOURCE))

            journal.run(payload, workbook, dry_run=False, roots=[root])
            saved = load_workbook(workbook)
            master = saved["2_Master_Log"]
            equity = saved["3_Equity_Timeline"]
            cards = saved["4_Daily_Journal"]
            dashboard_after = [(cell.value) for row in saved["1_Dashboard"].iter_rows() for cell in row]

            self.assertEqual(dashboard_after, dashboard)
            self.assertEqual([master.cell(78, column).value for column in range(1, 15)], master_tail)
            self.assertEqual(cards["B48"].value, journal_banner)
            self.assertEqual(journal.cell_date(master["A79"].value).isoformat(), "2026-11-04")
            self.assertEqual(master["B79"].value, "G1")
            self.assertEqual(master["D79"].value, "氣候D (冷卻重置期)")
            self.assertEqual(master["H79"].value, "低溫盲區受讓")
            self.assertEqual(master["E80"].value, "【$0 風控鎖倉】")
            self.assertEqual(master["F79"].value, "B盤 0.914")
            self.assertEqual(master["G79"].value, 0.55)
            self.assertEqual(master["L79"].value, "全輸")
            self.assertEqual(master["M79"].value, journal.pl_formula(79))
            self.assertEqual(master["A79"].number_format, "yyyy-mm-dd")
            self.assertIsNone(master["A81"].value)
            self.assertEqual(master["B80"].value, "G2")
            self.assertEqual(master["L80"].value, "鎖倉避險")
            self.assertEqual(master.auto_filter.ref, "A2:N80")

            self.assertEqual(journal.cell_date(equity["A12"].value).isoformat(), "2026-11-03")
            self.assertEqual(journal.cell_date(equity["A13"].value).isoformat(), "2026-11-04")
            self.assertEqual(equity["B13"].value, "氣候D (冷卻重置期)")
            self.assertEqual(equity["G13"].value, "=SUM(F$3:F13)")
            self.assertIn("TEXT(A13,\"yyyy-mm-dd\")", equity["I13"].value)
            self.assertIn("4_Daily_Journal", equity["I13"].value)
            chart = equity._charts[0]
            self.assertEqual(chart.series[0].val.numRef.f, "'3_Equity_Timeline'!$G$3:$G$13")
            self.assertEqual(chart.series[0].cat.numRef.f, "'3_Equity_Timeline'!$A$3:$A$13")
            self.assertEqual(chart.anchor._from.row, 14)

            self.assertEqual(journal.cell_date(cards["A48"].value).isoformat(), "2026-11-03")
            self.assertEqual(journal.cell_date(cards["A53"].value).isoformat(), "2026-11-04")
            self.assertEqual(journal.cell_date(cards["A58"].value).isoformat(), "2026-11-04")
            self.assertEqual(cards["B53"].value, "2026-11-04 | 氣候D (冷卻重置期)")
            self.assertEqual(cards["B54"].value, "【賽前量化分析與對話推論】")
            self.assertIn("賽前分析", cards["B55"].value)
            self.assertEqual(cards["B56"].value, "【賽後對戰結算與復盤日誌】")
            self.assertIn("賽後覆盤", cards["B57"].value)
            self.assertIn("3_Equity_Timeline", cards["B58"].value)
            self.assertIn("B53:I53", {str(merged) for merged in cards.merged_cells.ranges})
            saved.close()

            with zipfile.ZipFile(workbook) as book:
                names = book.namelist()
                xml = b"".join(book.read(name) for name in names if name.endswith(".xml"))
            self.assertNotIn(b"YYYY", xml)
            self.assertFalse(any("vml" in name.lower() for name in names))

            refreshed = sample_package()
            refreshed["post_game_text"] = "【2026-11-04 賽後覆盤】\n賽後已更新。"
            refreshed["games"][0]["outcome"] = "全贏"
            refreshed["games"][0]["score"] = "110 : 100"
            refreshed["games"][0]["signal_type"] = "低溫盲區受讓"
            payload.write_text(json.dumps(refreshed, ensure_ascii=False), encoding="utf-8")
            journal.main([str(payload), str(workbook)])
            again = load_workbook(workbook)
            self.assertIsNone(again["2_Master_Log"]["A81"].value)
            self.assertEqual(again["2_Master_Log"]["L79"].value, "全贏")
            self.assertEqual(again["2_Master_Log"]["J79"].value, "110 : 100")
            self.assertEqual(journal.cell_date(again["3_Equity_Timeline"]["A13"].value).isoformat(), "2026-11-04")
            self.assertIsNone(again["3_Equity_Timeline"]["A14"].value)
            self.assertIn("賽後已更新", again["4_Daily_Journal"]["B57"].value)
            again.close()
            short = sample_package()
            short["games"] = short["games"][:1]
            payload.write_text(json.dumps(short, ensure_ascii=False), encoding="utf-8")
            with self.assertRaises(SystemExit):
                journal.main([str(payload), str(workbook)])

            os.chmod(workbook, 0o444)
            try:
                with self.assertRaises(journal.JournalError):
                    journal.run(payload, workbook, dry_run=False, roots=[root])
            finally:
                os.chmod(workbook, 0o644)
        self.assertEqual(file_hash(SOURCE), before)


if __name__ == "__main__":
    unittest.main()
