# -*- coding: utf-8 -*-
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
PICK = ROOT / "jiujiu_pick" / "nba_pick.py"
APPLY = ROOT / "jiujiu_pick" / "apply_on_vps.py"


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


pick = load("nba_pick_module", PICK)
apply = load("apply_nba_pick", APPLY)


class GradeTests(unittest.TestCase):
    def test_selftest_examples(self):
        pick.selftest()

    def test_cover_receive_and_half_point(self):
        self.assertEqual(pick.grade_pick("老鷹受讓2分", "老鷹", "活塞", -1)["d"], "預測成功")
        self.assertEqual(pick.grade_pick("活塞讓2分", "老鷹", "活塞", -1)["c"], "活塞贏1分")
        self.assertEqual(pick.grade_pick("活塞讓2.5分", "老鷹", "活塞", -2)["d"], "預測失敗")
        self.assertEqual(pick.grade_pick("活塞讓2.5分", "老鷹", "活塞", -3)["d"], "預測成功")

    def test_display_token_follows_header_when_unsigned(self):
        self.assertEqual(pick.signed_from_display("(7+50)0.95", "老鷹", "老鷹受讓"), -7.5)
        self.assertEqual(pick.signed_from_display("(14.5)0.9", "湖人", "湖人讓分"), 14.5)
        self.assertEqual(pick.parse_spread_token("(-3.5)0.9"), -3.5)

    def test_base_pick_strips_result_mark(self):
        self.assertEqual(pick.base_pick("活塞讓2分輸"), "活塞讓2分")
        self.assertEqual(pick.base_pick("隨便寫"), "")


def build_game(ws):
    ws["A1"] = "11-23 04:30 活塞 @ 老鷹(主)"
    ws["A2"] = "時間"
    ws["B2"] = "老鷹（主）"
    ws["C2"] = "活塞"
    ws["D2"] = "老鷹受讓"
    ws["E2"] = "活塞讓分"
    ws.cell(2, 42, "讓分A")
    ws["A3"] = "07:00"
    ws["D3"] = "(2)0.95"
    ws.cell(3, 42, -2)


class SheetTests(unittest.TestCase):
    def test_menu_before_game_and_grade_after(self):
        wb = Workbook()
        ws = wb.active
        build_game(ws)
        notes = {"11-23 04:30 活塞 @ 老鷹(主)": {"賽前預測": "活塞讓2分", "比賽結果": ""}}
        pick.render_pick(ws, "11-23 04:30 活塞 @ 老鷹(主)", 4, notes, 12)
        self.assertEqual(ws["A4"].value, "賽前預測")
        self.assertEqual(ws["B4"].value, "活塞讓2分")
        self.assertIsNone(ws["D4"].value)
        self.assertIn("活塞讓2分", ws.data_validations.dataValidation[0].formula1)
        self.assertIn("老鷹受讓2分", ws.data_validations.dataValidation[0].formula1)

        ws["A5"] = "比賽結果"
        ws["B5"] = "客116-主111 客隊贏5"
        pick.render_pick(ws, "11-23 04:30 活塞 @ 老鷹(主)", 4, notes, 12)
        self.assertEqual(ws["B4"].value, "活塞讓2分贏")
        self.assertEqual(ws["C4"].value, "活塞贏5分")
        self.assertEqual(ws["D4"].value, "預測成功")
        self.assertTrue(str(ws["D4"].fill.fgColor.rgb).endswith("C6EFCE"))

    def test_failure_is_red_and_saved_line_does_not_follow_a_later_number(self):
        wb = Workbook()
        ws = wb.active
        build_game(ws)
        ws.cell(3, 42, -6)
        title = "11-23 04:30 活塞 @ 老鷹(主)"
        notes = {title: {"賽前預測": "活塞讓2分"}}
        ws["A5"] = "比賽結果"
        ws["B5"] = "活塞101:老鷹100"
        pick.render_pick(ws, title, 4, notes, 12)
        self.assertEqual(ws["B4"].value, "活塞讓2分輸")
        self.assertEqual(ws["C4"].value, "活塞贏1分")
        self.assertEqual(ws["D4"].value, "預測失敗")
        self.assertTrue(str(ws["D4"].fill.fgColor.rgb).endswith("FFC7CE"))
        self.assertEqual(notes[title]["賽前預測"], "活塞讓2分")

    def test_empty_pick_stays_blank_after_the_score(self):
        wb = Workbook()
        ws = wb.active
        build_game(ws)
        title = "11-23 04:30 活塞 @ 老鷹(主)"
        ws["A5"] = "比賽結果"
        ws["B5"] = "客116-主111"
        pick.render_pick(ws, title, 4, {title: {}}, 12)
        self.assertIsNone(ws["B4"].value)
        self.assertIsNone(ws["C4"].value)
        self.assertIsNone(ws["D4"].value)

    def test_merge_and_archive_keep_the_pick(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "時段觀察"
        build_game(ws)
        ws["A4"] = "賽前預測"
        ws["B4"] = "活塞讓2分贏"
        notes = {}
        pick.merge_picks(wb, notes)
        title = "11-23 04:30 活塞 @ 老鷹(主)"
        self.assertEqual(notes[title]["賽前預測"], "活塞讓2分")

        archive = wb.create_sheet("賽事筆記")
        archive["A1"] = "場次"
        archive["B1"] = "賽前想法"
        archive["C1"] = "比賽結果"
        archive["D1"] = "賽後想法"
        pick.archive_picks(wb, notes)
        self.assertEqual(archive["E1"].value, "賽前預測")
        self.assertEqual(archive["A2"].value, title)
        self.assertEqual(archive["E2"].value, "活塞讓2分")

        book = Workbook()
        book.active.title = "時段觀察"
        sheet = book.create_sheet("賽事筆記")
        sheet["A1"] = "場次"
        sheet["B1"] = "賽前想法"
        sheet["E1"] = "賽前預測"
        sheet["A2"] = title
        sheet["E2"] = "活塞讓2分"
        restored = {}
        pick.merge_picks(book, restored)
        self.assertEqual(restored[title]["賽前預測"], "活塞讓2分")


FAKE_NOTES = '''# -*- coding: utf-8 -*-
def merged_game_notes(wb):
    return {}

def write_game_notes(ws, title, row, game_notes, last_col):
    ws.cell(row, 1, "賽前想法")
    ws.cell(row + 1, 1, "比賽結果")
    ws.cell(row + 1, 2, "客116-主111 客隊贏5")
    ws.cell(row + 2, 1, "賽後想法")
    return row + 3

def write_game_note_archive(wb, game_notes):
    ws = wb.create_sheet("賽事筆記")
    ws["A1"] = "場次"
    ws["B1"] = "賽前想法"
    ws["A2"] = "11-23 04:30 活塞 @ 老鷹(主)"
'''


class InstallTests(unittest.TestCase):
    def test_wrapper_puts_the_menu_above_the_notes(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "nba_game_notes.py").write_text(FAKE_NOTES, encoding="utf-8")
            apply.install(root, PICK)
            sys.path.insert(0, str(root))
            try:
                notes = load("nba_game_notes_wrapped", root / "nba_game_notes.py")
                wb = Workbook()
                ws = wb.active
                build_game(ws)
                title = "11-23 04:30 活塞 @ 老鷹(主)"
                game_notes = {title: {"賽前預測": "活塞讓2分"}}
                notes.write_game_notes(ws, title, 4, game_notes, 12)
                self.assertEqual(ws["A4"].value, "賽前預測")
                self.assertEqual(ws["A5"].value, "賽前想法")
                self.assertEqual(ws["A6"].value, "比賽結果")
                self.assertEqual(ws["B4"].value, "活塞讓2分贏")
                self.assertEqual(ws["D4"].value, "預測成功")
                notes.write_game_note_archive(wb, game_notes)
                archive = wb["賽事筆記"]
                self.assertEqual(archive["C1"].value, "賽前預測")
                self.assertEqual(archive["C2"].value, "活塞讓2分")
                apply.install(root, PICK)
                text = (root / "nba_game_notes.py").read_text(encoding="utf-8")
                self.assertEqual(text.count(apply.MARKER), 1)
            finally:
                sys.path.remove(str(root))
                sys.modules.pop("nba_pick", None)


if __name__ == "__main__":
    unittest.main()
