# -*- coding: utf-8 -*-
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Border, Side

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
        self.assertEqual(pick.grade_pick("老鷹受讓2分", "老鷹", "活塞", -1)["d"], "預測正確")
        self.assertEqual(pick.grade_pick("活塞讓2分", "老鷹", "活塞", -1)["c"], "活塞贏1分")
        self.assertEqual(pick.grade_pick("活塞讓2.5分", "老鷹", "活塞", -2)["d"], "預測失敗")
        self.assertEqual(pick.grade_pick("活塞讓2.5分", "老鷹", "活塞", -3)["d"], "預測正確")

    def test_display_token_follows_header_when_unsigned(self):
        self.assertEqual(pick.signed_from_display("(7+50)0.95", "老鷹", "老鷹受讓"), -7)
        self.assertEqual(pick.signed_from_display("(2-25)0.95", "老鷹", "老鷹受讓"), -2)
        self.assertEqual(pick.signed_from_display("(2.5)0.95", "湖人", "湖人讓分"), 2.5)
        self.assertEqual(pick.parse_spread_token("0.95"), None)
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
    def test_phone_sheet_reads_the_spread_not_the_price(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "11-23 04:30 底特律活塞 @ 亞特蘭大老鷹(主)"
        ws["A2"] = "時間"
        ws["B2"] = "老鷹（主）"
        ws["C2"] = "活塞"
        ws["D2"] = "老鷹受讓"
        ws["E2"] = "活塞讓分"
        ws["A3"] = "11-23 04:02"
        ws["D3"] = 0.95
        ws["E3"] = "(2-25)0.95"
        title = ws["A1"].value
        pick.render_pick(ws, title, 4, {title: {}}, 12)
        self.assertEqual(ws["A4"].value, "賽前預測")
        self.assertEqual(ws["B4"].value, "AI推薦")
        self.assertEqual(ws["E4"].value, "我的推薦")
        formula = ws.data_validations.dataValidation[0].formula1
        self.assertIn("活塞讓2分-25", formula)
        self.assertIn("老鷹受讓2分贏25", formula)
        self.assertNotIn("0.95", formula)

    def test_two_picks_grade_each_side(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "11-26 09:30 快艇 @ 灰熊(主)"
        ws["A2"] = "時間"
        ws["B2"] = "灰熊（主）"
        ws["C2"] = "快艇"
        ws["D2"] = "灰熊讓分"
        ws["A3"] = "11-26 07:02"
        ws["D3"] = "(7-25)0.95"
        title = ws["A1"].value
        notes = {title: {"AI推薦": "快艇受讓7分贏25", "我的推薦": "灰熊讓7分-25"}}
        ws["A6"] = "比賽結果"
        ws["B6"] = "（客）快艇97/（主）灰熊115，（主）灰熊贏18分，讓分過盤"
        pick.render_pick(ws, title, 4, notes, 12)
        self.assertEqual(ws["B5"].value, "快艇受讓7分贏25")
        self.assertEqual(ws["C5"].value, "快艇輸18分")
        self.assertEqual(ws["D5"].value, "預測失敗")
        self.assertTrue(str(ws["D5"].fill.fgColor.rgb).endswith("FFC7CE"))
        self.assertEqual(ws["E5"].value, "灰熊讓7分-25")
        self.assertEqual(ws["F5"].value, "灰熊贏18分")
        self.assertEqual(ws["G5"].value, "預測正確")
        self.assertTrue(str(ws["G5"].fill.fgColor.rgb).endswith("C6EFCE"))

    def test_menu_before_game_and_grade_after(self):
        wb = Workbook()
        ws = wb.active
        build_game(ws)
        notes = {"11-23 04:30 活塞 @ 老鷹(主)": {"賽前預測": "活塞讓2分", "比賽結果": ""}}
        pick.render_pick(ws, "11-23 04:30 活塞 @ 老鷹(主)", 4, notes, 12)
        self.assertEqual(ws["A4"].value, "賽前預測")
        self.assertEqual(ws["E5"].value, "活塞讓2分")
        self.assertIsNone(ws["G5"].value)
        self.assertIn("活塞讓2分", ws.data_validations.dataValidation[0].formula1)
        self.assertIn("老鷹受讓2分", ws.data_validations.dataValidation[0].formula1)

        ws["A6"] = "比賽結果"
        ws["B6"] = "客116-主111 客隊贏5"
        pick.render_pick(ws, "11-23 04:30 活塞 @ 老鷹(主)", 4, notes, 12)
        self.assertEqual(ws["E5"].value, "活塞讓2分")
        self.assertEqual(ws["F5"].value, "活塞贏5分")
        self.assertEqual(ws["G5"].value, "預測正確")
        self.assertTrue(str(ws["G5"].fill.fgColor.rgb).endswith("C6EFCE"))

    def test_failure_is_red_and_saved_line_does_not_follow_a_later_number(self):
        wb = Workbook()
        ws = wb.active
        build_game(ws)
        ws.cell(3, 42, -6)
        title = "11-23 04:30 活塞 @ 老鷹(主)"
        notes = {title: {"賽前預測": "活塞讓2分"}}
        ws["A6"] = "比賽結果"
        ws["B6"] = "活塞101:老鷹100"
        pick.render_pick(ws, title, 4, notes, 12)
        self.assertEqual(ws["E5"].value, "活塞讓2分")
        self.assertEqual(ws["F5"].value, "活塞贏1分")
        self.assertEqual(ws["G5"].value, "預測失敗")
        self.assertTrue(str(ws["G5"].fill.fgColor.rgb).endswith("FFC7CE"))
        self.assertEqual(notes[title]["我的推薦"], "活塞讓2分")

    def test_empty_pick_stays_blank_after_the_score(self):
        wb = Workbook()
        ws = wb.active
        build_game(ws)
        title = "11-23 04:30 活塞 @ 老鷹(主)"
        ws["A6"] = "比賽結果"
        ws["B6"] = "客116-主111"
        pick.render_pick(ws, title, 4, {title: {}}, 12)
        self.assertIsNone(ws["B5"].value)
        self.assertIsNone(ws["E5"].value)
        self.assertIsNone(ws["G5"].value)

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
        self.assertEqual(archive["F1"].value, "我的推薦")
        self.assertEqual(archive["A2"].value, title)
        self.assertEqual(archive["F2"].value, "活塞讓2分")

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

    def test_archive_fits_note_text_and_draws_a_day_line(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "賽事筆記"
        ws["A1"] = "場次"
        ws["B1"] = "賽前想法"
        ws["C1"] = "比賽結果"
        ws.column_dimensions["B"].width = 20
        ws.column_dimensions["C"].width = 20
        long_note = "老鷹主場讓分從6.5到8.75，覺得是順向盤，看好老鷹贏9分以上。"
        rows = [
            (2, "11-20 09:00 勇士 @ 灰熊(主)", "短", "灰熊贏"),
            (3, "11-21 08:00 黃蜂 @ 老鷹(主)", long_note * 3, "黃蜂97-老鷹111，老鷹贏14分"),
            (4, "11-21 08:30 溜馬 @ 騎士(主)", "短", "騎士贏"),
        ]
        for row, title, pregame, result in rows:
            ws.cell(row, 1, title)
            ws.cell(row, 2, pregame)
            ws.cell(row, 3, result)
        ws.row_dimensions[3].height = 15
        for col in range(1, 4):
            ws.cell(3, col).border = Border(bottom=Side(style="thick", color="000000"))

        pick.archive_picks(wb, {})

        self.assertTrue(ws["B3"].alignment.wrap_text)
        self.assertEqual(ws["B3"].alignment.vertical, "top")
        self.assertTrue(ws["C4"].alignment.wrap_text)
        self.assertGreater(ws.row_dimensions[3].height, ws.row_dimensions[2].height)
        self.assertGreater(ws.row_dimensions[3].height, 15)
        self.assertEqual(ws["A2"].border.bottom.style, "thick")
        self.assertEqual(ws["B2"].border.bottom.style, "thick")
        self.assertEqual(ws["C2"].border.bottom.style, "thick")
        self.assertNotEqual(ws["A3"].border.bottom.style, "thick")
        self.assertNotEqual(ws["C3"].border.bottom.style, "thick")
        self.assertEqual(ws["A4"].border.bottom.style, "thick")
        self.assertEqual(ws["C4"].border.bottom.style, "thick")
        self.assertEqual(ws.column_dimensions["B"].width, 20)


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
                self.assertEqual(ws["B4"].value, "AI推薦")
                self.assertEqual(ws["A6"].value, "賽前想法")
                self.assertEqual(ws["A7"].value, "比賽結果")
                self.assertEqual(ws["E5"].value, "活塞讓2分")
                self.assertEqual(ws["G5"].value, "預測正確")
                notes.write_game_note_archive(wb, game_notes)
                archive = wb["賽事筆記"]
                self.assertEqual(archive["C1"].value, "AI推薦")
                self.assertEqual(archive["D1"].value, "我的推薦")
                self.assertEqual(archive["D2"].value, "活塞讓2分")
                apply.install(root, PICK)
                text = (root / "nba_game_notes.py").read_text(encoding="utf-8")
                self.assertEqual(text.count(apply.MARKER), 1)
            finally:
                sys.path.remove(str(root))
                sys.modules.pop("nba_pick", None)

    def test_wrapper_accepts_row_before_title(self):
        notes_text = FAKE_NOTES.replace(
            "def write_game_notes(ws, title, row, game_notes, last_col):",
            "def write_game_notes(ws, row, title, game_notes, last_col):",
        )
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "nba_game_notes.py").write_text(notes_text, encoding="utf-8")
            apply.install(root, PICK)
            sys.path.insert(0, str(root))
            try:
                notes = load("nba_game_notes_row_first", root / "nba_game_notes.py")
                wb = Workbook()
                ws = wb.active
                build_game(ws)
                title = "11-23 04:30 活塞 @ 老鷹(主)"
                game_notes = {title: {"賽前預測": "活塞讓2分"}}
                notes.write_game_notes(ws, 4, title, game_notes, 12)
                self.assertEqual(ws["A4"].value, "賽前預測")
                self.assertEqual(ws["A6"].value, "賽前想法")
                self.assertEqual(ws["E5"].value, "活塞讓2分")
                self.assertEqual(ws["G5"].value, "預測正確")
            finally:
                sys.path.remove(str(root))
                sys.modules.pop("nba_pick", None)


if __name__ == "__main__":
    unittest.main()
