# -*- coding: utf-8 -*-
"""產出 NBA26-27 量化操盤日記。

第 2 頁的比分是公開賽果。氣候、訊號、水位、注碼和第 4 頁文字是示範，
用來讓儀表板、資金曲線和日期連結一打開就能用。正式記帳請改第 2 頁和第 4 頁。
"""

from __future__ import annotations

from datetime import datetime

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "/agent/repos/extensions/日記範例/NBA26-27_Quant_Betting_Journal.xlsx"

NAVY = "1F4E78"
SECTION = "2F5496"
ICE = "DCE6F1"
ICE_ALT = "EDF3F9"
GREEN = "E2EFDA"
GREEN_ALT = "F4FBF2"
TOTAL = "D6DCE4"
ZEBRA = "F9FAFB"
WHITE = "FFFFFF"
ORANGE = "C65911"
LINE = "B0B7C3"
WIN = "C6EFCE"
LOSS = "F8CBAD"

FONT = "Calibri"
THIN = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
fill_navy = PatternFill("solid", fgColor=NAVY)
fill_section = PatternFill("solid", fgColor=SECTION)
fill_ice = PatternFill("solid", fgColor=ICE)
fill_ice_alt = PatternFill("solid", fgColor=ICE_ALT)
fill_green = PatternFill("solid", fgColor=GREEN)
fill_green_alt = PatternFill("solid", fgColor=GREEN_ALT)
fill_total = PatternFill("solid", fgColor=TOTAL)
fill_white = PatternFill("solid", fgColor=WHITE)
fill_zebra = PatternFill("solid", fgColor=ZEBRA)

font_title = Font(name=FONT, size=18, bold=True, color=NAVY)
font_note = Font(name=FONT, size=11, italic=True, color=ORANGE)
font_head = Font(name=FONT, size=12, bold=True, color=WHITE)
font_section = Font(name=FONT, size=12, bold=True, color=WHITE)
font_label = Font(name=FONT, size=11, bold=True, color="1A1A1A")
font_body = Font(name=FONT, size=11, color="1A1A1A")
font_kpi = Font(name=FONT, size=20, bold=True, color=NAVY)
font_small = Font(name=FONT, size=9, color="667085")

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
right_align = Alignment(horizontal="right", vertical="center")

LAST_DATA_ROW = 200
CLIMATES = ("氣候A-冷卻期", "氣候B-誘盤期", "氣候C-大收割期")
SIGNALS = ("RLM逆向推盤", "低溫盲區受讓", "KeyPoint鎖水", "高溫熔斷", "風控鎖倉")
OUTCOMES = ("全贏", "全輸", "平手", "鎖倉避險")

# 比分是公開賽果，寫成客分-主分。讓分差是推薦標的得分減對手。
# 五張點名贏票：籃網+4.5、熱火-9、鵜鶘+6.5、魔術+6.5、雷霆+4.5。
GAMES = [
    ("2026-10-27", "G1", "太陽客 @ 雷霆主", "氣候A-冷卻期", "鎖倉不進場", "—", None, "風控鎖倉", 0.0, "114-125", None, "鎖倉避險", "雷霆贏 11 分。這列示範鎖倉，注碼 0。"),
    ("2026-10-27", "G2", "公牛客 @ 賽爾提克主", "氣候B-誘盤期", "賽爾提克 -12.5", "A盤 0.900", 0.71, "高溫熔斷", 0.5, "99-134", 35, "全贏", "示範列。公開比分賽爾提克贏 35 分，深讓仍過盤。"),
    ("2026-10-28", "G1", "籃網客 @ 賽爾提克主", "氣候A-冷卻期", "籃網 +4.5", "C盤 0.900", 0.35, "低溫盲區受讓", 1.0, "124-125", -1, "全贏", "點名贏票。加時只輸 1 分，受讓 4.5 過盤。"),
    ("2026-10-28", "G2", "拓荒者客 @ 湖人主", "氣候B-誘盤期", "拓荒者 -7", "B盤 0.860", 0.52, "KeyPoint鎖水", 0.5, "136-129", 7, "平手", "公開比分剛好贏 7 分，示範平手，損益 0。"),
    ("2026-10-29", "G1", "熱火客 @ 老鷹主", "氣候C-大收割期", "熱火 -9", "C盤 0.844", 0.62, "RLM逆向推盤", 1.0, "129-117", 12, "全贏", "點名贏票。熱火贏 12 分，讓 9 分過盤。"),
    ("2026-10-29", "G2", "雷霆客 @ 獨行俠主", "氣候B-誘盤期", "雷霆 +8.5", "A盤 0.820", 0.41, "高溫熔斷", 0.5, "96-120", -24, "全輸", "示範敗票。這張不是雷霆 +4.5。雷霆輸 24 分。"),
    ("2026-10-30", "G1", "國王客 @ 鵜鶘主", "氣候A-冷卻期", "鵜鶘 +6.5", "C盤 0.880", 0.42, "低溫盲區受讓", 1.0, "120-132", 12, "全贏", "點名贏票。鵜鶘贏 12 分，受讓 6.5 過盤。"),
    ("2026-10-30", "G2", "騎士客 @ 老鷹主", "氣候B-誘盤期", "騎士 +2.5", "B盤 0.910", 0.47, "RLM逆向推盤", 0.5, "111-114", -3, "全輸", "示範敗票。騎士輸 3 分，受讓 2.5 沒過。"),
    ("2026-10-31", "G1", "魔術客 @ 暴龍主", "氣候C-大收割期", "魔術 +6.5", "C盤 0.920", 0.55, "KeyPoint鎖水", 1.0, "121-114", 7, "全贏", "點名贏票。魔術贏 7 分，受讓 6.5 過盤。"),
    ("2026-10-31", "G2", "活塞客 @ 籃網主", "氣候B-誘盤期", "籃網 +10", "A盤 0.800", 0.33, "高溫熔斷", 0.5, "110-94", -16, "全輸", "示範敗票。這張不是籃網 +4.5。籃網輸 16 分。"),
    ("2026-10-31", "G3", "溜馬客 @ 熱火主", "氣候C-大收割期", "鎖倉不進場", "—", None, "風控鎖倉", 0.0, "125-145", None, "鎖倉避險", "熱火贏 20 分。這列示範不進場。"),
    ("2026-11-01", "G1", "火箭客 @ 雷霆主", "氣候B-誘盤期", "雷霆 -7", "B盤 0.870", 0.68, "RLM逆向推盤", 1.0, "119-98", -21, "全輸", "示範敗票。雷霆輸 21 分。贏票在 11 月 3 日那張 +4.5。"),
    ("2026-11-01", "G2", "老鷹客 @ 鵜鶘主", "氣候A-冷卻期", "鎖倉不進場", "—", None, "風控鎖倉", 0.0, "99-121", None, "鎖倉避險", "鵜鶘贏 22 分。這列示範鎖倉。"),
    ("2026-11-01", "G3", "公鹿客 @ 騎士主", "氣候B-誘盤期", "公鹿 +9.5", "A盤 0.850", 0.38, "低溫盲區受讓", 0.5, "111-136", -25, "全輸", "示範敗票。公鹿輸 25 分，受讓 9.5 沒過。"),
    ("2026-11-02", "G1", "塞爾提克客 @ 魔術主", "氣候A-冷卻期", "魔術 +3", "C盤 0.900", 0.44, "低溫盲區受讓", 0.5, "108-97", -11, "全輸", "示範敗票。這張不是魔術 +6.5。魔術輸 11 分。"),
    ("2026-11-02", "G2", "溜馬客 @ 籃網主", "氣候C-大收割期", "籃網 -4", "B盤 0.860", 0.58, "RLM逆向推盤", 0.5, "90-98", 8, "全贏", "示範贏票。籃網贏 8 分，讓 4 分過盤。"),
    ("2026-11-02", "G3", "勇士客 @ 快艇主", "氣候B-誘盤期", "鎖倉不進場", "—", None, "風控鎖倉", 0.0, "124-112", None, "鎖倉避險", "勇士贏 12 分。這列示範不進場。"),
    ("2026-11-03", "G1", "雷霆客 @ 騎士主", "氣候C-大收割期", "雷霆 +4.5", "C盤 0.950", 0.48, "RLM逆向推盤", 1.0, "117-112", 5, "全贏", "點名贏票。雷霆贏 5 分，受讓 4.5 過盤。"),
    ("2026-11-03", "G2", "鵜鶘客 @ 灰熊主", "氣候A-冷卻期", "鵜鶘 -3.5", "B盤 0.800", 0.51, "KeyPoint鎖水", 0.5, "135-134", 1, "全輸", "加時只贏 1 分，讓 3.5 沒過。這張不是鵜鶘 +6.5。"),
    ("2026-11-03", "G3", "熱火客 @ 湖人主", "氣候B-誘盤期", "熱火 +6", "A盤 0.830", 0.36, "高溫熔斷", 0.5, "117-138", -21, "全輸", "示範敗票。熱火輸 21 分。"),
    ("2026-11-03", "G4", "籃網客 @ 尼克主", "氣候C-大收割期", "鎖倉不進場", "—", None, "風控鎖倉", 0.0, "90-103", None, "鎖倉避險", "籃網輸 13 分。這列示範鎖倉。"),
]


def apply_grid(ws):
    ws.sheet_view.showGridLines = True
    ws.sheet_view.zoomScale = 110
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.horizontalCentered = True
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.4
    ws.sheet_properties.tabColor = NAVY


def put(cell, value, font=font_body, fill=None, align=center, fmt=None):
    cell.value = value
    cell.font = font
    cell.alignment = align
    cell.border = THIN
    if fill is not None:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt


def widths(ws, pairs):
    for letter, width in pairs.items():
        ws.column_dimensions[letter].width = width


def pl_formula(row: int) -> str:
    return (
        f'=IF(A{row}="","",'
        f'IF(OR(L{row}="平手",L{row}="鎖倉避險"),0,'
        f'IF(L{row}="全輸",-N(I{row}),'
        f'IF(L{row}="全贏",N(I{row})*IFERROR(VALUE(TRIM(RIGHT(SUBSTITUTE(F{row}," ",REPT(" ",99)),99))),0),'
        f'""))))'
    )


def build_master(wb: Workbook):
    ws = wb.active
    ws.title = "2_Master_Log"
    apply_grid(ws)
    headers = [
        "日期", "場次", "對戰組合", "大盤氣候", "推薦標的", "盤口規格與水位",
        "尾盤金流CF%", "訊號類型", "建議注碼U", "終場比分", "讓分差",
        "盤口結果", "單場損益U", "賽後邏輯與偏差檢討",
    ]
    ws.merge_cells("A1:N1")
    put(
        ws["A1"],
        "賽前推薦與賽後對帳流水帳　　淡藍是輸入，淡綠是公式。改結果或注碼後，損益會重算。",
        font_title,
        align=Alignment(horizontal="left", vertical="center"),
    )
    ws["A1"].border = Border()
    ws.row_dimensions[1].height = 28
    for col, header in enumerate(headers, start=1):
        put(ws.cell(2, col), header, font_head, fill_navy, center)
    ws.row_dimensions[2].height = 30
    ws.auto_filter.ref = f"A2:N{LAST_DATA_ROW}"
    ws.freeze_panes = "B3"
    ws.auto_filter.ref = "A2:N21"
    ws.print_title_rows = "1:2"

    for index, game in enumerate(GAMES):
        row = 3 + index
        alt = index % 2 == 1
        ice = fill_ice_alt if alt else fill_ice
        green = fill_green_alt if alt else fill_green
        day, gid, matchup, climate, pick, line, cf, signal, stake, score, margin, outcome, note = game
        values = [
            datetime.strptime(day, "%Y-%m-%d").date(),
            gid, matchup, climate, pick, line, cf, signal, stake, score, margin, outcome, None, note,
        ]
        for col, value in enumerate(values, start=1):
            fill = green if col == 13 else ice
            align = left_wrap if col in (3, 5, 6, 14) else center
            put(ws.cell(row, col), value, font_body, fill, align)
        ws.cell(row, 1).number_format = "YYYY-MM-DD"
        ws.cell(row, 7).number_format = "0.0%"
        ws.cell(row, 9).number_format = '0.0"U"'
        ws.cell(row, 11).number_format = "0.0"
        ws.cell(row, 13).value = pl_formula(row)
        ws.cell(row, 13).number_format = '+0.000"U";-0.000"U";0.000"U"'
        ws.cell(row, 13).font = Font(name=FONT, size=11, bold=True)
        ws.row_dimensions[row].height = 32

    for row in range(3 + len(GAMES), LAST_DATA_ROW + 1):
        for col in range(1, 15):
            fill = fill_green if col == 13 else fill_white
            put(ws.cell(row, col), None, font_body, fill, center)
        ws.cell(row, 1).number_format = "YYYY-MM-DD"
        ws.cell(row, 7).number_format = "0.0%"
        ws.cell(row, 9).number_format = '0.0"U"'
        ws.cell(row, 11).number_format = "0.0"
        ws.cell(row, 13).value = pl_formula(row)
        ws.cell(row, 13).number_format = '+0.000"U";-0.000"U";0.000"U"'
        ws.row_dimensions[row].height = 20

    ws.cell(2, 11).comment = Comment("推薦標的的得分減掉對手。正數代表推薦的那隊贏分。", "日記", width=240, height=40)
    ws.cell(2, 13).comment = Comment("全贏＝注碼乘最後一個水位數字。全輸＝減去注碼。平手與鎖倉＝0。", "日記", width=260, height=48)

    outcome_dv = DataValidation(type="list", formula1='"全贏,全輸,平手,鎖倉避險"', allow_blank=True)
    climate_dv = DataValidation(type="list", formula1='"氣候A-冷卻期,氣候B-誘盤期,氣候C-大收割期"', allow_blank=True)
    signal_dv = DataValidation(type="list", formula1='"RLM逆向推盤,低溫盲區受讓,KeyPoint鎖水,高溫熔斷,風控鎖倉"', allow_blank=True)
    stake_dv = DataValidation(type="list", formula1='"0,0.5,1"', allow_blank=True)
    for dv in (outcome_dv, climate_dv, signal_dv, stake_dv):
        dv.error = "請用清單裡的項目"
        dv.errorTitle = "無法輸入"
        dv.prompt = "請由清單選擇"
        dv.promptTitle = "輸入"
        ws.add_data_validation(dv)
    outcome_dv.add(f"L3:L{LAST_DATA_ROW}")
    climate_dv.add(f"D3:D{LAST_DATA_ROW}")
    signal_dv.add(f"H3:H{LAST_DATA_ROW}")
    stake_dv.add(f"I3:I{LAST_DATA_ROW}")

    ws.conditional_formatting.add(
        f"L3:L{LAST_DATA_ROW}",
        FormulaRule(formula=['L3="全贏"'], fill=PatternFill("solid", fgColor=WIN)),
    )
    ws.conditional_formatting.add(
        f"L3:L{LAST_DATA_ROW}",
        FormulaRule(formula=['L3="全輸"'], fill=PatternFill("solid", fgColor=LOSS)),
    )
    ws.conditional_formatting.add(
        f"M3:M{LAST_DATA_ROW}",
        CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=FONT, size=11, bold=True, color="006100")),
    )
    ws.conditional_formatting.add(
        f"M3:M{LAST_DATA_ROW}",
        CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONT, size=11, bold=True, color="9C0006")),
    )

    widths(ws, {
        "A": 14, "B": 10, "C": 28, "D": 18, "E": 18, "F": 18, "G": 16,
        "H": 18, "I": 14, "J": 14, "K": 12, "L": 14, "M": 16, "N": 42,
    })
    ws.auto_filter.ref = f"A2:N{2 + len(GAMES)}"
    return ws


def build_equity(wb: Workbook):
    ws = wb.create_sheet("3_Equity_Timeline")
    apply_grid(ws)
    ws.merge_cells("A1:I1")
    put(
        ws["A1"],
        "每日資金曲線與時間軸　　數字由第 2 頁加總。最右欄可點進第 4 頁當天的文字。",
        font_title,
        align=Alignment(horizontal="left", vertical="center"),
    )
    ws["A1"].border = Border()
    ws.row_dimensions[1].height = 28
    headers = ["日期", "當日氣候", "下注場數", "鎖倉場數", "當日投入U", "當日淨損益U", "累積損益U", "當日勝率", "詳細日誌"]
    for col, header in enumerate(headers, start=1):
        put(ws.cell(2, col), header, font_head, fill_navy, center)
    ws.row_dimensions[2].height = 30
    ws.freeze_panes = "B3"
    ws.print_title_rows = "1:2"

    days = []
    for game in GAMES:
        if game[0] not in days:
            days.append(game[0])
    for index, day in enumerate(days):
        row = 3 + index
        alt = index % 2 == 1
        green = fill_green_alt if alt else fill_green
        ice = fill_ice_alt if alt else fill_ice
        put(ws.cell(row, 1), datetime.strptime(day, "%Y-%m-%d").date(), font_body, ice, center, "YYYY-MM-DD")
        put(
            ws.cell(row, 2),
            f'=IF(A{row}="","",TEXTJOIN("、",TRUE,UNIQUE(FILTER(\'2_Master_Log\'!$D$3:$D$500,\'2_Master_Log\'!$A$3:$A$500=A{row}))))',
            font_body, green, left_wrap,
        )
        put(ws.cell(row, 3), f'=COUNTIFS(\'2_Master_Log\'!A:A,A{row},\'2_Master_Log\'!I:I,">0")', font_body, green, center, "0")
        put(ws.cell(row, 4), f'=COUNTIFS(\'2_Master_Log\'!A:A,A{row},\'2_Master_Log\'!I:I,0)', font_body, green, center, "0")
        put(ws.cell(row, 5), f'=SUMIF(\'2_Master_Log\'!A:A,A{row},\'2_Master_Log\'!I:I)', font_body, green, center, '0.0"U"')
        put(ws.cell(row, 6), f'=SUMIF(\'2_Master_Log\'!A:A,A{row},\'2_Master_Log\'!M:M)', font_body, green, center, '+0.000"U";-0.000"U";0.000"U"')
        put(ws.cell(row, 7), f'=SUM(F$3:F{row})', font_label, green, center, '+0.000"U";-0.000"U";0.000"U"')
        put(ws.cell(row, 8), f'=IFERROR(COUNTIFS(\'2_Master_Log\'!A:A,A{row},\'2_Master_Log\'!L:L,"全贏")/C{row},0)', font_body, green, center, "0.0%")
        put(
            ws.cell(row, 9),
            f'=HYPERLINK("#\'4_Daily_Journal\'!A"&MATCH(A{row},\'4_Daily_Journal\'!A:A,0),"📖 查看當日分析對話與覆盤")',
            Font(name=FONT, size=11, bold=True, color=NAVY, underline="single"),
            green,
            center,
        )
        ws.row_dimensions[row].height = 24

    last = 2 + len(days)
    chart = LineChart()
    chart.title = "2026-27 NBA 量化操盤累積資金曲線 (Equity Curve)"
    chart.style = 10
    chart.y_axis.title = "累積損益U"
    chart.x_axis.title = "日期"
    chart.height = 8
    chart.width = 20
    chart.legend = None
    data = Reference(ws, min_col=7, min_row=2, max_row=last)
    cats = Reference(ws, min_col=1, min_row=3, max_row=last)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    series = chart.series[0]
    series.graphicalProperties.line.solidFill = NAVY
    series.graphicalProperties.line.width = 25000
    series.marker = Marker(symbol="circle", size=7)
    series.marker.graphicalProperties.solidFill = NAVY
    series.marker.graphicalProperties.line.solidFill = NAVY
    ws.add_chart(chart, "A14")

    widths(ws, {"A": 16, "B": 42, "C": 14, "D": 14, "E": 16, "F": 18, "G": 16, "H": 14, "I": 38})
    return ws, last


def _games_on(day: str) -> list[tuple]:
    return [game for game in GAMES if game[0] == day]


def _day_climate(day: str) -> str:
    seen: list[str] = []
    for game in _games_on(day):
        climate = game[3]
        if climate not in seen:
            seen.append(climate)
    return "、".join(seen)


def _pregame_text(day: str) -> str:
    lines = [
        "【賽前量化分析與對話推論】",
        "示範文字，對應第 2 頁同一天的推薦。這裡不是 NotebookLM 的原始對話。",
        "",
    ]
    for game in _games_on(day):
        _day, gid, matchup, climate, pick, line, cf, signal, stake, _score, _margin, _outcome, _note = game
        flow = "金流留白" if cf is None else f"尾盤金流 {cf:.0%}"
        stake_text = "不進場，注碼 0U" if not stake else f"建議注碼 {stake:.1f}U"
        lines.append(f"{gid}　{matchup}")
        lines.append(f"氣候判讀：{climate}。訊號：{signal}。{flow}。")
        lines.append(f"賽前結論：{pick}，{line}。{stake_text}。")
        lines.append("")
    return "\n".join(lines).strip()


def _postgame_text(day: str) -> str:
    lines = [
        "【賽後對戰結算與復盤日誌】",
        "比分是公開賽果。結果依第 2 頁的盤口結算。",
        "",
    ]
    for game in _games_on(day):
        _day, gid, matchup, _climate, pick, _line, _cf, _signal, stake, score, margin, outcome, note = game
        margin_text = "—" if margin is None else f"{margin:+.1f}"
        stake_text = "未下注" if not stake else f"{stake:.1f}U"
        lines.append(f"{gid}　{matchup}　推薦 {pick}")
        lines.append(f"終場 {score}　讓分差 {margin_text}　注碼 {stake_text}　盤口結果 {outcome}")
        lines.append(note)
        lines.append("")
    return "\n".join(lines).strip()


def _text_height(text: str) -> int:
    lines = text.count("\n") + 1
    return min(320, max(96, lines * 18 + 12))


def build_journal(wb: Workbook):
    ws = wb.create_sheet("4_Daily_Journal")
    apply_grid(ws)
    ws.sheet_view.zoomScale = 110
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    put(
        ws["B1"],
        "賽前對話與賽後檢討日誌　　A 欄的日期用來讓第 3 頁跳到這一列。閱讀文字在 B 欄。",
        font_title,
        align=Alignment(horizontal="left", vertical="center", wrap_text=True),
    )
    ws["B1"].border = Border()
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"
    ws.print_title_rows = "1:1"

    days = []
    for game in GAMES:
        if game[0] not in days:
            days.append(game[0])

    row = 3
    top_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
    for day in days:
        put(
            ws.cell(row, 1),
            datetime.strptime(day, "%Y-%m-%d").date(),
            font_head,
            fill_navy,
            center,
            "YYYY-MM-DD",
        )
        put(
            ws.cell(row, 2),
            f"日期 {day} | {_day_climate(day)}",
            font_head,
            fill_navy,
            Alignment(horizontal="left", vertical="center", wrap_text=True),
        )
        ws.row_dimensions[row].height = 28

        pregame = _pregame_text(day)
        put(ws.cell(row + 1, 1), None, font_body, fill_ice, top_align)
        put(ws.cell(row + 1, 2), pregame, font_body, fill_ice, top_align)
        ws.row_dimensions[row + 1].height = _text_height(pregame)

        postgame = _postgame_text(day)
        put(ws.cell(row + 2, 1), None, font_body, fill_green, top_align)
        put(ws.cell(row + 2, 2), postgame, font_body, fill_green, top_align)
        ws.row_dimensions[row + 2].height = _text_height(postgame)

        put(ws.cell(row + 3, 1), None, font_body, fill_white, center)
        put(
            ws.cell(row + 3, 2),
            '=HYPERLINK("#\'3_Equity_Timeline\'!A1","⬆️ 返回時間軸總覽")',
            Font(name=FONT, size=11, bold=True, color=NAVY, underline="single"),
            fill_white,
            Alignment(horizontal="left", vertical="center"),
        )
        ws.row_dimensions[row + 3].height = 22
        ws.row_dimensions[row + 4].height = 16
        row += 5

    widths(ws, {"A": 18, "B": 100})
    ws.auto_filter.ref = None
    ws.page_setup.orientation = "portrait"
    ws.sheet_view.showGridLines = True
    return ws


def build_dashboard(wb: Workbook, equity_last: int):
    ws = wb.create_sheet("1_Dashboard", 0)
    apply_grid(ws)
    ws.merge_cells("A1:L1")
    put(
        ws["A1"],
        "2026-27 NBA 量化操盤儀表板",
        font_title,
        align=Alignment(horizontal="left", vertical="center"),
    )
    ws["A1"].border = Border()
    ws.merge_cells("A2:L2")
    put(
        ws["A2"],
        "示範列使用 2026-10-27 至 2026-11-03 的公開比分。五張點名贏票已放進流水帳。第 3 頁可點進第 4 頁當天文字。氣候、訊號、水位、注碼和日誌文字是示範，正式記帳請改第 2 頁和第 4 頁。",
        font_note,
        align=Alignment(horizontal="left", vertical="center", wrap_text=True),
    )
    ws["A2"].border = Border()
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 32

    labels = [
        ("A", "B", "總下注場數", "=SUM('3_Equity_Timeline'!C:C)", "0", "有注碼的場次"),
        ("C", "D", "總鎖倉場數", "=SUM('3_Equity_Timeline'!D:D)", "0", "注碼為 0"),
        ("E", "F", "總投入資金 (U)", "=SUM('3_Equity_Timeline'!E:E)", '0.0"U"', "注碼加總"),
        ("G", "H", "累積淨損益 (U)", "=SUM('3_Equity_Timeline'!F:F)", '+0.000"U";-0.000"U";0.000"U"', "全期淨損益"),
        ("I", "J", "整體勝率", '=IFERROR(COUNTIF(\'2_Master_Log\'!L:L,"全贏")/(COUNTIF(\'2_Master_Log\'!L:L,"全贏")+COUNTIF(\'2_Master_Log\'!L:L,"全輸")),0)', "0.0%", "不含平手與鎖倉"),
        ("K", "L", "整體 ROI %", "=IFERROR(G5/E5,0)", "0.0%", "淨損益除以總投入"),
    ]
    for left, right, title, formula, fmt, note in labels:
        put(ws[f"{left}4"], title, font_head, fill_navy, center)
        put(ws[f"{right}4"], None, font_head, fill_navy, center)
        put(ws[f"{left}5"], formula, font_kpi, fill_green, center, fmt)
        put(ws[f"{right}5"], None, font_kpi, fill_green, center)
        put(ws[f"{left}6"], note, font_small, fill_green, center)
        put(ws[f"{right}6"], None, font_small, fill_green, center)
        ws.merge_cells(f"{left}4:{right}4")
        ws.merge_cells(f"{left}5:{right}5")
        ws.merge_cells(f"{left}6:{right}6")
    ws.row_dimensions[4].height = 22
    ws.row_dimensions[5].height = 36
    ws.row_dimensions[6].height = 18

    put(ws["A8"], "大盤氣候勝率與效益分析", font_section, fill_section, Alignment(horizontal="left", vertical="center"))
    for col in range(2, 10):
        put(ws.cell(8, col), None, font_section, fill_section, center)
    ws.merge_cells("A8:I8")
    climate_headers = ["大盤氣候", "總場數", "下注場數", "鎖倉場數", "勝場", "敗場", "勝率%", "淨損益U", "ROI%"]
    for col, header in enumerate(climate_headers, start=1):
        put(ws.cell(9, col), header, font_head, fill_navy, center)
    for index, name in enumerate(CLIMATES):
        row = 10 + index
        put(ws.cell(row, 1), name, font_label, fill_ice, left_wrap)
        put(ws.cell(row, 2), f"=COUNTIF('2_Master_Log'!D:D,A{row})", font_body, fill_green, center, "0")
        put(ws.cell(row, 3), f"=COUNTIFS('2_Master_Log'!D:D,A{row},'2_Master_Log'!I:I,\">0\")", font_body, fill_green, center, "0")
        put(ws.cell(row, 4), f"=COUNTIFS('2_Master_Log'!D:D,A{row},'2_Master_Log'!I:I,0)", font_body, fill_green, center, "0")
        put(ws.cell(row, 5), f"=COUNTIFS('2_Master_Log'!D:D,A{row},'2_Master_Log'!L:L,\"全贏\")", font_body, fill_green, center, "0")
        put(ws.cell(row, 6), f"=COUNTIFS('2_Master_Log'!D:D,A{row},'2_Master_Log'!L:L,\"全輸\")", font_body, fill_green, center, "0")
        put(ws.cell(row, 7), f"=IFERROR(E{row}/(E{row}+F{row}),0)", font_body, fill_green, center, "0.0%")
        put(ws.cell(row, 8), f"=SUMIF('2_Master_Log'!D:D,A{row},'2_Master_Log'!M:M)", font_body, fill_green, center, '+0.000"U";-0.000"U";0.000"U"')
        put(ws.cell(row, 9), f"=IFERROR(H{row}/SUMIFS('2_Master_Log'!I:I,'2_Master_Log'!D:D,A{row}),0)", font_body, fill_green, center, "0.0%")
        ws.row_dimensions[row].height = 22
    put(ws.cell(13, 1), "合計", font_label, fill_total, center)
    for col, letter in enumerate(["B", "C", "D", "E", "F"], start=2):
        put(ws.cell(13, col), f"=SUM({letter}10:{letter}12)", font_label, fill_total, center, "0")
    put(ws.cell(13, 7), "=IFERROR(E13/(E13+F13),0)", font_label, fill_total, center, "0.0%")
    put(ws.cell(13, 8), "=SUM(H10:H12)", font_label, fill_total, center, '+0.000"U";-0.000"U";0.000"U"')
    put(ws.cell(13, 9), '=IFERROR(H13/SUMIFS(\'2_Master_Log\'!I:I,\'2_Master_Log\'!D:D,"<>"),0)', font_label, fill_total, center, "0.0%")
    ws.row_dimensions[8].height = 24
    ws.row_dimensions[9].height = 22
    ws.row_dimensions[13].height = 22

    put(ws["A15"], "核心訊號確信度勝率追蹤", font_section, fill_section, Alignment(horizontal="left", vertical="center"))
    for col in range(2, 7):
        put(ws.cell(15, col), None, font_section, fill_section, center)
    ws.merge_cells("A15:F15")
    signal_headers = ["訊號類型", "進場次數", "勝場", "敗場", "勝率%", "淨損益U"]
    for col, header in enumerate(signal_headers, start=1):
        put(ws.cell(16, col), header, font_head, fill_navy, center)
    signal_rows = ("RLM逆向推盤", "低溫盲區受讓", "KeyPoint鎖水", "高溫熔斷")
    for index, name in enumerate(signal_rows):
        row = 17 + index
        put(ws.cell(row, 1), name, font_label, fill_ice, left_wrap)
        put(ws.cell(row, 2), f"=COUNTIFS('2_Master_Log'!H:H,A{row},'2_Master_Log'!I:I,\">0\")", font_body, fill_green, center, "0")
        put(ws.cell(row, 3), f"=COUNTIFS('2_Master_Log'!H:H,A{row},'2_Master_Log'!L:L,\"全贏\")", font_body, fill_green, center, "0")
        put(ws.cell(row, 4), f"=COUNTIFS('2_Master_Log'!H:H,A{row},'2_Master_Log'!L:L,\"全輸\")", font_body, fill_green, center, "0")
        put(ws.cell(row, 5), f"=IFERROR(C{row}/(C{row}+D{row}),0)", font_body, fill_green, center, "0.0%")
        put(ws.cell(row, 6), f"=SUMIF('2_Master_Log'!H:H,A{row},'2_Master_Log'!M:M)", font_body, fill_green, center, '+0.000"U";-0.000"U";0.000"U"')
        ws.row_dimensions[row].height = 22
    put(ws.cell(21, 1), "合計", font_label, fill_total, center)
    for col, letter in enumerate(["B", "C", "D"], start=2):
        put(ws.cell(21, col), f"=SUM({letter}17:{letter}20)", font_label, fill_total, center, "0")
    put(ws.cell(21, 5), "=IFERROR(C21/(C21+D21),0)", font_label, fill_total, center, "0.0%")
    put(ws.cell(21, 6), "=SUM(F17:F20)", font_label, fill_total, center, '+0.000"U";-0.000"U";0.000"U"')
    ws.row_dimensions[15].height = 24
    ws.row_dimensions[16].height = 22
    ws.row_dimensions[21].height = 22

    climate_chart = BarChart()
    climate_chart.type = "col"
    climate_chart.title = "三大氣候淨損益"
    climate_chart.y_axis.title = "淨損益U"
    climate_chart.style = 10
    climate_chart.height = 7
    climate_chart.width = 15
    climate_chart.legend = None
    pdata = Reference(ws, min_col=8, min_row=9, max_row=12)
    pcats = Reference(ws, min_col=1, min_row=10, max_row=12)
    climate_chart.add_data(pdata, titles_from_data=True)
    climate_chart.set_categories(pcats)
    climate_chart.shape = 4
    climate_chart.series[0].graphicalProperties.solidFill = NAVY
    ws.add_chart(climate_chart, "A23")

    signal_chart = BarChart()
    signal_chart.type = "col"
    signal_chart.title = "核心訊號勝率"
    signal_chart.y_axis.title = "勝率"
    signal_chart.style = 10
    signal_chart.height = 7
    signal_chart.width = 15
    signal_chart.legend = None
    signal_chart.y_axis.scaling.min = 0
    signal_chart.y_axis.scaling.max = 1
    sdata = Reference(ws, min_col=5, min_row=16, max_row=20)
    scats = Reference(ws, min_col=1, min_row=17, max_row=20)
    signal_chart.add_data(sdata, titles_from_data=True)
    signal_chart.set_categories(scats)
    signal_chart.series[0].graphicalProperties.solidFill = SECTION
    ws.add_chart(signal_chart, "E23")

    widths(ws, {
        "A": 22, "B": 14, "C": 16, "D": 14, "E": 18, "F": 16,
        "G": 18, "H": 16, "I": 14, "J": 14, "K": 16, "L": 16,
    })
    ws.freeze_panes = "A4"
    ws.oddHeader.left.text = "2026-27 NBA 量化操盤日記"
    return ws


def main():
    wb = Workbook()
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    build_master(wb)
    _, equity_last = build_equity(wb)
    build_journal(wb)
    build_dashboard(wb, equity_last)
    # 打開時先看儀表板
    wb.active = wb["1_Dashboard"]
    wb.properties.title = "2026-27 NBA 量化操盤日記"
    wb.properties.creator = "NBA 資料助手"
    wb.save(OUT)
    print(OUT)


if __name__ == "__main__":
    main()
