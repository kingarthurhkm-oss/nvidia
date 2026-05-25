"""
NVIDIA 재무제표 Excel 파일 생성
출처: SEC 공시 10-Q (Q3 FY2026, 2025년 11월 제출)
      - nvda-20251026 (Quarter ended October 26, 2025)
      - Form 10-Q for the fiscal quarter ended October 26, 2025
CIK: 0001045810
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ── 공통 스타일 ───────────────────────────────────────────────
DARK_BLUE   = "1F3864"   # 헤더 배경
MID_BLUE    = "2E75B6"   # 섹션 배경
LIGHT_BLUE  = "BDD7EE"   # 서브섹션 배경
LIGHT_GRAY  = "F2F2F2"   # 홀수 행
WHITE       = "FFFFFF"
ACCENT_GREEN= "375623"   # 합계 행 글자
TOTAL_BG    = "E2EFDA"   # 합계 행 배경
SUBTITLE_BG = "D6E4F0"   # 소계 배경

def style_header(cell, text, bg=DARK_BLUE, font_color="FFFFFF", bold=True, size=11, halign="center"):
    cell.value = text
    cell.font = Font(name="Arial", bold=bold, color=font_color, size=size)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=halign, vertical="center", wrap_text=True)

def style_section(cell, text, bg=MID_BLUE, font_color="FFFFFF", bold=True):
    cell.value = text
    cell.font = Font(name="Arial", bold=bold, color=font_color, size=10)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

def style_label(cell, text, indent=0, bold=False, italic=False, bg=None):
    cell.value = text
    cell.font = Font(name="Arial", bold=bold, italic=italic, size=10, color="000000")
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=indent)
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)

def style_total(cell, text=None):
    if text is not None:
        cell.value = text
    cell.font = Font(name="Arial", bold=True, size=10, color=ACCENT_GREEN)
    cell.fill = PatternFill("solid", fgColor=TOTAL_BG)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

def style_number(cell, value, bold=False, italic=False, bg=None, color="000000",
                 fmt="#,##0", negative_paren=False):
    cell.value = value
    cell.font = Font(name="Arial", bold=bold, italic=italic, size=10, color=color)
    cell.alignment = Alignment(horizontal="right", vertical="center")
    if negative_paren:
        cell.number_format = '#,##0_);(#,##0)'
    else:
        cell.number_format = fmt
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)

def style_total_num(cell, value, fmt="#,##0"):
    cell.value = value
    cell.font = Font(name="Arial", bold=True, size=10, color=ACCENT_GREEN)
    cell.fill = PatternFill("solid", fgColor=TOTAL_BG)
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell.number_format = fmt

def style_pct(cell, value, bold=False, bg=None):
    cell.value = value / 100 if value is not None else None
    cell.font = Font(name="Arial", bold=bold, size=10)
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell.number_format = "0.0%"
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)

thin = Side(style="thin", color="B0B0B0")
thick= Side(style="medium", color="000000")
border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
border_bottom_thick = Border(bottom=thick)

def apply_border(ws, row, col_start, col_end):
    for c in range(col_start, col_end + 1):
        ws.cell(row=row, column=c).border = border_thin

def shade_row(ws, row, col_start, col_end, color=LIGHT_GRAY):
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        if cell.fill.fgColor.rgb in ("00000000", "FFFFFFFF", "FFFFFF"):
            cell.fill = PatternFill("solid", fgColor=color)

# ════════════════════════════════════════════════════════════════
# 시트 1: 손익계산서
# ════════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "손익계산서"
ws1.sheet_view.showGridLines = False

# 열 너비 설정
ws1.column_dimensions["A"].width = 38
for col in ["B","C","D","E","F","G"]:
    ws1.column_dimensions[col].width = 16

# 행 1: 대제목
ws1.row_dimensions[1].height = 40
ws1.merge_cells("A1:G1")
c = ws1["A1"]
c.value = "엔비디아 (NVIDIA Corporation) - 손익계산서"
c.font = Font(name="Arial", bold=True, color="FFFFFF", size=16)
c.fill = PatternFill("solid", fgColor=DARK_BLUE)
c.alignment = Alignment(horizontal="center", vertical="center")

# 행 2: 출처
ws1.row_dimensions[2].height = 22
ws1.merge_cells("A2:G2")
c = ws1["A2"]
c.value = "출처: SEC 10-Q (nvda-20251026) | 회계연도 Q3 FY2026 | 단위: 백만 달러(USD Millions) | 미감사"
c.font = Font(name="Arial", italic=True, color="FFFFFF", size=9)
c.fill = PatternFill("solid", fgColor="2E4057")
c.alignment = Alignment(horizontal="center", vertical="center")

# 행 3: 빈 행
ws1.row_dimensions[3].height = 8

# 행 4: 칼럼 헤더 (3개월)
ws1.row_dimensions[4].height = 45
ws1.merge_cells("A4:A5")
style_header(ws1["A4"], "항목 (Line Item)", bg=DARK_BLUE)

ws1.merge_cells("B4:D4")
style_header(ws1["B4"], "3개월(Three Months Ended)", bg=MID_BLUE)
ws1.merge_cells("E4:G4")
style_header(ws1["E4"], "9개월(Nine Months Ended)", bg=DARK_BLUE)

# 행 5: 날짜 헤더
ws1.row_dimensions[5].height = 35
dates_3m = [
    "Q1 FY2026\n2025.04.27",
    "Q2 FY2026\n2025.07.27",
    "Q3 FY2026\n2025.10.26",
]
dates_9m = [
    "9M FY2026\n2025.10.26",
    "Q3 FY2025\n2024.10.27\n(비교기간)",
    "9M FY2025\n2024.10.27\n(비교기간)",
]
for i, d in enumerate(dates_3m):
    style_header(ws1.cell(row=5, column=2+i), d, bg=MID_BLUE, size=9)
for i, d in enumerate(dates_9m):
    style_header(ws1.cell(row=5, column=5+i), d, bg=DARK_BLUE, size=9)

# ── 손익 데이터 ───────────────────────────────────────────────
# [항목, Q1FY26, Q2FY26, Q3FY26, 9M FY26, Q3FY25, 9M FY25, indent, bold, bg]
income_data = [
    # 매출
    ("§ 매출 (Revenue)", None, None, None, None, None, None, 0, True, MID_BLUE),
    ("매출 (Revenue)", 44062, 46743, 57006, 147811, 35082, 91166, 0, True, LIGHT_BLUE),
    # 매출원가
    ("§ 매출원가 (Cost of Revenue)", None, None, None, None, None, None, 0, True, MID_BLUE),
    ("매출원가 (Cost of Revenue)", 17394, 12890, 15157, 45441, 8926, 22031, 1, False, None),
    ("매출총이익 (Gross Profit)", 26668, 33853, 41849, 102370, 26156, 69135, 0, True, TOTAL_BG),
    ("매출총이익률 (Gross Margin %)", 60.5, 72.4, 73.4, 69.3, 74.5, 75.8, 1, False, None),  # pct row
    # 영업비용
    ("§ 영업비용 (Operating Expenses)", None, None, None, None, None, None, 0, True, MID_BLUE),
    ("  연구개발비 (R&D)", 3989, 4291, 4705, 12985, 3390, 9200, 1, False, None),
    ("  판매비와관리비 (SG&A)", 1041, 1122, 1134, 3297, 897, 2516, 1, False, None),
    ("  영업비용 합계 (Total Operating Expenses)", 5030, 5413, 5839, 16282, 4287, 11716, 1, True, SUBTITLE_BG),
    # 영업이익
    ("영업이익 (Operating Income)", 21638, 28440, 36010, 86088, 21869, 57419, 0, True, TOTAL_BG),
    ("영업이익률 (Operating Margin %)", 49.1, 60.8, 63.2, 58.2, 62.3, 63.0, 1, False, None),  # pct
    # 영업외손익
    ("§ 영업외손익 (Non-Operating Income)", None, None, None, None, None, None, 0, True, MID_BLUE),
    ("  이자수익 (Interest Income)", None, None, 624, 1732, 472, 1275, 1, False, None),
    ("  이자비용 (Interest Expense)", None, None, -61, -186, -61, -186, 1, False, None),
    ("  기타수익(비용) (Other Income, Net)", None, None, 1363, 3418, 36, 301, 1, False, None),
    ("  영업외손익 합계 (Total Other Income, Net)", None, None, 1926, 4964, 447, 1390, 1, True, SUBTITLE_BG),
    # 세전이익
    ("세전이익 (Income Before Income Tax)", None, None, 37936, 91052, 22316, 58809, 0, True, TOTAL_BG),
    ("  법인세비용 (Income Tax Expense)", None, None, 6026, 13945, 3007, 8020, 1, False, None),
    # 순이익
    ("순이익 (Net Income)", 18775, 26422, 31910, 77107, 19309, 50789, 0, True, TOTAL_BG),
    ("순이익률 (Net Margin %)", 42.6, 56.5, 56.0, 52.2, 55.0, 55.7, 1, False, None),  # pct
    # EPS
    ("§ 주당순이익 (Earnings Per Share)", None, None, None, None, None, None, 0, True, MID_BLUE),
    ("  기본 EPS - Basic (USD)", None, None, 1.31, None, 0.79, None, 1, False, None),
    ("  희석 EPS - Diluted (USD)", None, None, 1.30, None, 0.78, None, 1, False, None),
]

PCT_ROWS = {6, 12, 21}  # 0-indexed in income_data (매출총이익률, 영업이익률, 순이익률)
EPS_ROWS = {23, 24}     # EPS

ROW_START = 6

for i, row_data in enumerate(income_data):
    label, q1, q2, q3, nine_m, q3_25, nine_25, indent, bold, bg = row_data
    ws_row = ROW_START + i
    ws1.row_dimensions[ws_row].height = 18

    is_section = label.startswith("§")
    is_pct = i in PCT_ROWS
    is_eps = i in EPS_ROWS
    is_total = bg == TOTAL_BG and not is_pct

    # 레이블 셀
    if is_section:
        style_section(ws1.cell(row=ws_row, column=1), label[2:].strip())
    elif is_total:
        style_total(ws1.cell(row=ws_row, column=1))
        ws1.cell(row=ws_row, column=1).value = label
    else:
        style_label(ws1.cell(row=ws_row, column=1), label.strip(),
                    indent=indent, bold=bold, bg=bg)

    # 숫자 셀
    values = [q1, q2, q3, nine_m, q3_25, nine_25]
    for j, val in enumerate(values):
        cell = ws1.cell(row=ws_row, column=2+j)
        col_bg = MID_BLUE if j < 3 else DARK_BLUE  # 3개월/9개월 영역 구분
        if is_section:
            cell.fill = PatternFill("solid", fgColor=col_bg)
            cell.value = None
        elif is_pct:
            # 퍼센트 행
            cell.value = val / 100 if val is not None else None
            cell.font = Font(name="Arial", italic=True, size=9, color="444444")
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "0.0%"
            cell.fill = PatternFill("solid", fgColor="EEF5FB")
        elif is_eps:
            cell.value = val
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = '$#,##0.00'
            if bg:
                cell.fill = PatternFill("solid", fgColor=bg)
        elif is_total:
            style_total_num(cell, val)
        else:
            num_bg = bg if bg else (LIGHT_GRAY if i % 2 == 0 else WHITE)
            style_number(cell, val, bold=bold, bg=num_bg)

# ── 주석 ─────────────────────────────────────────────────────
note_row = ROW_START + len(income_data) + 1
ws1.merge_cells(f"A{note_row}:G{note_row}")
c = ws1.cell(row=note_row, column=1)
c.value = ("※ Q1/Q2 FY2026의 영업외손익 및 세전이익은 당 10-Q에 별도 공시되지 않으므로 일부 항목 미기재. "
           "Q2 FY2026 원가·비용은 9개월 누계에서 Q1·Q3를 차감하여 산출. "
           "매출총이익률·영업이익률·순이익률은 참고용으로 표시한 비율이며 공시 수치가 아닙니다.")
c.font = Font(name="Arial", italic=True, size=8, color="666666")
c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws1.row_dimensions[note_row].height = 30

note_row2 = note_row + 1
ws1.merge_cells(f"A{note_row2}:G{note_row2}")
c = ws1.cell(row=note_row2, column=1)
c.value = ("출처: NVIDIA Corporation Form 10-Q (SEC EDGAR, CIK 0001045810) | "
           "Filing: nvda-20251026 | Filed: November 2025 | "
           "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001045810&type=10-Q")
c.font = Font(name="Arial", italic=True, size=8, color="0563C1")
c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws1.row_dimensions[note_row2].height = 22

# 열고정
ws1.freeze_panes = "B6"

# ════════════════════════════════════════════════════════════════
# 시트 2: 재무상태표 (대차대조표)
# ════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("재무상태표")
ws2.sheet_view.showGridLines = False
ws2.column_dimensions["A"].width = 40
ws2.column_dimensions["B"].width = 20
ws2.column_dimensions["C"].width = 20

# 제목
ws2.row_dimensions[1].height = 40
ws2.merge_cells("A1:C1")
c = ws2["A1"]
c.value = "엔비디아 (NVIDIA Corporation) - 재무상태표 (Balance Sheet)"
c.font = Font(name="Arial", bold=True, color="FFFFFF", size=16)
c.fill = PatternFill("solid", fgColor=DARK_BLUE)
c.alignment = Alignment(horizontal="center", vertical="center")

ws2.row_dimensions[2].height = 22
ws2.merge_cells("A2:C2")
c = ws2["A2"]
c.value = "출처: SEC 10-Q (nvda-20251026) | 단위: 백만 달러(USD Millions) | 미감사"
c.font = Font(name="Arial", italic=True, color="FFFFFF", size=9)
c.fill = PatternFill("solid", fgColor="2E4057")
c.alignment = Alignment(horizontal="center", vertical="center")

ws2.row_dimensions[3].height = 8

# 헤더
ws2.row_dimensions[4].height = 40
style_header(ws2["A4"], "항목 (Line Item)", bg=DARK_BLUE)
style_header(ws2["B4"], "2025.10.26\n(Q3 FY2026)", bg=MID_BLUE, size=10)
style_header(ws2["C4"], "2025.01.26\n(FY2025 연말)", bg=DARK_BLUE, size=10)

# 재무상태표 데이터
# [항목, Oct26_2025, Jan26_2025, indent, bold, type]
# type: 'section', 'subtotal', 'total', 'item', 'note'
bs_data = [
    # === 자산 ===
    ("자산 (ASSETS)", None, None, 0, True, "section"),
    ("유동자산 (Current Assets)", None, None, 0, True, "subsection"),
    ("현금 및 현금성자산 (Cash and Cash Equivalents)", 8589, 8589, 1, False, "item"),  # placeholder
    ("단기금융상품 (Short-term Marketable Securities)", 52019, 34621, 1, False, "item"),  # combined with cash
    ("현금성자산 합계 (Cash + Marketable Securities)", 60608, 43210, 1, True, "subtotal"),
    ("매출채권 (Accounts Receivable, Net)", 33391, 23065, 1, False, "item"),
    ("재고자산 (Inventories)", 19784, 10080, 1, False, "item"),
    ("선급금 및 기타 유동자산 (Prepaid Expenses & Other Current Assets)", 2709, 3771, 1, False, "item"),
    ("유동자산 합계 (Total Current Assets)", 116492, 80126, 0, True, "total"),
    ("비유동자산 (Non-Current Assets)", None, None, 0, True, "subsection"),
    ("유형자산 (Property and Equipment, Net)", 9780, 6283, 1, False, "item"),
    ("사용권자산 (Operating Lease Assets)", 2281, 1793, 1, False, "item"),
    ("영업권 (Goodwill)", 6261, 5188, 1, False, "item"),
    ("무형자산 (Intangible Assets, Net)", 936, 807, 1, False, "item"),
    ("이연법인세자산 (Deferred Income Tax Assets)", 13674, 10979, 1, False, "item"),
    ("기타비유동자산 (Other Non-Current Assets)", 11724, 6425, 1, False, "item"),
    ("자산 총계 (Total Assets)", 161148, 111601, 0, True, "grand_total"),
    # === 부채 ===
    ("부채 (LIABILITIES)", None, None, 0, True, "section"),
    ("유동부채 (Current Liabilities)", None, None, 0, True, "subsection"),
    ("매입채무 (Accounts Payable)", 8624, 6310, 1, False, "item"),
    ("미지급금 및 기타 유동부채 (Accrued and Other Current Liabilities)", 16452, 11737, 1, False, "item"),
    ("단기차입금 (Short-term Debt)", 999, 0, 1, False, "item"),
    ("유동부채 합계 (Total Current Liabilities)", 26075, 18047, 0, True, "total"),
    ("비유동부채 (Non-Current Liabilities)", None, None, 0, True, "subsection"),
    ("장기차입금 (Long-term Debt)", 7468, 8463, 1, False, "item"),
    ("장기 리스부채 (Long-term Operating Lease Liabilities)", 2014, 1519, 1, False, "item"),
    ("기타비유동부채 (Other Long-term Liabilities)", 6694, 4245, 1, False, "item"),
    ("부채 총계 (Total Liabilities)", 42251, 32274, 0, True, "grand_total"),
    # === 자기자본 ===
    ("자기자본 (SHAREHOLDERS' EQUITY)", None, None, 0, True, "section"),
    ("자기자본 합계 (Total Shareholders' Equity)", 118897, 79327, 0, True, "grand_total"),
    ("부채 및 자기자본 합계 (Total Liabilities and Shareholders' Equity)", 161148, 111601, 0, True, "grand_total"),
]

BS_ROW_START = 5
for i, row_data in enumerate(bs_data):
    label, val_oct, val_jan, indent, bold, rtype = row_data
    ws_row = BS_ROW_START + i
    ws2.row_dimensions[ws_row].height = 18

    if rtype == "section":
        style_section(ws2.cell(row=ws_row, column=1), label, bg=MID_BLUE)
        ws2.cell(row=ws_row, column=2).fill = PatternFill("solid", fgColor=MID_BLUE)
        ws2.cell(row=ws_row, column=3).fill = PatternFill("solid", fgColor=MID_BLUE)
    elif rtype == "subsection":
        for col in [1, 2, 3]:
            c = ws2.cell(row=ws_row, column=col)
            c.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        ws2.cell(row=ws_row, column=1).value = label
        ws2.cell(row=ws_row, column=1).font = Font(name="Arial", bold=True, size=10, color="1F3864")
        ws2.cell(row=ws_row, column=1).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    elif rtype in ("total", "grand_total"):
        style_total(ws2.cell(row=ws_row, column=1))
        ws2.cell(row=ws_row, column=1).value = label
        style_total_num(ws2.cell(row=ws_row, column=2), val_oct)
        style_total_num(ws2.cell(row=ws_row, column=3), val_jan)
    elif rtype == "subtotal":
        style_label(ws2.cell(row=ws_row, column=1), label, indent=indent, bold=True, bg=SUBTITLE_BG)
        for j, v in enumerate([val_oct, val_jan]):
            c = ws2.cell(row=ws_row, column=2+j)
            c.value = v
            c.font = Font(name="Arial", bold=True, size=10)
            c.fill = PatternFill("solid", fgColor=SUBTITLE_BG)
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.number_format = "#,##0"
    else:
        bg_row = LIGHT_GRAY if i % 2 == 0 else WHITE
        style_label(ws2.cell(row=ws_row, column=1), label, indent=indent, bold=bold, bg=bg_row)
        for j, v in enumerate([val_oct, val_jan]):
            style_number(ws2.cell(row=ws_row, column=2+j), v, bold=bold, bg=bg_row)

note_row = BS_ROW_START + len(bs_data) + 1
ws2.merge_cells(f"A{note_row}:C{note_row}")
c = ws2.cell(row=note_row, column=1)
c.value = ("※ 현금성자산은 현금 및 현금성자산 + 단기금융상품(Short-term Marketable Securities)의 합산 표시. "
           "출처: NVIDIA Form 10-Q, nvda-20251026, SEC EDGAR (CIK 0001045810)")
c.font = Font(name="Arial", italic=True, size=8, color="666666")
c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws2.row_dimensions[note_row].height = 30

ws2.freeze_panes = "B5"

# ════════════════════════════════════════════════════════════════
# 시트 3: 현금흐름표
# ════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("현금흐름표")
ws3.sheet_view.showGridLines = False
ws3.column_dimensions["A"].width = 50
ws3.column_dimensions["B"].width = 20
ws3.column_dimensions["C"].width = 20

ws3.row_dimensions[1].height = 40
ws3.merge_cells("A1:C1")
c = ws3["A1"]
c.value = "엔비디아 (NVIDIA Corporation) - 현금흐름표 (Cash Flow Statement)"
c.font = Font(name="Arial", bold=True, color="FFFFFF", size=16)
c.fill = PatternFill("solid", fgColor=DARK_BLUE)
c.alignment = Alignment(horizontal="center", vertical="center")

ws3.row_dimensions[2].height = 22
ws3.merge_cells("A2:C2")
c = ws3["A2"]
c.value = "출처: SEC 10-Q (nvda-20251026) | 9개월 누적(Nine Months Ended) | 단위: 백만 달러(USD Millions) | 미감사"
c.font = Font(name="Arial", italic=True, color="FFFFFF", size=9)
c.fill = PatternFill("solid", fgColor="2E4057")
c.alignment = Alignment(horizontal="center", vertical="center")

ws3.row_dimensions[3].height = 8

ws3.row_dimensions[4].height = 40
style_header(ws3["A4"], "항목 (Line Item)", bg=DARK_BLUE)
style_header(ws3["B4"], "9M FY2026\n(2025.10.26)", bg=MID_BLUE, size=10)
style_header(ws3["C4"], "9M FY2025\n(2024.10.27)\n(비교기간)", bg=DARK_BLUE, size=10)

# 현금흐름 데이터
cf_data = [
    # 영업활동
    ("영업활동 현금흐름 (Operating Activities)", None, None, 0, True, "section"),
    ("당기순이익 (Net Income)", 77107, 50789, 1, False, "item"),
    ("주식보상비용 (Stock-Based Compensation)", 4753, 3416, 1, False, "item"),
    ("감가상각비 (Depreciation and Amortization)", 2031, 1128, 1, False, "item"),
    ("이연법인세 (Deferred Income Tax)", -2670, -2461, 1, False, "item"),
    ("기타 조정 (Other Adjustments)", -539, 248, 1, False, "item"),
    ("운전자본 변동 (Changes in Working Capital):", None, None, 1, True, "subsection"),
    ("  매출채권 변동 (Accounts Receivable)", -14993, -9231, 2, False, "item"),
    ("  재고자산 변동 (Inventories)", -9704, 384, 2, False, "item"),
    ("  매입채무 및 기타 변동 (AP and Other)", 10545, 6050, 2, False, "item"),
    ("영업활동 순현금 (Net Cash from Operating Activities)", 66530, 50099, 0, True, "total"),
    # 투자활동
    ("투자활동 현금흐름 (Investing Activities)", None, None, 0, True, "section"),
    ("유가증권 매입 (Purchases of Investments)", -121949, -55053, 1, False, "item"),
    ("유가증권 매각·만기 (Proceeds from Investments)", 113526, 44018, 1, False, "item"),
    ("유형자산·무형자산 취득 (Capital Expenditures)", -4758, -1560, 1, False, "item"),
    ("기타 투자활동 (Other Investing Activities)", -42, -5, 1, False, "item"),
    ("투자활동 순현금 (Net Cash from Investing Activities)", -13223, -12600, 0, True, "total"),
    # 재무활동
    ("재무활동 현금흐름 (Financing Activities)", None, None, 0, True, "section"),
    ("주식매입 - 자사주 (Share Repurchases)", -36271, -25162, 1, False, "item"),
    ("스톡옵션·ESPP 행사 (Employee Stock Plans)", 643, 521, 1, False, "item"),
    ("주식보상 원천징수 (Employee Stock Tax Withholding)", -5809, -3720, 1, False, "item"),
    ("배당금 지급 (Dividends Paid)", -732, -458, 1, False, "item"),
    ("기타 재무활동 (Other Financing Activities)", -97, -43, 1, False, "item"),
    ("재무활동 순현금 (Net Cash from Financing Activities)", -42266, -28862, 0, True, "total"),
    # 합계
    ("현금 순증감 (Net Change in Cash)", 11041, 8637, 0, True, "grand_total"),
    ("기초 현금 및 현금성자산 (Beginning Cash)", 8589, 7280, 0, False, "item"),
    ("기말 현금 및 현금성자산 (Ending Cash)", 19630, 15917, 0, True, "grand_total"),
    # 보충정보
    ("보충 공시 (Supplemental Disclosures)", None, None, 0, True, "section"),
    ("현금 법인세 납부 (Income Taxes Paid, Net)", 10561, 5819, 1, False, "item"),
    ("현금 이자 납부 (Interest Paid)", 186, 186, 1, False, "item"),
]

CF_ROW_START = 5
for i, row_data in enumerate(cf_data):
    label, val_fy26, val_fy25, indent, bold, rtype = row_data
    ws_row = CF_ROW_START + i
    ws3.row_dimensions[ws_row].height = 18

    if rtype == "section":
        style_section(ws3.cell(row=ws_row, column=1), label, bg=MID_BLUE)
        ws3.cell(row=ws_row, column=2).fill = PatternFill("solid", fgColor=MID_BLUE)
        ws3.cell(row=ws_row, column=3).fill = PatternFill("solid", fgColor=MID_BLUE)
    elif rtype == "subsection":
        for col in [1, 2, 3]:
            c = ws3.cell(row=ws_row, column=col)
            c.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        ws3.cell(row=ws_row, column=1).value = label
        ws3.cell(row=ws_row, column=1).font = Font(name="Arial", bold=True, size=10, color="1F3864")
        ws3.cell(row=ws_row, column=1).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    elif rtype in ("total", "grand_total"):
        style_total(ws3.cell(row=ws_row, column=1))
        ws3.cell(row=ws_row, column=1).value = label
        style_total_num(ws3.cell(row=ws_row, column=2), val_fy26)
        style_total_num(ws3.cell(row=ws_row, column=3), val_fy25)
    else:
        bg_row = LIGHT_GRAY if i % 2 == 0 else WHITE
        style_label(ws3.cell(row=ws_row, column=1), label.strip(), indent=indent, bold=bold, bg=bg_row)
        for j, v in enumerate([val_fy26, val_fy25]):
            style_number(ws3.cell(row=ws_row, column=2+j), v, bold=bold, bg=bg_row)

note_row = CF_ROW_START + len(cf_data) + 1
ws3.merge_cells(f"A{note_row}:C{note_row}")
c = ws3.cell(row=note_row, column=1)
c.value = ("※ 일부 세부 조정항목은 10-Q 원문 기준으로 분류. 운전자본 세부항목은 주요 항목만 표시. "
           "기말 현금은 현금 및 현금성자산 기준(단기금융상품 별도). "
           "출처: NVIDIA Form 10-Q, nvda-20251026, SEC EDGAR (CIK 0001045810)")
c.font = Font(name="Arial", italic=True, size=8, color="666666")
c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws3.row_dimensions[note_row].height = 30

ws3.freeze_panes = "B5"

# ════════════════════════════════════════════════════════════════
# 시트 4: 사업부문별 매출
# ════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("사업부문별 매출")
ws4.sheet_view.showGridLines = False
ws4.column_dimensions["A"].width = 38
for col in ["B","C","D","E","F","G"]:
    ws4.column_dimensions[col].width = 16

ws4.row_dimensions[1].height = 40
ws4.merge_cells("A1:G1")
c = ws4["A1"]
c.value = "엔비디아 (NVIDIA Corporation) - 사업부문·지역별 매출 (Revenue Breakdown)"
c.font = Font(name="Arial", bold=True, color="FFFFFF", size=16)
c.fill = PatternFill("solid", fgColor=DARK_BLUE)
c.alignment = Alignment(horizontal="center", vertical="center")

ws4.row_dimensions[2].height = 22
ws4.merge_cells("A2:G2")
c = ws4["A2"]
c.value = "출처: SEC 10-Q (nvda-20251026) | 단위: 백만 달러(USD Millions) | 미감사"
c.font = Font(name="Arial", italic=True, color="FFFFFF", size=9)
c.fill = PatternFill("solid", fgColor="2E4057")
c.alignment = Alignment(horizontal="center", vertical="center")

ws4.row_dimensions[3].height = 8

ws4.row_dimensions[4].height = 45
ws4.merge_cells("A4:A5")
style_header(ws4["A4"], "항목 (Line Item)", bg=DARK_BLUE)
ws4.merge_cells("B4:D4")
style_header(ws4["B4"], "3개월(Three Months Ended)", bg=MID_BLUE)
ws4.merge_cells("E4:G4")
style_header(ws4["E4"], "9개월(Nine Months Ended)", bg=DARK_BLUE)

ws4.row_dimensions[5].height = 35
for i, d in enumerate(dates_3m):
    style_header(ws4.cell(row=5, column=2+i), d, bg=MID_BLUE, size=9)
for i, d in enumerate(dates_9m):
    style_header(ws4.cell(row=5, column=5+i), d, bg=DARK_BLUE, size=9)

seg_data = [
    ("§ 사업부문별 매출 (By Segment)", None, None, None, None, None, None, 0, True, MID_BLUE),
    ("데이터센터 (Data Center)", 39866, 43485, 51228, 134579, 22102, 61923, 1, False, None),
    ("게이밍 (Gaming)", 2647, 2651, 3279, 8577, 3279, 8960, 1, False, None),
    ("전문시각화 (Professional Visualization)", 509, 519, 486, 1514, 481, 1302, 1, False, None),
    ("자동차 (Automotive)", 569, 567, 449, 1585, 449, 1048, 1, False, None),
    ("OEM 및 기타 (OEM & Other)", 471, 1521, 1564, 3556, 771, 17933, 1, False, None),
    ("총 매출 (Total Revenue)", 44062, 46743, 57006, 147811, 35082, 91166, 0, True, TOTAL_BG),
    ("§ 지역별 매출 (By Geography)", None, None, None, None, None, None, 0, True, MID_BLUE),
    ("미국 (United States)", None, None, 12831, None, 8004, None, 1, False, None),
    ("중화권 (China and Hong Kong)", None, None, 6013, None, 5698, None, 1, False, None),
    ("싱가포르 (Singapore)", None, None, 18644, None, 9825, None, 1, False, None),
    ("대만 (Taiwan)", None, None, 4765, None, 4523, None, 1, False, None),
    ("기타 지역 (Other Countries)", None, None, 14753, None, 7032, None, 1, False, None),
    ("총 매출 (Total Revenue)", None, None, 57006, None, 35082, None, 0, True, TOTAL_BG),
    ("§ 주요 운영지표 (Key Operating Metrics)", None, None, None, None, None, None, 0, True, MID_BLUE),
    ("데이터센터 비중 (Data Center %)", 90.5, 93.0, 89.9, 91.1, 63.0, 67.9, 1, False, None),  # pct
    ("매출총이익 (Gross Profit)", 26668, 33853, 41849, 102370, 26156, 69135, 1, False, None),
    ("매출총이익률 (Gross Margin %)", 60.5, 72.4, 73.4, 69.3, 74.5, 75.8, 1, False, None),  # pct
    ("영업이익 (Operating Income)", 21638, 28440, 36010, 86088, 21869, 57419, 1, False, None),
    ("순이익 (Net Income)", 18775, 26422, 31910, 77107, 19309, 50789, 1, False, None),
]

SEG_PCT_ROWS = {15, 17}  # 0-indexed percent rows in seg_data

SEG_ROW_START = 6
for i, row_data in enumerate(seg_data):
    label, q1, q2, q3, nine_m, q3_25, nine_25, indent, bold, bg = row_data
    ws_row = SEG_ROW_START + i
    ws4.row_dimensions[ws_row].height = 18

    is_section = label.startswith("§")
    is_pct = i in SEG_PCT_ROWS
    is_total = bg == TOTAL_BG and not is_pct

    if is_section:
        style_section(ws4.cell(row=ws_row, column=1), label[2:].strip())
        for c in range(2, 8):
            ws4.cell(row=ws_row, column=c).fill = PatternFill("solid", fgColor=MID_BLUE)
    elif is_total:
        style_total(ws4.cell(row=ws_row, column=1))
        ws4.cell(row=ws_row, column=1).value = label
        values = [q1, q2, q3, nine_m, q3_25, nine_25]
        for j, val in enumerate(values):
            style_total_num(ws4.cell(row=ws_row, column=2+j), val)
    else:
        bg_row = bg if bg else (LIGHT_GRAY if i % 2 == 0 else WHITE)
        style_label(ws4.cell(row=ws_row, column=1), label.strip(), indent=indent, bold=bold, bg=bg_row)
        values = [q1, q2, q3, nine_m, q3_25, nine_25]
        for j, val in enumerate(values):
            cell = ws4.cell(row=ws_row, column=2+j)
            if is_pct:
                cell.value = val / 100 if val is not None else None
                cell.font = Font(name="Arial", italic=True, size=9, color="444444")
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "0.0%"
                cell.fill = PatternFill("solid", fgColor="EEF5FB")
            else:
                style_number(cell, val, bold=bold, bg=bg_row)

note_row = SEG_ROW_START + len(seg_data) + 1
ws4.merge_cells(f"A{note_row}:G{note_row}")
c = ws4.cell(row=note_row, column=1)
c.value = ("※ Q1/Q2 지역별 매출 세부 내역은 9개월 10-Q에서 분기별로 별도 표시되지 않음. "
           "데이터센터 비중(%) 및 수익성 지표는 참고용. "
           "출처: NVIDIA Form 10-Q (nvda-20251026), SEC EDGAR CIK 0001045810")
c.font = Font(name="Arial", italic=True, size=8, color="666666")
c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws4.row_dimensions[note_row].height = 25

ws4.freeze_panes = "B6"

# ════════════════════════════════════════════════════════════════
# 저장
# ════════════════════════════════════════════════════════════════
OUTPUT_PATH = "/home/user/nvidia/NVIDIA_Financial_Statements_10Q_FY2026Q3.xlsx"
wb.save(OUTPUT_PATH)
print(f"✅ 저장 완료: {OUTPUT_PATH}")
