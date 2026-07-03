# -*- coding: utf-8 -*-
"""거시경제 상황 + 주가 영향 분석 레포트(엑셀) 생성 스크립트.

데이터 기준: 2026년 7월 초 (2026년 6월 발표치·전망치 중심)
목적: 미시·거시·국제 지표를 정리하고, 각 항목이 '주가(증시)'에 미치는
      영향(방향·수혜/피해 업종)과 결론을 도출.
출처: 공개 자료(연준/BLS, 한국은행/KDI/KIET/KIEP/KIF/KCIF/KITA/HRI 등,
      IMF/EIA/Eurostat/中 NBS) 웹 취합. 일부는 잠정·전망치.

※ 본 자료는 참고용이며 투자자문·매매권유가 아님. 주가 영향은
  일반적 시장 메커니즘에 근거한 시나리오 해석임.
"""

import math
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- 색상/스타일
NAVY = "1F3864"
BLUE = "2E5496"
LTBLUE = "D6E0F0"
STRIPE = "F2F5FB"
AMBER = "BF8F00"      # 결론·주가영향 헤더
LTAMBER = "FFF2CC"    # 결론 박스 배경
GREY = "808080"
GREEN = "2E7D32"
RED = "C00000"

TITLE_FONT = Font(name="맑은 고딕", color="FFFFFF", bold=True, size=16)
SUB_FONT = Font(name="맑은 고딕", color="FFFFFF", size=10)
HDR_FONT = Font(name="맑은 고딕", bold=True, size=10)
CELL_FONT = Font(name="맑은 고딕", size=10)
NOTE_FONT = Font(name="맑은 고딕", size=9, italic=True, color=GREY)
SECTION_FONT = Font(name="맑은 고딕", color="FFFFFF", bold=True, size=11)
CALLOUT_HDR_FONT = Font(name="맑은 고딕", color="FFFFFF", bold=True, size=10)

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
LEFT_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)


def fill(color):
    return PatternFill("solid", fgColor=color)


def style_header_row(ws, row, ncols, start_col=1):
    for c in range(start_col, start_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill(LTBLUE)
        cell.font = HDR_FONT
        cell.alignment = CENTER
        cell.border = BORDER


def write_table(ws, start_row, headers, rows, col_widths=None, dir_col=None):
    """헤더 + 데이터 표. dir_col(0-index) 지정 시 ▲녹/▼적/중립회색 색상."""
    ncols = len(headers)
    for j, h in enumerate(headers, start=1):
        ws.cell(row=start_row, column=j, value=h)
    style_header_row(ws, start_row, ncols)

    for i, rowdata in enumerate(rows):
        r = start_row + 1 + i
        for j, val in enumerate(rowdata, start=1):
            cell = ws.cell(row=r, column=j, value=val)
            cell.font = CELL_FONT
            cell.border = BORDER
            cell.alignment = LEFT if j == 1 else CENTER
            if i % 2 == 1:
                cell.fill = fill(STRIPE)
            if dir_col is not None and j - 1 == dir_col:
                s = str(val)
                col = GREEN if s.startswith("▲") else RED if s.startswith("▼") else GREY
                cell.font = Font(name="맑은 고딕", size=10, bold=True, color=col)
    if col_widths:
        for j, w in enumerate(col_widths, start=1):
            ws.column_dimensions[get_column_letter(j)].width = w
    return start_row + 1 + len(rows) + 1


def section_bar(ws, row, text, ncols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    cell = ws.cell(row=row, column=1, value=text)
    cell.fill = fill(BLUE)
    cell.font = SECTION_FONT
    cell.alignment = LEFT
    ws.row_dimensions[row].height = 22
    return row + 1


def callout(ws, start_row, lines, ncols, total_width, title="▶ 설명 · 결론 · 주가 영향"):
    """설명/결론/주가영향 콜아웃 박스. lines: (라벨, 내용) 리스트."""
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=ncols)
    h = ws.cell(row=start_row, column=1, value=title)
    h.fill = fill(AMBER)
    h.font = CALLOUT_HDR_FONT
    h.alignment = LEFT
    ws.row_dimensions[start_row].height = 18

    cpl = max(20, int(total_width / 2.0))  # 병합폭 기준 한 줄 대략 글자수
    r = start_row + 1
    for label, text in lines:
        full = f"[{label}] {text}"
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
        cell = ws.cell(row=r, column=1, value=full)
        cell.fill = fill(LTAMBER)
        cell.font = CELL_FONT if label != "주가" else Font(name="맑은 고딕", size=10, bold=True)
        cell.alignment = LEFT_TOP
        cell.border = BORDER
        nlines = max(1, math.ceil(len(full) / cpl))
        ws.row_dimensions[r].height = 15 * nlines + 4
        r += 1
    return r + 1


def title_block(ws, title, subtitle, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    t = ws.cell(row=1, column=1, value=title)
    t.fill = fill(NAVY); t.font = TITLE_FONT
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 34
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    s = ws.cell(row=2, column=1, value=subtitle)
    s.fill = fill(NAVY); s.font = SUB_FONT
    s.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 18


def notes_block(ws, start_row, notes, ncols):
    r = start_row + 1
    for note in notes:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
        c = ws.cell(row=r, column=1, value=note)
        c.font = NOTE_FONT; c.alignment = LEFT
        r += 1
    return r


def diag_lines(ws, r, lines, ncols, height=20):
    for i, s in enumerate(lines):
        ws.merge_cells(start_row=r + i, start_column=1, end_row=r + i, end_column=ncols)
        cell = ws.cell(row=r + i, column=1, value=s)
        cell.font = CELL_FONT; cell.alignment = LEFT; cell.border = BORDER
        ws.row_dimensions[r + i].height = height
    return r + len(lines) + 1


wb = Workbook()

# ============================================================ 0) 표지/요약
ws = wb.active
ws.title = "요약(Cover)"
ws.sheet_view.showGridLines = False
NC = 4
title_block(ws, "거시경제 상황 & 주가 영향 레포트", "데이터 기준 2026년 7월 초 · 작성일 2026-07-03 · 목적: 지표→주가 시사점 도출", NC)

r = 4
r = section_bar(ws, r, "■ 리포트 개요", NC)
overview = [
    ["구분", "내용"],
    ["작성 목적", "국내외 경제 지표를 정리하고, 각 항목이 '주가(증시)'에 미치는 영향과 결론 도출"],
    ["구성", "①미시 ②거시 ③국제 ④국내기관전망 ⑤주가전망 (+요약·출처)"],
    ["항목별 구성", "각 항목 = 데이터표 + [설명·결론·주가영향] 콜아웃(수혜/피해 업종 명시)"],
    ["데이터 기준", "2026년 6월 발표치 및 주요 기관 전망(2026-07-03 취합)"],
    ["참고 레포트", "국내 15개+ 기관(기재부·한은·KDI·KIET·KIEP·KIF·KCIF·KITA·HRI·삼일PwC 등)"],
]
for i, row in enumerate(overview):
    ws.merge_cells(start_row=r + i, start_column=2, end_row=r + i, end_column=4)
    for j, val in enumerate(row, start=1):
        cell = ws.cell(row=r + i, column=j, value=val)
        cell.font = HDR_FONT if i == 0 else CELL_FONT
        cell.alignment = LEFT; cell.border = BORDER
        if i == 0:
            cell.fill = fill(LTBLUE)
    ws.row_dimensions[r + i].height = 18
r += len(overview) + 1

r = section_bar(ws, r, "■ 한눈에 보는 핵심 지표 (2026년 6월)", NC)
kpi = [
    ["지표", "한국", "미국", "1차 주가 함의"],
    ["정책금리", "2.50%", "3.50~3.75%", "고금리 지속 → 성장·고밸류주 부담"],
    ["소비자물가", "2.6%(4월)", "4.2%(5월)", "인플레 재점화 → 밸류에이션 압박"],
    ["성장률('26)", "2.6%", "2.2%", "반도체가 韓 지수 하단 방어"],
    ["환율·달러", "USD/KRW≈1,554", "DXY≈100.7", "외국인 코스피 순매도 압력"],
    ["증시(YTD)", "코스피 변동성↑", "S&P500 +9.6%", "AI가 상단, 유가·금리가 하단"],
]
for i, row in enumerate(kpi):
    for j, val in enumerate(row, start=1):
        cell = ws.cell(row=r + i, column=j, value=val)
        cell.border = BORDER
        cell.alignment = CENTER if i == 0 or (j > 1 and j < 4) else LEFT
        cell.font = HDR_FONT if i == 0 else CELL_FONT
        if i == 0:
            cell.fill = fill(LTBLUE)
        elif i % 2 == 0:
            cell.fill = fill(STRIPE)
    ws.row_dimensions[r + i].height = 17
r += len(kpi) + 1

r = section_bar(ws, r, "■ 최종 결론 3줄 (주가 관점)", NC)
r = diag_lines(ws, r, [
    "1) 국면: 유가 급등·고금리·지정학이 지수 상단을 누르고, AI·반도체 실적이 하단을 받치는 '변동성 큰 박스권'.",
    "2) 전략: 지수(베타)보다 종목·업종 차별화(스톡피커) 장세 — 반도체·에너지·방산·조선 우위, 항공·정유화학·건설·내수 열위.",
    "3) 변수: 호르무즈 정상화 시 유가↓·금리인하 재개 → 강세 전환 / 전쟁 확전·유가 $120+ 시 스태그플레이션 → 조정.",
], NC, height=30)

for col, w in zip("ABCD", [16, 22, 22, 34]):
    ws.column_dimensions[col].width = w
notes_block(ws, r, [
    "※ 본 자료는 공개자료 요약·해석이며 투자자문·매매권유가 아닙니다. 주가 영향은 일반적 시장 메커니즘 기반 시나리오입니다.",
    "※ 개별 종목명은 대표 '업종/테마' 예시이며 특정 종목 추천이 아닙니다.",
], NC)

# ============================================================ 1) 미시경제
ws1 = wb.create_sheet("미시경제")
ws1.sheet_view.showGridLines = False
NC = 5
W1 = 22 + 10 + 16 + 22 + 26
title_block(ws1, "① 미시경제 (Microeconomics)", "가계·노동·산업·부문 세부 + 항목별 주가 영향 · 2026년 6월", NC)

r = 4
r = section_bar(ws1, r, "■ 가계·소비 (Household & Consumption)", NC)
r = write_table(ws1, r,
    ["지표", "국가", "값", "전기/전년비", "비고"],
    [
        ["시간당 평균임금(AHE)", "미국", "$37.64", "+0.3% MoM / +3.5% YoY", "6월 민간"],
        ["실질임금", "미국", "둔화", "명목 3.5% < CPI 4.2%", "구매력 감소"],
        ["가계부채", "한국", "높은 수준", "GDP 대비 89.3%", "금리인하 제약"],
        ["소비심리", "한국", "회복 둔화", "-", "고물가·고환율"],
    ], col_widths=[22, 10, 16, 22, 26])
r = callout(ws1, r, [
    ("설명", "임금 상승률(3.5%)이 물가(4.2%)를 밑돌아 실질구매력이 감소. 韓은 가계부채·고물가로 소비 여력이 제한됨."),
    ("결론", "내수 소비 모멘텀 약화 → 소비 회복은 완만·선별적. 필수소비 > 임의소비 구도."),
    ("주가", "내수유통·백화점·의류·외식 ▼ / 저가·가성비 소비주 상대 우위 / 필수소비(식품·생필품) ▲ 방어적."),
], NC, W1)

r = section_bar(ws1, r, "■ 노동시장 세부 (美 2026.06)", NC)
r = write_table(ws1, r,
    ["항목", "값", "방향", "주가 영향", "코멘트"],
    [
        ["비농업 신규고용", "+57천명", "▼ 둔화", "▼ 경기민감주 부담", "예상치 하회"],
        ["실업률", "4.2%", "▼ 12개월 최저", "중립", "고용 견조 신호 혼재"],
        ["경제활동참가율", "61.5%", "▼", "중립", "'21.3 이후 최저"],
        ["레저·접객 고용", "-61천명", "▼", "▼ 여행·호텔·카지노", "월드컵·계절요인"],
        ["전문·헬스케어 고용", "+36/+22천명", "▲", "▲ 헬스케어·서비스", "고용 증가 주도"],
    ], col_widths=[20, 12, 14, 20, 22], dir_col=3)
r = callout(ws1, r, [
    ("설명", "고용 둔화(+57천명)는 경기 둔화 신호이나, 그만큼 연준의 추가 긴축 명분을 약화시켜 '금리 기대'엔 완충."),
    ("결론", "'나쁜 지표=금리인하 기대'와 '경기 둔화=실적 우려'가 상충 → 지표 발표마다 변동성 확대."),
    ("주가", "여행·항공·레저 ▼ / 헬스케어·필수서비스 ▲ / 지수 전반은 방향성보다 변동성 국면."),
], NC, W1)

r = section_bar(ws1, r, "■ 산업·부문 (한국)", NC)
r = write_table(ws1, r,
    ["부문", "상태", "주가 영향", "근거/리스크", ""],
    [
        ["반도체/IT", "호조", "▲ 반도체·장비·소부장", "AI 수요·수출 사상최고", ""],
        ["조선·바이오헬스", "호조", "▲ 조선·바이오", "견고한 성장(KIET)", ""],
        ["철강·석유화학·정유", "침체", "▼ 소재·정유주", "가동률↓·마진 압박", ""],
        ["자동차·섬유", "정체", "중립~▼ 완성차·부품", "성장 정체", ""],
        ["건설·부동산 PF", "부실 우려", "▼ 건설·증권·저축은행", "PF 건전성 리스크", ""],
    ], col_widths=[16, 12, 22, 22, 4], dir_col=2)
r = callout(ws1, r, [
    ("설명", "반도체·AI가 성장·수출·이익을 독식하는 '외끌이' 구조. 소재·건설은 침체로 산업 양극화 심화(KIET·KDI)."),
    ("결론", "지수 상승은 반도체 편중 → 반도체 빠진 '지수 착시' 주의. 소재·건설은 실적 바닥 확인 전까지 비중 축소."),
    ("주가", "반도체·장비·소부장·조선·바이오 ▲ / 정유·화학·철강·건설·증권(부동산 익스포저) ▼ / 완성차 중립."),
], NC, W1)

r = section_bar(ws1, r, "■ 반도체 심층 (KITA·KIET)", NC)
r = write_table(ws1, r,
    ["항목", "값/내용", "출처", "주가 함의", ""],
    [
        ["'25 반도체 수출", "사상최고 ≈1,419억달러", "무역협회", "▲ 실적 서프라이즈 기대", ""],
        ["글로벌 반도체시장'26", "+17.8% → 9,098억달러", "무역협회", "▲ 업황 확장 국면", ""],
        ["수급", "범용DRAM 부족+AI수요 급증", "KITA", "▲ 메모리 가격 상승", ""],
        ["구조", "고부가(HBM)·AI 중심 재편", "KITA", "▲ HBM·장비 밸류 재평가", ""],
    ], col_widths=[20, 26, 14, 24, 4], dir_col=3)
r = callout(ws1, r, [
    ("설명", "AI 투자로 범용 DRAM 공급이 부족해지며 메모리 가격·믹스 개선. HBM 등 고부가 비중 확대로 이익률 레벨업."),
    ("결론", "2026년 국내 증시의 핵심 상승 동력. 다만 이미 주가에 상당폭 선반영 → 실적 '눈높이'와 밸류에이션 부담 점검 필요."),
    ("주가", "메모리 대형주·HBM·반도체 장비·소재 ▲(핵심 주도주) / 단, 고밸류·기대 과열 시 되돌림 변동성 유의."),
], NC, W1)

r = section_bar(ws1, r, "■ 가계부채·부동산 (금융위·우리금융)", NC)
r = write_table(ws1, r,
    ["항목", "값/내용", "주가 영향", "코멘트", ""],
    [
        ["GDP 대비 가계부채", "89.3%('25 3Q)", "중립~▼ 은행 성장성", "'21 98.7%→하향"],
        ["국제비교", "주요국 中 5위", "▼ 소비 여력 제약", "여전히 높음"],
        ["가계부채 관리방안", "관리 강화(2026.04)", "▼ 대출성장 둔화", "생산적 자금 유도"],
        ["다주택 규제", "규제지역 만기연장 불허", "▼ 건설·부동산·리츠", "양도세 중과 재개"],
    ], col_widths=[18, 22, 24, 20, 4], dir_col=2)
r = callout(ws1, r, [
    ("설명", "부채는 완만히 줄지만 여전히 높아 소비를 짓누름. 규제 강화로 은행 대출성장·건설 분양은 둔화."),
    ("결론", "은행은 '대출성장↓ vs 건전성 관리·배당' 혼재 → 고배당 방어주 성격. 건설·부동산은 하방 우위."),
    ("주가", "은행 중립(배당 매력) / 건설·부동산·증권·리츠 ▼ / 내수 소비주 ▼."),
], NC, W1)

notes_block(ws1, r, [
    "※ 미시 = 가계·기업·부문 등 개별 주체 수준. 韓 부문·반도체는 KIET·KITA·금융위·우리금융 취합.",
    "※ '주가 영향'은 대표 업종/테마 기준 방향성이며 특정 종목 추천이 아닙니다.",
], NC)

# ============================================================ 2) 거시경제
ws2 = wb.create_sheet("거시경제")
ws2.sheet_view.showGridLines = False
NC = 6
W2 = 18 + 16 + 14 + 12 + 12 + 22
title_block(ws2, "② 거시경제 (Macroeconomics)", "성장·물가·금리·고용 총량지표 + 항목별 주가 영향 · 2026년 6월", NC)

r = 4
r = section_bar(ws2, r, "■ 한국 (Korea)", NC)
r = write_table(ws2, r,
    ["지표", "최근치", "전망'26", "전망'27", "주가 영향", "출처"],
    [
        ["기준금리", "2.50% 동결", "인하 지연", "-", "▼ 성장주 밸류 부담", "한국은행"],
        ["소비자물가", "2.6%(4월)", "2.7%", "2.3%", "▼ 금리인하 제약", "유가 상방"],
        ["실질GDP", "-", "2.6%", "2.1%", "▲ 반도체가 지수 방어", "한은 상향"],
        ["경상수지", "GDP 5.9%", "흑자 지속", "-", "▲ 수출주 실적·원화 지지", "반도체 주도"],
    ], col_widths=[16, 14, 12, 10, 24, 14], dir_col=4)
r = callout(ws2, r, [
    ("설명", "물가 반등으로 한은의 금리인하가 지연·불투명. 성장은 반도체 수출이 견인해 지수 하단을 방어."),
    ("결론", "'금리 부담(밸류에이션) vs 반도체 실적(EPS)'의 줄다리기. 코스피는 반도체 이익 개선에 연동된 종목 장세."),
    ("주가", "반도체·수출 대형주 ▲ / 금리 민감 성장·바이오·부동산 ▼ / 지수는 반도체 비중효과로 견조."),
], NC, W2)

r = section_bar(ws2, r, "■ 미국 (United States)", NC)
r = write_table(ws2, r,
    ["지표", "최근치", "전망'26", "기타", "주가 영향", "출처"],
    [
        ["연방기금금리", "3.50~3.75%", "연말 3.6~4.1%", "매파적", "▼ 고밸류 기술주 압박", "FOMC"],
        ["CPI", "4.2%(5월)", "3.6%", "다년 최고", "▼ 변동성 확대", "근원 2.9%"],
        ["실질GDP", "-", "2.2%", "-0.2%p", "▼ 경기민감주 부담", "연준 하향"],
        ["실업률", "4.2%(6월)", "4.3%", "-0.1%p", "중립", "고용 둔화"],
    ], col_widths=[16, 14, 14, 12, 24, 14], dir_col=4)
r = callout(ws2, r, [
    ("설명", "유가발 물가로 연준이 '인상 가능성'까지 시사(매파 전환). 고금리는 할인율↑로 고밸류 성장주의 밸류에이션을 압박."),
    ("결론", "그럼에도 S&P는 AI 실적으로 연 +9.6% — '금리 역풍 vs AI 이익'의 힘겨루기. 밸류 부담 큰 구간은 되돌림 위험."),
    ("주가", "AI·반도체 실적주 ▲(단 변동성) / 고PER·무이익 성장주·리츠·유틸(금리) ▼ / 에너지·방산 ▲."),
], NC, W2)

r = section_bar(ws2, r, "■ 국내기관 '26 성장률 컨센서스", NC)
r = write_table(ws2, r,
    ["기관", "성장률", "물가", "발표", "특징", ""],
    [
        ["기획재정부", "2.0%(목표)", "-", "'26상반기", "잠재성장 반등 목표"],
        ["한국은행", "1.8→2.6%", "2.7%", "'26상반기", "반도체 반영 상향"],
        ["KDI", "2.5%", "-", "2026.05", "설비투자 3.3%"],
        ["산업연구원", "2.5%", "-", "'26", "AI·반도체 강세"],
        ["한국금융연구원", "2.1%", "-", "'26", "금융완화 효과"],
        ["현대경제硏/NABO", "1.9%", "1.9%", "'25.9/'26", "대외리스크 중시"],
    ], col_widths=[18, 14, 10, 12, 20, 4])
r = callout(ws2, r, [
    ("설명", "전망 편차(1.8~2.6%)는 발표시점 차이 — 2025년말 전망은 관세·수출둔화로 보수적, '26상반기 갱신치는 반도체 호조 반영."),
    ("결론", "'반도체 낙관파(2.5%) vs 대외리스크 신중파(1.9%)'로 양분. 반도체 사이클 지속 여부가 지수 방향을 가름."),
    ("주가", "반도체 사이클 유지 시 코스피 레벨업 / 사이클 정점·기저효과 우려 부각 시 지수 조정 압력."),
], NC, W2)

r = section_bar(ws2, r, "■ 거시 국면 진단 & 스태그플레이션 점검", NC)
r = diag_lines(ws2, r, [
    "· 물가↑ + 성장↓ 조합 = 스태그플레이션 우려 → 전반적 밸류에이션 하락 압력, 방어주·실물자산 상대 강세.",
    "· 통화정책: 韓·美 모두 인하 지연/매파 → 유동성 장세 부재, '실적 있는 종목'만 오르는 차별화.",
    "· 주가 결론: 지수 방향성은 제한적, 섹터·종목 알파가 수익률을 좌우. 현금·방어 비중 확보로 변동성 대응.",
], NC, height=28)
notes_block(ws2, r, [
    "※ 거시 = 경제 전체 총량(성장·물가·고용·금리). 전망치는 한은·연준 및 국내 주요 연구기관 공식 전망.",
], NC)

# ============================================================ 3) 국제경제
ws3 = wb.create_sheet("국제경제")
ws3.sheet_view.showGridLines = False
NC = 5
W3 = 18 + 18 + 12 + 20 + 22
title_block(ws3, "③ 국제경제 (International)", "세계성장·원자재·외환·주요국 + 항목별 주가 영향 · 2026년 6월", NC)

r = 4
r = section_bar(ws3, r, "■ 세계 성장 (Global Growth)", NC)
r = write_table(ws3, r,
    ["지역/기관", "성장'26", "주가 영향", "비고", ""],
    [
        ["세계(KIEP)", "3.0%", "중립", "'완충된 둔화'", ""],
        ["미국", "2.2%(연준)/1.6%(KIEP)", "중립~▼", "AI가 상쇄", ""],
        ["유로존", "1.1%", "▼ 유럽주 부진", "에너지 압박", ""],
        ["중국", "5.0%(1Q)", "▲ 대중 수출·소재주", "목표 4.5~5%", ""],
    ], col_widths=[18, 20, 20, 16, 4], dir_col=2)
r = callout(ws3, r, [
    ("설명", "선진국 저성장(유럽 1.1%·일본 0.6%) 속 중국이 5%대 성장을 유지. 세계 교역은 완만한 둔화."),
    ("결론", "글로벌 경기 모멘텀 약화로 지수 상단 제한. 중국 회복은 한국 중간재·소재·화장품·여행주엔 국지적 호재."),
    ("주가", "유럽 익스포저주 ▼ / 중국 소비·소재·면세·화장품 ▲ / 글로벌 경기민감주(산업재) 중립~▼."),
], NC, W3)

r = section_bar(ws3, r, "■ 원자재·에너지 (유가)", NC)
r = write_table(ws3, r,
    ["품목", "최근가", "전망'26평균", "주가 영향", "비고"],
    [
        ["Brent", "≈$95.06", "$82~85", "▲ 에너지 / ▼ 항공·화학", "호르무즈 봉쇄"],
        ["WTI", "≈$92.32", "'26말 $90~100(서베이)", "▲ 정유(정제마진 변수)", "美·이란 전쟁"],
        ["시나리오", "Q2≈$106", "Q4≈$89→연말$70", "유가↓ 시 증시 전반 ▲", "해협 정상화 가정"],
        ["상방리스크", ">$120", "봉쇄 3Q 지속시", "▼ 스태그플레이션 조정", "GS"],
    ], col_widths=[16, 16, 20, 24, 20], dir_col=3)
r = callout(ws3, r, [
    ("설명", "호르무즈 봉쇄로 유가 급등 → 인플레·금리 상방의 '진앙'. 유가는 이번 국면 증시의 최대 스윙 변수."),
    ("결론", "유가 방향 = 증시 방향의 역(逆). 유가 안정 시 물가·금리 완화로 강세 전환, 재급등 시 조정."),
    ("주가", "에너지·정유·조선(해양) ▲ / 항공·해운·석유화학·전력·유틸 ▼ / 유가 하락 반전 시 성장주 전반 ▲."),
], NC, W3)

r = section_bar(ws3, r, "■ 외환·금융시장 (FX & Markets)", NC)
r = write_table(ws3, r,
    ["지표", "값", "주가 영향", "비고", ""],
    [
        ["달러인덱스(DXY)", "≈100.7", "▼ 신흥국 증시 자금유출", "'25.5 이후 최고"],
        ["USD/KRW", "≈1,554", "▼ 외국인 코스피 순매도", "환손실 우려"],
        ["원화 약세 효과", "수출단가", "▲ 수출주 원화환산 실적", "IT·車·조선 수혜"],
        ["S&P500", "+9.6% YTD", "▲(AI) 단 6월 -1%", "밸류 부담"],
        ["나스닥", "3월말+21%", "▲ 변동성 확대", "AI 주도"],
    ], col_widths=[18, 20, 24, 18, 4], dir_col=2)
r = callout(ws3, r, [
    ("설명", "달러 강세·원화 약세는 '양날의 칼' — 수출기업 원화 실적은 늘지만, 외국인은 환차손 탓에 코스피를 순매도."),
    ("결론", "고환율 국면에선 외국인 수급이 지수 발목. 원화 실적 개선 수혜(수출주)와 수급 악재(외국인 매도)가 공존."),
    ("주가", "수출 대형주(반도체·車·조선) 원화실적 ▲ / 외국인 수급 의존 高 종목·내수·항공(원가) ▼ / 지수는 수급 부담."),
], NC, W3)

r = section_bar(ws3, r, "■ 국내기관 세계진단 & 3대 리스크 (KIEP·KCIF·KITA)", NC)
r = write_table(ws3, r,
    ["항목", "내용", "출처", "주가 함의", ""],
    [
        ["세계성장'26", "3.0%, '완충된 둔화'", "KIEP", "중립(상단 제한)"],
        ["원/달러", "달러약세에도 원화강세 제한", "KCIF", "▼ 외국인 수급"],
        ["韓 수출'26", "-0.5%, 무역흑자 675억달러", "KITA", "중립(반도체 편중)"],
        ["3대 리스크", "중동에너지·美통상·재정/국채", "KCIF·KIEP", "▼ 변동성 확대 요인"],
    ], col_widths=[16, 26, 14, 18, 4], dir_col=3)
r = callout(ws3, r, [
    ("설명", "국내기관 공통으로 ①중동 에너지충격 ②美 관세·통상 불확실성 ③주요국 재정·국채 불안을 3대 하방리스크로 지목."),
    ("결론", "세 리스크 모두 '변동성 확대' 방향 → 위험자산 프리미엄 요구 상승, 지수 상단 제한·방어적 대응 유효."),
    ("주가", "방산·에너지·안전자산(금 관련) ▲ / 관세 노출 수출·경기민감주 변동성 / 고부채·고밸류주 ▼."),
], NC, W3)

notes_block(ws3, r, [
    "※ 국제 = 세계성장·원자재·외환 등 대외환경. 유가·환율은 변동성 큰 시점 스냅샷.",
    "※ 세계진단 = KIEP·국제금융센터(KCIF)·무역협회(KITA) 취합.",
], NC)

# ============================================================ 4) 국내기관전망
wsR = wb.create_sheet("국내기관전망")
wsR.sheet_view.showGridLines = False
NC = 6
WR = 22 + 12 + 8 + 40 + 12 + 10
title_block(wsR, "④ 국내 주요기관 레포트 종합", "15개+ 기관 전망 취합 + 주가 시사점 · 2025년말~2026년 상반기", NC)

r = 4
r = section_bar(wsR, r, "■ 기관별 2026 한국경제 전망 비교", NC)
r = write_table(wsR, r,
    ["기관", "성장률", "물가", "핵심 메시지", "발표", "구분"],
    [
        ["기획재정부", "2.0%(목표)", "-", "'경제大도약 원년', 잠재성장 반등", "'26상반", "정부"],
        ["한국은행", "1.8→2.6%", "2.7%", "반도체로 상향, 금리 2.50% 동결", "'26상반", "중앙은행"],
        ["KDI", "2.5%", "-", "반도체+내수, 설비투자 3.3%·소비 2.2%", "'26.05", "국책"],
        ["산업연구원(KIET)", "2.5%", "-", "AI·반도체 강세, 산업 양극화", "'26", "국책"],
        ["대외경제정책연(KIEP)", "-", "-", "세계 3.0%, '완충된 둔화'", "'26", "국책"],
        ["한국금융연구원(KIF)", "2.1%", "-", "금융완화·정책효과", "'26", "연구원"],
        ["삼일PwC", "2% 내외", "-", "재정 내수가 수출둔화 상쇄, 저성장", "'25.12", "민간"],
        ["삼정KPMG", "저성장", "-", "국내 경제·산업 전망", "'25.12", "민간"],
        ["현대경제연구원(HRI)", "1.9%", "1.9%", "회복 미약, 실업률 3.0%", "'25.09", "민간"],
        ["국회예산정책처(NABO)", "1.9%", "-", "독립 재정기구 전망", "'26", "국회"],
        ["한국무역협회(KITA)", "-", "-", "수출 -0.5%, 흑자 675억달러", "'25.12", "협회"],
        ["국제금융센터(KCIF)", "-", "-", "글로벌 3.1→2.9%, 원화강세 제한", "'26", "연구원"],
        ["하나/우리금융연구소", "-", "-", "고환율·가계부채(GDP대비 92%)", "'26", "민간"],
    ], col_widths=[22, 12, 8, 40, 12, 10])
r = callout(wsR, r, [
    ("설명", "성장률 컨센서스는 잠재성장(약 2%) 부근. 반도체 낙관파(KDI·KIET 2.5%)와 대외리스크 신중파(HRI·NABO 1.9%)로 갈림."),
    ("결론", "기관 공통: ①반도체가 성장·증시의 엔진 ②물가·유가·환율·가계부채가 하방 ③정책은 가계부채 관리·재정 내수 보강."),
    ("주가", "컨센서스 상향(반도체 반영)은 코스피 우호적 / 단, 대외 3대 리스크가 밸류에이션 상단을 제한."),
], NC, WR)

r = section_bar(wsR, r, "■ 국내기관 공통 진단 (Consensus Themes)", NC)
r = diag_lines(wsR, r, [
    "1) 성장: 반도체 사이클 낙관(2.5%) vs 대외리스크 신중(1.9%)으로 양분 → 반도체 지속성이 지수 방향 결정.",
    "2) 물가: 목표(2%) 부근 안정 전망이 다수였으나, 중동발 유가 급등으로 상방 위험 재부각(한은 2.7%).",
    "3) 수출: 반도체·AI가 핵심 엔진이나, 기저효과·교역둔화로 총수출은 소폭 감소(-0.5%, KITA) 전망.",
    "4) 리스크: 중동 에너지·美 통상·가계부채/부동산 PF·고환율 → 국내기관 공통 하방 요인.",
    "5) 정책: 가계부채 관리 강화(금융위)·재정 내수 보강(기재부), 통화정책은 물가·환율에 제약.",
], NC, height=28)
notes_block(wsR, r, [
    "※ 성장률 편차는 발표시점(2025하반 vs 2026상반)·반도체/대외리스크 가정 차이. '저성장' 등은 서술형 전망.",
], NC)

# ============================================================ 5) 주가전망 (신규)
wsS = wb.create_sheet("주가전망")
wsS.sheet_view.showGridLines = False
NC = 5
WS = 20 + 12 + 34 + 26 + 8
title_block(wsS, "⑤ 주가전망 (Stock Outlook)", "지표 종합 → 증시 방향·섹터 승패·시나리오 · 2026년 6월 기준", NC)

r = 4
r = section_bar(wsS, r, "■ 종합 시장 View", NC)
r = diag_lines(wsS, r, [
    "· 큰 그림: '유가·고금리·지정학(상단 억제)' vs 'AI·반도체 실적(하단 지지)' → 방향성 약한 '변동성 큰 박스권'.",
    "· 성격: 유동성(금리인하) 장세 부재 → 지수(베타)보다 실적 있는 종목·업종의 차별화(스톡피커) 장세.",
    "· 한국: 반도체가 지수를 방어하나 외국인 수급(고환율)이 발목 → 반도체 편중·양극화 심화.",
    "· 미국: AI가 지수를 끌지만 고금리로 고밸류주 되돌림 반복 → 실적·현금흐름 우량주 우위.",
], NC, height=28)

r = section_bar(wsS, r, "■ 섹터별 주가 방향 (Winners & Losers)", NC)
r = write_table(wsS, r,
    ["섹터/테마", "방향", "근거", "대표 업종", ""],
    [
        ["반도체·AI·HBM", "▲ 강세", "AI 수요·메모리 업황·실적 서프라이즈", "메모리·장비·소부장", ""],
        ["에너지", "▲ 강세", "유가 급등 수혜", "정유(원유재고)·가스·에너지", ""],
        ["방산", "▲ 강세", "지정학 리스크·수출 확대", "방위산업", ""],
        ["조선", "▲ 우위", "수주 사이클·해양", "조선·기자재", ""],
        ["바이오·헬스케어", "▲ 우위", "방어적·고용 견조", "제약·바이오·의료", ""],
        ["필수소비(방어)", "▲ 상대우위", "경기·물가 방어", "식품·생필품·통신", ""],
        ["은행", "중립", "예대마진 유지 vs 대출성장·건전성", "은행(고배당)", ""],
        ["수출 대형주", "중립", "원화실적 ▲ vs 외국인수급 ▼", "반도체·車·조선", ""],
        ["항공·해운·운송", "▼ 약세", "유가(연료비) 부담", "항공·해운", ""],
        ["정유화학·철강", "▼ 약세", "마진 압박·수요 둔화", "석유화학·철강", ""],
        ["건설·부동산·증권", "▼ 약세", "PF 부실·규제·거래 위축", "건설·리츠·증권", ""],
        ["내수유통·레저", "▼ 약세", "실질구매력·소비심리 위축", "유통·백화점·여행", ""],
        ["고밸류 성장주·유틸", "▼ 약세", "고금리 할인율 부담", "고PER 기술·리츠·유틸", ""],
    ], col_widths=[20, 12, 34, 26, 4], dir_col=1)
r = callout(wsS, r, [
    ("설명", "이번 국면의 이익·정책 흐름을 업종에 매핑한 결과. 유가·금리·지정학·반도체가 승패를 가르는 4대 축."),
    ("결론", "'반도체·에너지·방산·조선·바이오' 우위, '항공·정유화학·건설·내수·고밸류' 열위의 뚜렷한 차별화."),
    ("주가", "코어(방어+반도체) + 위성(에너지·방산)의 바벨 전략 유효 / 고밸류·경기민감·유가피해 업종 비중 축소."),
], NC, WS)

r = section_bar(wsS, r, "■ 시나리오별 증시 경로 (Scenario)", NC)
r = write_table(wsS, r,
    ["시나리오", "핵심 조건", "유가/금리", "증시 방향", "가능성"],
    [
        ["Bull(강세)", "호르무즈 정상화·유가 하락", "유가↓·금리인하 재개", "▲ 강세 전환(반도체+멀티플)", "중"],
        ["Base(기본)", "유가 고착·고금리 지속", "유가 $85±·동결", "→ 박스권·종목 차별화", "높음"],
        ["Bear(약세)", "전쟁 확전·유가 $120+", "유가↑·추가 긴축", "▼ 스태그플레이션 조정", "중하"],
    ], col_widths=[14, 24, 22, 30, 10], dir_col=3)
r = callout(wsS, r, [
    ("설명", "증시 경로는 사실상 '유가·금리' 두 변수로 압축. 유가가 물가→금리→밸류에이션의 방아쇠."),
    ("결론", "기본(Base)은 박스권 종목 장세. 관전 포인트는 '호르무즈 재개 여부'와 '연준의 인하 재개 시점'."),
    ("주가", "Bull=성장·반도체 전반 ▲ / Base=반도체·방어주 선별 ▲ / Bear=에너지·방산·현금 외 전반 ▼."),
], NC, WS)

r = section_bar(wsS, r, "■ 한국 vs 미국 vs 신흥국 (상대 매력)", NC)
r = write_table(wsS, r,
    ["시장", "평가", "포인트", "리스크", ""],
    [
        ["미국(S&P·나스닥)", "우위", "AI 실적·기업 이익 견조", "고밸류·고금리 되돌림", ""],
        ["한국(코스피)", "선별", "반도체 이익·저평가 매력", "외국인 수급(고환율)·양극화", ""],
        ["신흥국(EM)", "열위", "성장 여지", "달러강세發 자금유출", ""],
        ["유럽", "열위", "밸류 저평가", "에너지·저성장", ""],
    ], col_widths=[18, 10, 30, 26, 4])
r = callout(wsS, r, [
    ("설명", "달러 강세 국면에선 미국>한국>신흥국 순의 상대 우위가 일반적. 한국은 반도체라는 강력한 개별 모멘텀 보유."),
    ("결론", "美 AI 실적주 코어 + 韓 반도체 위성의 조합. 신흥국·유럽은 달러·에너지 부담으로 후순위."),
    ("주가", "美 AI·우량주 ▲ / 韓 반도체 선별 ▲ / EM·유럽 상대 ▼(달러 약세 전환 시 반전 가능)."),
], NC, WS)

r = section_bar(wsS, r, "■ 최종 결론 (Bottom Line)", NC)
r = diag_lines(wsS, r, [
    "1) 방향: 지수는 방향성 약한 박스권 — 상단은 유가·금리, 하단은 AI·반도체 실적이 규정.",
    "2) 전략: 종목 차별화 장세. 반도체·에너지·방산·조선·바이오 비중 확대, 항공·정유화학·건설·내수 축소.",
    "3) 트리거: 유가(호르무즈)와 연준 인하 시점이 최대 변수 — 유가 안정+인하 재개가 강세 전환의 열쇠.",
    "4) 리스크관리: 스태그플레이션 시나리오 대비 방어주·현금·에너지로 변동성 헤지.",
], NC, height=30)
notes_block(wsS, r, [
    "※ 본 주가전망은 거시·미시·국제 지표의 일반적 시장 파급을 정리한 시나리오이며, 투자자문·매매권유가 아닙니다.",
    "※ 개별 종목이 아닌 업종/테마 방향성 기준이며, 실제 성과는 개별 실적·수급·정책에 따라 달라질 수 있습니다.",
], NC)

# ============================================================ 6) 출처
ws4 = wb.create_sheet("출처(Sources)")
ws4.sheet_view.showGridLines = False
NC = 3
title_block(ws4, "출처 및 유의사항 (Sources & Notes)", "2026-07-03 웹 취합 · 참고용", NC)
r = 4
r = write_table(ws4, r,
    ["구분", "기관/자료", "내용"],
    [
        ["국내(정부)", "기획재정부(재정경제부)", "2026년 경제성장전략"],
        ["국내(정부)", "금융위원회(FSC)", "2026년 가계부채 관리방안(2026.04.01)"],
        ["국내(중앙은행)", "한국은행(BOK)", "2026 경제전망보고서·경제통계·기준금리"],
        ["국내(국책)", "KDI 한국개발연구원", "KDI 경제전망 2026 상반기(2026.05.13)·경제동향"],
        ["국내(국책)", "산업연구원(KIET)", "2026년 경제·산업 전망"],
        ["국내(국책)", "대외경제정책연구원(KIEP)", "2026년 세계경제 전망(업데이트)"],
        ["국내(국회)", "국회예산정책처(NABO)", "2026 경제전망(성장률 1.9%)"],
        ["국내(연구원)", "한국금융연구원(KIF)", "2026년 경제전망과 정책적 시사점"],
        ["국내(연구원)", "국제금융센터(KCIF)", "2026 세계경제·국제금융시장 이슈 및 전망"],
        ["국내(연구원)", "자본시장연구원(KCMI)", "2026년 거시경제 주요 이슈"],
        ["국내(협회)", "한국무역협회(KITA)", "2025 수출입 평가·2026 전망·반도체 브리프"],
        ["국내(민간)", "현대경제연구원(HRI)", "2026년 한국경제 전망(통권 996호)"],
        ["국내(민간)", "삼일PwC / 삼정KPMG", "2026년 국내외 경제·산업 전망"],
        ["국내(민간)", "하나금융연구소/우리금융경영연구소", "외환시장 전망·가계부채 분석"],
        ["해외", "美 연준(FOMC)/BLS", "2026.06.17 FOMC·6월 고용상황"],
        ["해외", "IMF/EIA/World Bank/GS", "WEO·STEO·원자재 전망"],
        ["해외", "Eurostat/中 NBS", "유로존·중국 GDP"],
        ["시장", "CNBC/CNN/Trading Economics 등", "증시·환율·유가 시세(취합)"],
    ], col_widths=[16, 34, 46])
notes_block(ws4, r, [
    "※ 본 레포트는 공개 자료를 요약·재구성한 참고용 문서이며, 투자자문·매매권유가 아닙니다.",
    "※ '주가 영향/전망'은 지표의 일반적 시장 파급을 정리한 시나리오 해석이며, 특정 종목 추천이 아닙니다.",
    "※ 수치는 발표기관·시점·개정에 따라 달라질 수 있으며, 정확한 값은 원출처를 확인하십시오.",
    "※ 작성 2026-07-03 / 데이터 기준 2026년 6월(일부 잠정·전망 포함).",
], NC)

fname = "거시경제_주가영향_레포트_2026-07.xlsx"
wb.save(fname)
print("saved:", fname)
for s in wb.sheetnames:
    print(" -", s)
