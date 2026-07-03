# -*- coding: utf-8 -*-
"""거시경제 상황 요약 레포트(엑셀) 생성 스크립트.

데이터 기준: 2026년 7월 초 (2026년 6월 발표치·전망치 중심)
출처: 공개 자료(연준/BLS, 한국은행/KDI, IMF/EIA, Eurostat, 중국 NBS 등) 웹 검색 취합.
일부 수치는 잠정치 또는 기관 전망치이며, 시트 하단 주석·출처 참조.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- 색상/스타일
NAVY = "1F3864"       # 제목 배경
BLUE = "2E5496"       # 섹션 헤더
LTBLUE = "D6E0F0"     # 표 헤더
STRIPE = "F2F5FB"     # 줄무늬
GREY = "808080"
GREEN = "548235"
RED = "C00000"

WHITE_BOLD = Font(name="맑은 고딕", color="FFFFFF", bold=True)
TITLE_FONT = Font(name="맑은 고딕", color="FFFFFF", bold=True, size=16)
SUB_FONT = Font(name="맑은 고딕", color="FFFFFF", size=10)
HDR_FONT = Font(name="맑은 고딕", bold=True, size=10)
CELL_FONT = Font(name="맑은 고딕", size=10)
NOTE_FONT = Font(name="맑은 고딕", size=9, italic=True, color=GREY)
SECTION_FONT = Font(name="맑은 고딕", color="FFFFFF", bold=True, size=11)

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")


def fill(color):
    return PatternFill("solid", fgColor=color)


def style_header_row(ws, row, ncols, start_col=1):
    for c in range(start_col, start_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill(LTBLUE)
        cell.font = HDR_FONT
        cell.alignment = CENTER
        cell.border = BORDER


def write_table(ws, start_row, headers, rows, col_widths=None, value_align=None):
    """헤더 + 데이터 표 작성. 반환: 표 다음 빈 행 번호."""
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
            if value_align and j - 1 in value_align:
                cell.alignment = value_align[j - 1]
            elif j == 1:
                cell.alignment = LEFT
            else:
                cell.alignment = CENTER
            if i % 2 == 1:
                cell.fill = fill(STRIPE)
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


def title_block(ws, title, subtitle, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    t = ws.cell(row=1, column=1, value=title)
    t.fill = fill(NAVY)
    t.font = TITLE_FONT
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 34

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    s = ws.cell(row=2, column=1, value=subtitle)
    s.fill = fill(NAVY)
    s.font = SUB_FONT
    s.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 18


def notes_block(ws, start_row, notes, ncols):
    r = start_row + 1
    for note in notes:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
        c = ws.cell(row=r, column=1, value=note)
        c.font = NOTE_FONT
        c.alignment = LEFT
        r += 1
    return r


wb = Workbook()

# ============================================================ 0) 표지/요약 시트
ws = wb.active
ws.title = "요약(Cover)"
ws.sheet_view.showGridLines = False
NC = 4
title_block(ws, "거시경제 상황 요약 레포트", "데이터 기준: 2026년 7월 초 · 작성일 2026-07-03", NC)

r = 4
r = section_bar(ws, r, "■ 리포트 개요", NC)
overview = [
    ["구분", "내용", "", ""],
    ["작성 목적", "현재 국내외 경제 상황을 미시·거시·국제 관점으로 요약", "", ""],
    ["구성", "3개 시트 — ①미시경제 ②거시경제 ③국제경제", "", ""],
    ["데이터 기준", "2026년 6월 발표치 및 주요 기관 전망치(2026-07-03 취합)", "", ""],
    ["주요 출처", "美 연준·BLS, 한국은행·KDI, IMF·EIA·세계은행, Eurostat, 中 NBS", "", ""],
]
for i, row in enumerate(overview):
    for j, val in enumerate(row, start=1):
        cell = ws.cell(row=r + i, column=j, value=val)
        cell.font = HDR_FONT if i == 0 else CELL_FONT
        cell.alignment = LEFT
        cell.border = BORDER
        if i == 0:
            cell.fill = fill(LTBLUE)
    ws.merge_cells(start_row=r + i, start_column=2, end_row=r + i, end_column=4)
r += len(overview) + 1

r = section_bar(ws, r, "■ 한눈에 보는 핵심 지표 (2026년 6월 기준)", NC)
kpi = [
    ["지표", "한국", "미국", "비고"],
    ["정책(기준)금리", "2.50%", "3.50~3.75%", "美 6/17 동결·매파적"],
    ["소비자물가(YoY)", "2.6% (4월)", "4.2% (5월, CPI)", "유가발 상방 압력"],
    ["성장률 전망('26)", "2.6%", "2.2%", "한은/연준 전망"],
    ["실업률", "약 2.8%", "4.2% (6월)", "美 고용 둔화"],
    ["환율·달러", "USD/KRW ≈1,554", "DXY ≈100.7", "원화 약세·달러 강세"],
    ["증시(YTD)", "코스피 변동성 확대", "S&P500 +9.6%", "AI랠리·전쟁 변수"],
]
for i, row in enumerate(kpi):
    for j, val in enumerate(row, start=1):
        cell = ws.cell(row=r + i, column=j, value=val)
        cell.border = BORDER
        cell.alignment = CENTER if i == 0 or j > 1 else LEFT
        if i == 0:
            cell.fill = fill(LTBLUE); cell.font = HDR_FONT
        else:
            cell.font = CELL_FONT
            if i % 2 == 0:
                cell.fill = fill(STRIPE)
r += len(kpi) + 1

r = section_bar(ws, r, "■ 현 국면 3줄 요약", NC)
summary = [
    "1) 中東(중동) 분쟁·호르무즈 해협 봉쇄 장기화로 국제유가가 급등하며 글로벌 인플레이션이 재점화됨.",
    "2) 美 연준은 금리를 동결했으나 물가 상방으로 매파적 선회(연내 인상 가능성), 달러 강세·원화 약세가 심화.",
    "3) 한국은 반도체 수출 호조로 성장 전망은 상향됐으나, 고환율·가계부채·부동산 PF가 하방 리스크로 상존.",
]
for i, s in enumerate(summary):
    ws.merge_cells(start_row=r + i, start_column=1, end_row=r + i, end_column=NC)
    cell = ws.cell(row=r + i, column=1, value=s)
    cell.font = CELL_FONT
    cell.alignment = LEFT
    cell.border = BORDER
    ws.row_dimensions[r + i].height = 20
r += len(summary) + 1

ws.column_dimensions["A"].width = 20
ws.column_dimensions["B"].width = 24
ws.column_dimensions["C"].width = 24
ws.column_dimensions["D"].width = 26
notes_block(ws, r, [
    "※ 본 자료는 공개된 통계·전망을 요약한 참고용이며, 투자판단의 근거로 사용될 수 없습니다.",
    "※ 일부 수치는 잠정치/기관 전망치로, 발표기관·시점에 따라 달라질 수 있습니다.",
], NC)

# ============================================================ 1) 미시경제 시트
ws1 = wb.create_sheet("미시경제")
ws1.sheet_view.showGridLines = False
NC = 5
title_block(ws1, "① 미시경제 (Microeconomics)", "가계·노동·산업·부문별 세부 지표 · 2026년 6월 기준", NC)

r = 4
r = section_bar(ws1, r, "■ 가계·소비 (Household & Consumption)", NC)
r = write_table(
    ws1, r,
    ["지표", "국가", "값", "전기/전년 대비", "비고"],
    [
        ["시간당 평균임금(AHE)", "미국", "$37.64", "+0.3% MoM / +3.5% YoY", "6월 민간부문"],
        ["실질임금", "미국", "둔화", "명목 3.5% < CPI 4.2%", "실질 구매력 감소"],
        ["가계부채", "한국", "높은 수준 지속", "GDP 대비 高", "추가 금리인상 제약 요인"],
        ["소비심리", "한국", "회복세 둔화", "-", "고물가·고환율 부담"],
        ["소비·내수", "한국", "완만한 개선", "수출 대비 약함", "한은 평가"],
    ],
    col_widths=[22, 10, 16, 22, 26],
)

r = section_bar(ws1, r, "■ 노동시장 세부 (Labor Market Detail, 美 2026년 6월)", NC)
r = write_table(
    ws1, r,
    ["항목", "값", "전월", "방향", "비고"],
    [
        ["비농업 신규고용(NFP)", "+57천명", "+129천명(하향)", "▼ 둔화", "예상치 +110천명 하회"],
        ["실업률", "4.2%", "-", "▼ 12개월 최저", "-11bp"],
        ["경제활동참가율", "61.5%", "61.8%", "▼", "2021.3 이후 최저"],
        ["전문·비즈니스 서비스", "+36천명", "-", "▲", "고용 증가 주도"],
        ["헬스케어", "+22천명", "-", "▲", "-"],
        ["사회복지(social assist.)", "+25천명", "-", "▲", "-"],
        ["레저·접객(hospitality)", "-61천명", "-", "▼", "월드컵·계절요인"],
    ],
    col_widths=[24, 12, 14, 14, 24],
)

r = section_bar(ws1, r, "■ 산업·부문 (Sectors, 한국)", NC)
r = write_table(
    ws1, r,
    ["부문", "상태", "코멘트", "리스크", ""],
    [
        ["반도체/IT", "호조", "수출·경상흑자 견인, 6월 수출 사상 첫 1,000억달러 돌파(잠정)", "글로벌 수요·재고", ""],
        ["부동산 PF", "부실 우려", "프로젝트파이낸싱 건전성 리스크", "금융 연쇄 부담", ""],
        ["건설·내수", "부진", "내수 회복 지연", "고금리·심리 위축", ""],
        ["수출 제조업", "개선", "반도체 등 기술제품 중심 흑자 확대", "환율·통상 환경", ""],
    ],
    col_widths=[16, 12, 40, 18, 4],
)

r = section_bar(ws1, r, "■ 물가 세부 (Prices Breakdown)", NC)
r = write_table(
    ws1, r,
    ["구성", "미국(5월)", "한국(4월)", "특징", ""],
    [
        ["헤드라인 CPI", "4.2% YoY", "2.6% YoY", "에너지·식품 포함, 상승세", ""],
        ["근원(Core) CPI", "2.9% YoY", "목표 부근", "에너지·식품 제외", ""],
        ["주요 상승요인", "유가·에너지", "유가·환율", "호르무즈 봉쇄발 공급 충격", ""],
    ],
    col_widths=[16, 14, 14, 34, 4],
)

notes_block(ws1, r, [
    "※ 미시 지표는 가계·기업·부문 등 개별 경제주체 수준의 흐름을 보여줍니다.",
    "※ 美 고용 = BLS 2026년 6월 고용상황 보고서. 韓 지표 = 한국은행·언론 취합(일부 잠정).",
], NC)

# ============================================================ 2) 거시경제 시트
ws2 = wb.create_sheet("거시경제")
ws2.sheet_view.showGridLines = False
NC = 6
title_block(ws2, "② 거시경제 (Macroeconomics)", "성장·물가·금리·고용 등 총량 지표 · 2026년 6월 기준", NC)

r = 4
r = section_bar(ws2, r, "■ 한국 (Korea)", NC)
r = write_table(
    ws2, r,
    ["지표", "최근치", "시점", "전망('26)", "전망('27)", "출처/비고"],
    [
        ["기준금리", "2.50%", "2026.05 동결", "-", "-", "한국은행"],
        ["소비자물가(YoY)", "2.6%", "2026.04", "2.7%", "2.3%", "3월 2.2%→상승"],
        ["실질GDP 성장률", "-", "-", "2.6%", "2.1%", "한은, 2.0%→2.6% 상향"],
        ["실업률", "약 2.8%", "2026 상반기", "-", "-", "완만"],
        ["경상수지", "GDP 대비 5.9%", "'25.6 4Q누적", "흑자 지속", "-", "반도체 주도"],
        ["성장 동력", "소비+수출 개선", "-", "-", "-", "수출 기여 큼"],
    ],
    col_widths=[18, 16, 14, 12, 12, 22],
)

r = section_bar(ws2, r, "■ 미국 (United States)", NC)
r = write_table(
    ws2, r,
    ["지표", "최근치", "시점", "전망('26)", "기타", "출처/비고"],
    [
        ["연방기금금리", "3.50~3.75%", "2026.06.17 동결", "연말 3.6~4.1%", "매파적", "FOMC, 점도표 상향"],
        ["소비자물가(CPI)", "4.2%", "2026.05", "3.6%(헤드라인)", "다년 최고", "근원 2.9%"],
        ["근원 물가", "2.9%", "2026.05", "3.3%(core)", "-", "연준 전망 상향"],
        ["실질GDP 성장률", "-", "-", "2.2%", "-0.2%p(3월대비)", "연준 전망 하향"],
        ["실업률", "4.2%", "2026.06", "4.3%", "-0.1%p", "고용 둔화"],
        ["통화정책 기조", "동결·매파적", "-", "인상 가능성 시사", "-", "물가 재점화 대응"],
    ],
    col_widths=[16, 14, 16, 16, 14, 20],
)

r = section_bar(ws2, r, "■ 거시 국면 진단 (Macro Assessment)", NC)
diag = [
    "· 물가: 유가 급등이 헤드라인 물가를 끌어올리며, 美 연준의 물가 전망이 3월 대비 큰 폭 상향(2.7%→3.6%).",
    "· 성장: 美 성장 전망은 소폭 하향(2.2%), 韓은 반도체 수출 힘입어 상향(2.6%)으로 대조.",
    "· 통화정책: 韓·美 모두 동결 상태이나, 물가 상방으로 완화 사이클 지연 또는 매파 전환 압력.",
    "· 리스크: 스태그플레이션 우려(물가↑+성장↓), 고환율, 가계부채·부동산 PF.",
]
for i, s in enumerate(diag):
    ws2.merge_cells(start_row=r + i, start_column=1, end_row=r + i, end_column=NC)
    cell = ws2.cell(row=r + i, column=1, value=s)
    cell.font = CELL_FONT
    cell.alignment = LEFT
    cell.border = BORDER
    ws2.row_dimensions[r + i].height = 20
r += len(diag) + 1

notes_block(ws2, r, [
    "※ 거시 지표는 한 경제 전체의 총량 흐름(성장·물가·고용·금리)을 나타냅니다.",
    "※ 전망치는 각 중앙은행(한국은행·美 연준)의 공식 전망(2026년 상반기 기준).",
], NC)

# ============================================================ 3) 국제경제 시트
ws3 = wb.create_sheet("국제경제")
ws3.sheet_view.showGridLines = False
NC = 5
title_block(ws3, "③ 국제경제 (International / Global)", "세계 성장·원자재·외환·주요국 · 2026년 6월 기준", NC)

r = 4
r = section_bar(ws3, r, "■ 세계 성장 (Global Growth)", NC)
r = write_table(
    ws3, r,
    ["지역/기관", "지표", "값", "이전/비교", "비고"],
    [
        ["신흥국(IMF)", "성장률('26)", "3.9%", "1월 4.2%→하향", "중동 분쟁 영향"],
        ["유로존", "GDP(1Q '26)", "+0.1% QoQ", "'25 2Q 이후 최저", "에너지 공급 압박"],
        ["유로존", "성장률('26/'27)", "1.1% / 1.5%", "-", "완만한 회복"],
        ["중국", "GDP(1Q '26)", "+5.0% YoY", "목표 4.5~5%", "회복력 시현"],
        ["미국", "성장률('26)", "2.2%", "연준 전망", "소폭 하향"],
    ],
    col_widths=[16, 16, 14, 18, 22],
)

r = section_bar(ws3, r, "■ 원자재·에너지 (Commodities & Energy)", NC)
r = write_table(
    ws3, r,
    ["품목", "최근가", "시점", "전망('26 평균)", "비고"],
    [
        ["Brent 유", "≈$95.06/bbl", "2026.06.02", "$82~85 (IMF/GS)", "호르무즈 봉쇄로 급등·변동성"],
        ["WTI 유", "≈$92.32/bbl", "2026.06.02", "-", "-"],
        ["유가 시나리오", "Q2 ≈$106", "EIA STEO", "Q4 ≈$89→연말 $70", "해협 정상화 가정"],
        ["상방 리스크", ">$120 (Brent)", "GS", "-", "봉쇄 3Q까지 지속 시"],
    ],
    col_widths=[16, 16, 14, 20, 26],
)

r = section_bar(ws3, r, "■ 외환·금융시장 (FX & Markets)", NC)
r = write_table(
    ws3, r,
    ["지표", "값", "시점", "방향", "비고"],
    [
        ["달러인덱스(DXY)", "≈100.7", "2026.06", "▲ 강세", "'25.5 이후 최고, 연준 매파"],
        ["USD/KRW", "≈1,554 (월말)", "2026.06", "▲ 원화 약세", "6월 평균 ≈1,538"],
        ["S&P 500", "≈7,354 (+9.6% YTD)", "2026.06", "▲", "AI랠리, 6월 -1%(월간)"],
        ["나스닥", "3월말 대비 +21%", "2026 2Q", "▲", "전쟁발 급락 후 반등"],
        ["코스피", "변동성 확대", "2026.06", "환율 연동", "고환율 시 외국인 부담"],
    ],
    col_widths=[18, 20, 12, 16, 24],
)

r = section_bar(ws3, r, "■ 국제 리스크 (Global Risks)", NC)
risk = [
    "· 지정학: 중동 분쟁 및 호르무즈 해협 봉쇄 3개월+ 지속 → 원유 수송 차질, 글로벌 물가·성장에 직접 충격.",
    "· 에너지: 유가·LNG 공급 차질이 유로존 등 에너지 수입 경제의 성장을 제약.",
    "· 정책: 美 달러 강세·연준 매파 전환이 신흥국 자본유출·통화 약세 압력으로 작용.",
    "· 무역: 반도체 등 기술 수출 사이클은 한국 등 수출국에 긍정적이나 대외 여건 변화에 민감.",
]
for i, s in enumerate(risk):
    ws3.merge_cells(start_row=r + i, start_column=1, end_row=r + i, end_column=NC)
    cell = ws3.cell(row=r + i, column=1, value=s)
    cell.font = CELL_FONT
    cell.alignment = LEFT
    cell.border = BORDER
    ws3.row_dimensions[r + i].height = 20
r += len(risk) + 1

notes_block(ws3, r, [
    "※ 국제 지표는 세계 성장·원자재·외환 등 대외 환경을 보여줍니다.",
    "※ 유가·환율은 변동성이 큰 시점의 스냅샷으로, 실시간 시세와 차이가 있을 수 있습니다.",
], NC)

# ============================================================ 4) 출처 시트
ws4 = wb.create_sheet("출처(Sources)")
ws4.sheet_view.showGridLines = False
NC = 3
title_block(ws4, "출처 및 유의사항 (Sources & Notes)", "2026-07-03 웹 취합", NC)
r = 4
r = write_table(
    ws4, r,
    ["구분", "기관/자료", "내용"],
    [
        ["미국", "美 연준(FOMC)", "2026.06.17 FOMC 성명·경제전망(점도표)"],
        ["미국", "美 노동통계국(BLS)", "2026년 6월 고용상황 보고서"],
        ["한국", "한국은행(BOK)", "기준금리·물가·성장 전망, 경제통계"],
        ["한국", "KDI", "2026 상반기 경제전망"],
        ["국제", "IMF", "2026년 4월 세계경제전망(WEO)"],
        ["국제", "EIA / Goldman Sachs / World Bank", "단기에너지전망(STEO)·원자재 전망"],
        ["국제", "Eurostat / 中 국가통계국(NBS)", "유로존·중국 GDP"],
        ["시장", "CNBC/CNN/Trading Economics 등", "증시·환율·유가 시세(취합)"],
    ],
    col_widths=[10, 34, 46],
)
notes_block(ws4, r, [
    "※ 본 레포트는 공개 자료를 요약·재구성한 참고용 문서입니다.",
    "※ 수치는 발표 시점·기관·개정에 따라 달라질 수 있으며, 정확한 값은 원출처를 확인하십시오.",
    "※ 작성: 2026-07-03 / 데이터 기준: 2026년 6월(일부 잠정·전망 포함).",
], NC)

fname = "거시경제_상황_레포트_2026-07.xlsx"
wb.save(fname)
print("saved:", fname)
for s in wb.sheetnames:
    print(" -", s)
