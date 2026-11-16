from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter
import sys

OUT = sys.argv[1]
F = "Malgun Gothic"
NAVY, BLUE, PALE, GRAY = "0B1E3D", "005AAB", "DCE6F2", "F1F4F8"
INPUT = PatternFill("solid", fgColor="FFF6D5")  # 입력 칸
EXF = PatternFill("solid", fgColor="EDEDED")    # 예시 행
thin = Side(style="thin", color="C9D3E0")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
STATUS = ["미착수", "진행중", "완료", "해당없음"]

wb = Workbook()

def font(**k):
    k.setdefault("name", F); k.setdefault("size", 10)
    return Font(**k)

def title(ws, text, sub):
    ws["A1"] = text; ws["A1"].font = font(size=16, bold=True, color=NAVY)
    ws["A2"] = sub; ws["A2"].font = font(size=10, color="5B6676")
    ws.row_dimensions[1].height = 28

def table(ws, start_row, headers, widths, rows, example=None, input_cols=None, status_col=None, blank_rows=0, dv_lists=None, wrap_cols=()):
    """headers row at start_row; example row (gray) right below; then data rows; then blank rows.
    input_cols: 1-based column indexes the owner fills (yellow). returns (first_data_row, last_row)."""
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(start_row, i, h)
        c.font = font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=BLUE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BOX
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[start_row].height = 32
    r = start_row + 1
    if example:
        for i, v in enumerate(example, 1):
            c = ws.cell(r, i, v); c.fill = EXF; c.font = font(italic=True, color="666666"); c.border = BOX
            c.alignment = Alignment(vertical="center", wrap_text=True)
        r += 1
    first = r
    allrows = list(rows) + [[None] * len(headers) for _ in range(blank_rows)]
    for row in allrows:
        for i in range(1, len(headers) + 1):
            v = row[i - 1] if i - 1 < len(row) else None
            c = ws.cell(r, i, v); c.border = BOX; c.font = font()
            c.alignment = Alignment(vertical="center", wrap_text=(i in wrap_cols))
            if input_cols and i in input_cols:
                c.fill = INPUT
        if status_col and ws.cell(r, status_col).value is None:
            ws.cell(r, status_col, "미착수")
        r += 1
    last = r - 1
    if status_col:
        col = get_column_letter(status_col)
        dv = DataValidation(type="list", formula1='"' + ",".join(STATUS) + '"', allow_blank=True)
        ws.add_data_validation(dv); dv.add(f"{col}{first}:{col}{last}")
        rng = f"{col}{first}:{col}{last}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"완료"'], fill=PatternFill("solid", fgColor="D9F2E3")))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"진행중"'], fill=PatternFill("solid", fgColor="FFE9C7")))
    for colidx, items in (dv_lists or {}).items():
        col = get_column_letter(colidx)
        dv = DataValidation(type="list", formula1='"' + ",".join(items) + '"', allow_blank=True)
        ws.add_data_validation(dv); dv.add(f"{col}{first}:{col}{last}")
    ws.freeze_panes = ws.cell(start_row + 1, 1)
    return first, last

REG = {}  # sheet -> (status col letter, first, last)

# ---------------- 1. 안내 ----------------
ws = wb.active; ws.title = "안내"
title(ws, "(주)토인 회사소개서 리뉴얼 — 자료 수집 양식", "회사소개서 v1(국문·영문 23p)의 '입력 필요' 항목을 채우기 위한 양식입니다. 제출 마감: 2026-12-12 · 최종 배포: 2027-01-05 (확정)")
ws.column_dimensions["A"].width = 22; ws.column_dimensions["B"].width = 70; ws.column_dimensions["C"].width = 16; ws.column_dimensions["D"].width = 14
rows = [
    ("작성 방법", None),
    ("노란색 칸", "담당자가 입력하는 칸입니다. 그 외 칸(흰색)은 미리 채워둔 기준 정보이며, 틀렸으면 비고에 적어 주세요."),
    ("회색 행 (예시)", "입력 형식을 보여주는 예시입니다. 덮어쓰지 말고 아래 행부터 작성해 주세요. 진행률 계산에서 제외됩니다."),
    ("상태", "각 행 마지막 '상태' 칸에서 미착수 / 진행중 / 완료 / 해당없음 중 선택. '진행현황' 시트에 자동 집계됩니다."),
    ("단위", "길이 mm, 면적 ㎡, 출력 kW, 톤수 ton, 회전수 min⁻¹(rpm). 숫자만 입력하고 단위는 열 제목을 따릅니다."),
    ("근거", "수치·문구에는 근거(카탈로그 페이지, 인증서, 결산서 등)를 '근거/출처' 칸에 적어 주세요. 근거 없는 홍보 문구는 싣지 않습니다."),
    ("사진 파일 이름", "페이지번호_사업부_내용_순번.jpg  예) p11_판금_MURATEC_펀치_01.jpg  · 장변 3,000px 이상, 원본 그대로 공유폴더에 업로드"),
    (None, None),
    ("시트별 담당", None),
]
r = 4
for a, b in rows:
    ws.cell(r, 1, a).font = font(bold=True, size=11 if b is None else 10, color=NAVY if b is None else "1A2230")
    if b: ws.cell(r, 2, b).font = font(); ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
    if b: ws.row_dimensions[r].height = 30; ws.cell(r, 1).alignment = Alignment(vertical="top")
    if a == "노란색 칸": ws.cell(r, 1).fill = INPUT
    if a == "회색 행 (예시)": ws.cell(r, 1).fill = EXF
    r += 1
sheets_owner = [
    ("공통_회사정보", "경영관리부", "p2–6, 21–23"),
    ("판금_장비사양", "판금기계사업부", "p11"),
    ("판금_브랜드확인", "판금기계사업부", "p9–10, 12, 22"),
    ("공작_라인업", "공작기계사업부", "p14"),
    ("공작_고객사례", "공작기계사업부", "p15"),
    ("자동차_품목", "자동차부품사업부", "p18"),
    ("자동차_공장설비", "자동차부품사업부", "p17, 19–20"),
    ("사진_체크리스트", "각 사업부", "전체"),
]
first, last = table(ws, r, ["시트", "담당 부서 / 담당자", "관련 페이지", "마감"], [22, 70, 16, 14],
                    [(s, o + " / 오상원 부장", p, "2026-11-21" if s == "공통_회사정보" else "2026-12-12") for s, o, p in sheets_owner], input_cols={2})
ws.cell(last + 2, 1, "전체 일정").font = font(bold=True, size=11, color=NAVY)
for k, (a, b) in enumerate([("11/17–11/21", "사실관계 확정 (공통_회사정보)"), ("11/17–12/12", "자료 수집 · 사진 촬영 (각 사업부)"), ("12/14–12/19", "v2 반영 · 디자인"), ("12/21–12/26", "검수 (각 사업부: 사양 · 표기 · 오탈자)"), ("2027-01-05", "최종 배포 (PPT · PDF · 홈페이지) — 확정")]):
    ws.cell(last + 3 + k, 1, a).font = font(bold=(k == 4)); ws.cell(last + 3 + k, 2, b).font = font(bold=(k == 4))
for rr in range(first, last + 1):
    ws.cell(rr, 1).hyperlink = f"#'{ws.cell(rr,1).value}'!A1"; ws.cell(rr, 1).font = font(color=BLUE, underline="single")
ws.freeze_panes = None

# ---------------- 2. 진행현황 (filled after) ----------------
dash = wb.create_sheet("진행현황")

# ---------------- 3. 공통_회사정보 ----------------
ws = wb.create_sheet("공통_회사정보")
title(ws, "공통 · 회사 정보 확인", "담당: 경영관리부 · 자료마다 수치가 달랐던 항목을 확정합니다 (현재값 = 회사소개서 v1에 들어간 값)")
H = ["No", "구분", "항목", "현재값 (v1)", "확정값", "근거 / 출처", "관련 페이지", "비고", "상태"]
items = [
    ("기본", "설립 연월", "1992.02 (확정)", "p3, 4"),
    ("기본", "갑류 무역대리업 등록 연도", "1993.07", "p4"),
    ("재무", "2025년 매출액 (억원, 결산 확정)", "430 (잠정)", "p3, 5"),
    ("재무", "2024년 매출액 확인 (그래프 438 / 표지 430 혼재)", "438", "p5"),
    ("인원", "임직원 수 (2026년 기준)", "89명 (2025)", "p3"),
    ("인원", "본사 인원", "23명", "p6"),
    ("인원", "대구지사 인원 (공작·자동차·판금 기술지원2팀 포함)", "66명", "p6"),
    ("인원", "자동차부품사업부 생산 인력", "40명", "p16"),
    ("조직", "최신 조직도 반영 (부서명: 경영관리부 / 경영지원부 통일)", "경영관리부", "p6"),
    ("연혁", "2007 ISO 9001 인증 (증빙)", "2007.03 표기", "p4"),
    ("연혁", "2013 기업부설연구소 설립 (증빙)", "2013 표기", "p4"),
    ("연혁", "2013 IBK 우수거래기업 (증빙)", "2013 표기", "p4"),
    ("인증", "IATF 16949 취득 연월 · 유효기간", "", "p4, 20"),
    ("인증", "ISO 14001 취득 연월 · 유효기간", "", "p4, 20"),
    ("인증", "T-SQ / P-SQ / SSQ 취득 연월 · 유효기간", "", "p4, 20"),
    ("인증", "ISO 9001 현재 보유 여부", "", "p20"),
    ("대표", "대표 인사말 확정 (400~600자)", "초안 작성됨", "p2"),
    ("대표", "대표이사 사진 (배경 제거본)", "", "p2"),
    ("연락처", "대표 이메일 / A/S 이메일", "", "p21, 23"),
    ("연락처", "본사 · 대구 팩스번호", "", "p23"),
    ("연락처", "QR 연결 주소 (홈페이지 또는 카탈로그 URL)", "www.toinint.com", "p23"),
    ("브랜드", "로고 원본 (ai/eps) · 영문 로고 유무", "", "p1, 23"),
    ("고객", "자동차부품 고객 표현: '현대·기아 1차 협력' vs '경창산업 등 협력사 납품'", "경창산업 등 협력사 납품", "p7, 18"),
]
first, last = table(ws, 4, H, [5, 9, 46, 22, 22, 26, 11, 20, 10],
                    [(i + 1, a, b, c, None, None, d) for i, (a, b, c, d) in enumerate(items)],
                    example=("예시", "재무", "2025년 매출액 (억원)", "430 (잠정)", "432", "2025 결산 재무제표", "p3, 5", "", "예시"),
                    input_cols={5, 6, 8}, status_col=9, wrap_cols=(3, 6, 8))
REG["공통_회사정보"] = ("I", first, last)

# ---------------- 4. 판금_장비사양 ----------------
ws = wb.create_sheet("판금_장비사양")
title(ws, "판금기계사업부 · 대표 장비 사양", "담당: 판금기계사업부 · 회사소개서 p11 표 · 값은 회사 보유 메이커 카탈로그(구글 드라이브 '카탈로그' 폴더)에서 추출 — 확인·수정해 주세요")
H = ["No", "브랜드", "장비군", "대표 모델", "가공 범위\n(X×Y, mm)", "레이저 출력\n(kW)", "펀치·절곡 톤수\n(ton)", "최대 가공 두께\n(mm, 재질)", "핵심 특징\n(30자 이내)", "사진 파일명", "카탈로그 / 근거", "상태"]
pre = [
    (1, "MURATEC", "터렛 펀치 프레스", "MOTORUM M3048TG", "1,250 × 2,500", "-", 30, "6.35 (기준 재질 확인)", "서보 램 · 펀칭 전력 85% 절감", None, "M3048TG_M3058TG.pdf 사양 페이지", "진행중"),
    (2, "MURATEC", "파이버 레이저", "LS3015GC", "3,070 × 1,550", "2.5 / 4.0 / 6.0", "-", "6 kW: 연강 25 / SUS·AL 약 14 / 황동·구리 약 10", "고속 모드 가공시간 15% 단축", None, "LS3015GC 카타로그(한).pdf p07 · 두께는 그래프 판독치(2.5/4.0 kW: 연강 약 19/25, SUS 약 10/12)", "진행중"),
    (3, "MURATEC", "듀얼 드라이브 절곡기", "BH13530", "절곡 길이 3,100", "-", 135, "-", "볼스크류 + 유압 듀얼 드라이브", None, "종합카다로그(MURATA).pdf p21-22 · 판매중 확인(2026.11)", "진행중"),
    (4, "HSG", "평판 파이버 레이저", "HS-G3015A (XLASE)", "3,000 × 1,500", 4, "-", "연강 22 / SUS 12 / AL 12", "팔레트 교환 8초 · 최대 140 m/min", None, "G3015A-4KW Machine 카탈로그.pdf", "진행중"),
    (5, "HSG", "파이프 레이저", None, None, None, "-", None, None, None, "카탈로그 PDF 용량 초과(500MB+) — 사양 페이지만 따로 저장 필요", "미착수"),
    (6, "KYOWA", "AC 서보 패널벤더", "KMP3023LFT", "3,050 × 1,525", "-", 60, "SPCC 2.3 / SUS 1.5", "전축 AC 서보 · 유압 미사용", None, "KMP-카달로그(영문).pdf 사양 페이지", "진행중"),
    (7, "ADIRA", "절곡기", "PA·PH·PF (모델 확인)", None, "-", None, None, None, None, "회사 카탈로그 폴더에 없음 — 확보 필요", "미착수"),
    (8, "ADIRA", "절단기", "GH·GV (모델 확인)", None, "-", None, None, None, None, "회사 카탈로그 폴더에 없음 — 확보 필요", "미착수"),
    (9, "AIZAWA", "프레스 브레이크", None, None, "-", None, None, None, None, "회사 카탈로그 폴더에 없음 — 확보 필요", "미착수"),
]
first, last = table(ws, 4, H, [5, 12, 20, 20, 16, 14, 14, 22, 30, 26, 34, 10], pre,
                    example=("예시", "HSG", "평판 파이버 레이저", "예) ○○-3015", "3,050 × 1,530", 12, "-", "연강 25 / SUS 20", "고속 절단 · 자동 노즐 교환", "p11_판금_HSG_레이저_01.jpg", "HSG 카탈로그 2025 p.12", "예시"),
                    input_cols={4, 5, 6, 7, 8, 9, 10, 11}, status_col=12, blank_rows=3,
                    dv_lists={2: ["MURATEC", "HSG", "ADIRA", "KYOWA", "AIZAWA", "KYOKKO", "CONIC", "기타"]}, wrap_cols=(8, 9, 11))
for rr in range(first, last + 1):
    ws.cell(rr, 7).number_format = "#,##0"
ws.cell(last + 2, 1, "※ 해당 없는 칸은 '-'. 모델명은 실제 판매 중인 모델만. 제외: M2048HL(현재 미판매), Muratec-TOYO HPB·HYB·APB 절곡기·HSS 절단기(단종).").font = font(color="B86E00", bold=True)
REG["판금_장비사양"] = ("L", first, last)

# ---------------- 5. 판금_브랜드확인 ----------------
ws = wb.create_sheet("판금_브랜드확인")
title(ws, "판금기계사업부 · 브랜드 / 문구 확인", "담당: 판금기계사업부 · 회사소개서의 * 표시 문구와 취급 브랜드 목록을 확정합니다")
H = ["No", "확인 항목", "현재 표기 (v1)", "확정 내용", "근거 / 출처", "관련 페이지", "상태"]
items = [
    ("ADIRA 홍보 문구", "유럽 60년 역사*", "p10"), ("KYOWA 홍보 문구", "전축 서보 구동*", "p10"), ("AIZAWA 홍보 문구", "100년 역사 · 40,000대 출하*", "p10"),
    ("KYOKKO · CONIC 현재 거래 여부 (부품)", "p22에 표기", "p22"), ("3/25 초안의 'DGDY' 브랜드 정체 / 삭제 여부", "삭제됨", "p10"),
    ("공정–브랜드 매핑 검수 (MURATEC 절곡기 · 자동화 범위 등)", "p9 표기대로", "p9"), ("전국 가동 대수 (2026 기준)", "1,500+", "p3, 8"),
    ("해외 판매국 (현재 기준)", "중국 · 베트남 · 인도 · 미국", "p8, 12"), ("본사 전시장 현재 전시 장비", "평판 레이저 · MURATEC 절곡기 · 협동로봇", "p12"),
    ("해외 납품 사례 1~2건 (국가 · 장비 · 연도)", "", "p12"), ("브랜드 로고 원본 5종 수집", "", "p10, 22"),
]
first, last = table(ws, 4, H, [5, 46, 34, 34, 26, 11, 10], [(i + 1, a, b, None, None, c) for i, (a, b, c) in enumerate(items)],
                    example=("예시", "ADIRA 홍보 문구", "유럽 60년 역사*", "1956년 설립 포르투갈 절곡기 제조사", "ADIRA 공식 홈페이지 회사소개", "p10", "예시"),
                    input_cols={4, 5}, status_col=7, wrap_cols=(2, 3, 4, 5))
REG["판금_브랜드확인"] = ("G", first, last)

# ---------------- 6. 공작_라인업 ----------------
ws = wb.create_sheet("공작_라인업")
title(ws, "공작기계사업부 · MURATEC CNC 선반 라인업", "담당: 공작기계사업부 · 회사소개서 p14 표 · 값은 MURATEC 홈페이지 기준 잠정치(2026.11) — 국내 판매 사양으로 확인·수정해 주세요")
H = ["No", "장비 유형", "시리즈", "대표 모델\n(홈페이지)", "국내 판매 모델\n(확정)", "최대 척\n(inch)", "최대 주축 회전수\n(min⁻¹)", "최대 가공경\n(mm)", "주요 특징\n(30자 이내)", "국내 판매 여부", "사진 파일명", "카탈로그 / 근거", "상태"]
WEB = "MURATEC 홈페이지 (잠정, 2026.11)"
pre = [
    (1, "트윈 스핀들 CNC 처커", "MW", "MW120", None, "8 (ø210 mm)", "6,000", None, "2주축 동시 가공 · 갠트리 로더 일체", None, None, WEB, "진행중"),
    (2, "싱글 스핀들 CNC 처커", "MS", "MS20 · MS50 · MS60 · MS90 · MS100", None, "5–10", None, None, "부품 크기별 모델 선택", None, None, WEB, "진행중"),
    (3, "Y축 트윈 스핀들 복합기", "MWR", "MWR120", None, None, None, None, "Y축 ±60 mm · 회전공구 5.5 kW", None, None, WEB, "진행중"),
    (4, "회전공구 트윈 스핀들 터닝센터", "MD", "MD120 · MD200", None, "8–10", "3,500 (MD200)", None, "전 터렛 회전공구 · 갠트리 로더", None, None, WEB, "진행중"),
    (5, "CNC 샤프트 선반 (심압대)", "MJ", "MJ60 · MJ100 · MJ200", None, "8–10", "3,500–6,000", None, "샤프트류 자동화 · 유압 심압대", None, None, WEB, "진행중"),
    (6, "대향 트윈 스핀들 복합기", "MT", "MT12 · MT25 · MT100", None, "6–10", "3,500–5,000", None, "대향 2주축 · 2~3 터렛 복합 가공", None, None, WEB, "진행중"),
    (7, "(취급 여부 확인)", "ML", None, None, None, None, None, "홈페이지에서 확인 안 됨", None, None, None, "미착수"),
]
first, last = table(ws, 4, H, [5, 26, 8, 26, 18, 12, 16, 12, 30, 13, 26, 26, 10], pre,
                    example=("예시", "트윈 스핀들 CNC 처커", "MW", "MW120", "MW120G", "8", "6,000", 120, "갠트리 로더 일체형", "판매중", "p14_공작_MW120_01.jpg", "MURATEC 카탈로그 p.5", "예시"),
                    input_cols={5, 6, 7, 8, 9, 10, 11, 12}, status_col=13, blank_rows=2,
                    dv_lists={3: ["MW", "MWR", "MS", "MD", "MJ", "MT", "ML", "기타"], 10: ["판매중", "미판매", "확인중"]}, wrap_cols=(4, 9, 12))
for rr in range(first, last + 1): ws.cell(rr, 8).number_format = "#,##0"
ws.cell(last + 2, 1, "※ 빈칸: MS 주축 회전수, MWR120 척·회전수, 최대 가공경(홈페이지에 없음). 홈페이지 안에서도 버전별 수치가 달라(MW120 척 6\"/8\", 3,000~6,000 min⁻¹) 국내 판매 사양 기준으로 확정해 주세요.").font = font(color="B86E00", bold=True)
REG["공작_라인업"] = ("M", first, last)

# ---------------- 7. 공작_고객사례 ----------------
ws = wb.create_sheet("공작_고객사례")
title(ws, "공작기계사업부 · 고객사 로고 / 공정 개선 사례", "담당: 공작기계사업부 · 회사소개서 p15 · 로고는 현재 거래처이고 게재에 문제없는 곳만 싣습니다")
H = ["No", "고객사 (2025.03본 로고)", "현재 거래 여부", "로고 게재 가능", "주요 납품 장비", "비고", "상태"]
cust = ["일진 (ILJIN)", "S&S INC", "OTO", "PHC Valeo", "삼익THK", "한국케이지(주)", "부산주공", "경창산업", "경창정공", "한국파워트레인 (KAPEC)", "평화정공 (PHA)", "디아이씨 (DIC)", "유창정밀공업", "S&S M-TECH"]
first, last = table(ws, 4, H, [5, 26, 14, 14, 28, 26, 10], [(i + 1, c) for i, c in enumerate(cust)],
                    example=("예시", "경창산업", "거래중", "가능", "MW120 6대", "", "예시"),
                    input_cols={3, 4, 5, 6}, status_col=7, blank_rows=3,
                    dv_lists={3: ["거래중", "거래종료", "확인중"], 4: ["가능", "불가", "확인중"]})
REG["공작_고객사례"] = ("G", first, last)
r0 = last + 3
ws.cell(r0, 1, "공정 개선 사례 (1~2건)").font = font(bold=True, size=12, color=NAVY)
H2 = ["No", "고객사", "적용 장비", "대상 부품", "개선 전 사이클타임\n(초)", "개선 후 사이클타임\n(초)", "단축률", "개선 전 불량률", "개선 후 불량률", "한 줄 요약", "상태"]
cf, cl = table(ws, r0 + 1, H2, [5, 26, 14, 14, 16, 16, 10, 13, 13, 30, 10], [(1,), (2,)],
               example=("예시", "○○정공", "MW200", "허브", 95, 62, None, 0.012, 0.004, "트윈 스핀들 전환으로 사이클타임 35% 단축", "예시"),
               input_cols={2, 3, 4, 5, 6, 8, 9, 10}, status_col=11, wrap_cols=(10,))
for rr in range(cf - 1, cl + 1):
    ws.cell(rr, 7, f'=IFERROR(IF(OR(E{rr}="",F{rr}=""),"",1-F{rr}/E{rr}),"")'); ws.cell(rr, 7).number_format = "0.0%"
    ws.cell(rr, 8).number_format = "0.0%"; ws.cell(rr, 9).number_format = "0.0%"
ws.cell(cf - 1, 7).fill = EXF
ws.cell(cl + 1, 1, "※ '단축률'은 자동 계산(1 − 개선 후 ÷ 개선 전). 불량률은 % 형식으로 입력(예: 1.2%).").font = font(color="5B6676")
ws.freeze_panes = "A5"
REG["공작_고객사례(사례)"] = ("K", cf, cl)

# ---------------- 8. 자동차_품목 ----------------
ws = wb.create_sheet("자동차_품목")
title(ws, "자동차부품사업부 · 생산 품목", "담당: 자동차부품사업부 · 회사소개서 p18 · 자료마다 품목명이 달라 현재 양산 품목으로 확정합니다")
H = ["No", "품목명 (자료상 표기)", "현재 양산 여부", "회사소개서 표기명 (확정)", "적용 변속기 / 차종", "납품처", "월 생산 능력\n(개)", "관리 공차\n(mm)", "주요 공정", "사진 상태", "상태"]
prods = ["HUB 37R CLUTCH", "HUB 46 CLUTCH", "감마 IVT 부품", "PLATE MID SUN CONN", "PLATE SUN MID", "KAPPA CVT 부품 (3/25본)", "U/D BREAK (3/25본)", "HUB & SHAFT (3/25본)", "SHALL (2025.03본)"]
first, last = table(ws, 4, H, [5, 26, 13, 24, 22, 18, 13, 12, 24, 16, 10], [(i + 1, p) for i, p in enumerate(prods)],
                    example=("예시", "HUB 37R CLUTCH", "양산", "HUB 37R CLUTCH", "8속 자동변속기", "경창산업", 30000, "±0.01", "CNC 선삭 → 연삭 → 검사", "재촬영 필요", "예시"),
                    input_cols={3, 4, 5, 6, 7, 8, 9, 10}, status_col=11, blank_rows=3,
                    dv_lists={3: ["양산", "단종", "개발중", "확인중"], 10: ["재촬영 필요", "원본 있음", "촬영 완료"]}, wrap_cols=(9,))
for rr in range(first, last + 1): ws.cell(rr, 7).number_format = "#,##0"
ws.cell(last + 2, 1, "※ v1 품목 사진은 2025.03본 화면 캡처에서 잘라낸 저해상도 — 전 품목 재촬영 권장 (흰/회색 배경, 동일 각도)").font = font(color="B86E00", bold=True)
REG["자동차_품목"] = ("K", first, last)

# ---------------- 9. 자동차_공장설비 ----------------
ws = wb.create_sheet("자동차_공장설비")
title(ws, "자동차부품사업부 · 공장 면적 / 보유 설비", "담당: 자동차부품사업부 · 회사소개서 p17, p19, p20")
ws.cell(4, 1, "① 동별 면적 (㎡)").font = font(bold=True, size=12, color=NAVY)
H = ["구역", "1층", "2층", "3층", "합계", "비고", "상태"]
af, al = table(ws, 5, H, [12, 12, 12, 12, 12, 30, 10], [("B동", 2238, 430.8, 252), ("C동", 2238, 430.8, 252), ("D동",), ("기타",)],
               input_cols={2, 3, 4, 6}, status_col=7)
for rr in range(af, al + 1):
    ws.cell(rr, 5, f"=IF(COUNT(B{rr}:D{rr})=0,\"\",SUM(B{rr}:D{rr}))")
    for cc in range(2, 6): ws.cell(rr, cc).number_format = "#,##0.0"
tr = al + 1
ws.cell(tr, 1, "건축 합계 (㎡)").font = font(bold=True)
for cc in range(2, 6):
    col = get_column_letter(cc); ws.cell(tr, cc, f"=SUM({col}{af}:{col}{al})"); ws.cell(tr, cc).number_format = "#,##0.0"; ws.cell(tr, cc).font = font(bold=True)
ws.cell(tr + 1, 1, "평 환산"); ws.cell(tr + 1, 5, f"=E{tr}/3.3058"); ws.cell(tr + 1, 5).number_format = '#,##0" 평"'
ws.cell(tr + 1, 6, "회사소개서 표기: 2,815평 / 대지 16,529㎡(5,000평) — 합계와 다르면 확인").font = font(color="5B6676")
REG["자동차_공장설비(면적)"] = ("G", af, al)

r0 = tr + 4
ws.cell(r0, 1, "② 보유 설비").font = font(bold=True, size=12, color=NAVY)
H = ["설비 종류", "대수", "주요 메이커 / 모델", "배치 (동)", "비고", "상태"]
ef, el = table(ws, r0 + 1, H, [24, 12, 30, 12, 30, 10], [("CNC 선반",), ("연삭기",), ("로봇",), ("갠트리 로더",), ("컨베이어 라인",), ("측정·검사 장비",), ("열처리 설비",)],
               example=("CNC 선반", 48, "MURATEC MW120 외", "B·C동", "", "예시"),
               input_cols={2, 3, 4, 5}, status_col=6, blank_rows=3)
for rr in range(ef, el + 1): ws.cell(rr, 2).number_format = "#,##0"
ws.cell(el + 1, 1, "합계 (대)").font = font(bold=True); ws.cell(el + 1, 2, f"=SUM(B{ef}:B{el})"); ws.cell(el + 1, 2).font = font(bold=True)
REG["자동차_공장설비(설비)"] = ("F", ef, el)

r0 = el + 4
ws.cell(r0, 1, "③ 문구 근거 확인").font = font(bold=True, size=12, color=NAVY)
H = ["확인 항목", "현재 표기 (v1)", "확정 내용", "근거", "비고", "상태"]
qf, ql = table(ws, r0 + 1, H, [24, 12, 30, 12, 30, 10],
               [("정밀도 관리 수준", "±0.01mm*"), ("무인 가동 시간", "24시간*"), ("검사 공정 흐름 (수입→공정→출하→추적)", "p20 표기"), ("공장 배치도 최신본", "2025.03본 이미지")],
               input_cols={3, 4}, status_col=6, wrap_cols=(1, 3))
ws.freeze_panes = None
REG["자동차_공장설비(문구)"] = ("F", qf, ql)

# ---------------- 10. 사진_체크리스트 ----------------
ws = wb.create_sheet("사진_체크리스트")
title(ws, "사진 · 이미지 수집 체크리스트", "장변 3,000px 이상 원본 · 파일명 규칙: 페이지번호_사업부_내용_순번.jpg")
H = ["No", "페이지", "필요 사진 / 이미지", "현재 상태", "담당", "파일명", "상태"]
photos = [("p1", "대표 장비 고해상도 컷 (레이저 또는 펀치)", "저해상도", "판금"), ("p1, 23", "로고 원본 (ai/eps, 국문·영문, 흰색 버전)", "PPT 추출본", "경영관리"),
          ("p2", "대표이사 사진 (배경 제거본)", "없음", "경영관리"), ("p7, 8", "판금 대표 장비 (MURATEC 펀치)", "저해상도", "판금"),
          ("p10, 22", "브랜드 로고 5종 (MURATEC·HSG·ADIRA·KYOWA·AIZAWA)", "없음", "판금"), ("p11", "브랜드별 대표 장비 사진 (모델당 1장)", "없음", "판금"),
          ("p12", "본사 전시장 최신 사진", "2025.03본", "판금"), ("p7, 13", "MURATEC CNC 선반 대표 컷", "저해상도", "공작"),
          ("p14", "시리즈별 대표 모델 사진", "없음", "공작"), ("p15", "고객사 로고 (게재 가능 업체)", "2025.03본", "공작"),
          ("p16", "대구 공장 항공사진 (대구공장 여부 확인)", "있음", "자동차"), ("p17", "공장 배치도 최신본", "2025.03본", "자동차"),
          ("p18", "생산 품목 사진 (전 품목 재촬영)", "캡처본", "자동차"), ("p19", "자동화 라인 사진 (로봇·갠트리)", "있음", "자동차"),
          ("p20", "인증서 최신 사본 스캔 (IATF·ISO 14001·T-SQ·P-SQ·SSQ)", "2025.03본", "자동차"), ("p20", "검사 장비 / 측정실 사진", "없음", "자동차"),
          ("p21", "설치·교육·A/S 현장 사진", "없음", "판금·공작"), ("p22", "글로벌 네트워크 지도 (브랜드 확정 후 재제작)", "2025.03본", "디자인")]
first, last = table(ws, 4, H, [5, 10, 50, 14, 12, 34, 10], [(i + 1, a, b, c, d) for i, (a, b, c, d) in enumerate(photos)],
                    example=("예시", "p11", "HSG 평판 레이저", "없음", "판금", "p11_판금_HSG_레이저_01.jpg", "예시"),
                    input_cols={6}, status_col=7, blank_rows=3, wrap_cols=(3,))
REG["사진_체크리스트"] = ("G", first, last)

# ---------------- 진행현황 ----------------
ws = dash
title(ws, "진행현황", "각 시트의 '상태' 칸을 자동 집계합니다 (예시 행 제외, '해당없음'은 완료로 계산)")
H = ["시트", "구간", "전체 항목", "완료", "진행중", "미착수", "해당없음", "진행률"]
for i, (h, w) in enumerate(zip(H, [24, 12, 11, 9, 9, 9, 10, 12]), 1):
    c = ws.cell(4, i, h); c.font = font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center"); c.border = BOX
    ws.column_dimensions[get_column_letter(i)].width = w
r = 5
for key, (col, f, l) in REG.items():
    sheet = key.split("(")[0]
    part = key[len(sheet):].strip("()") or "전체"
    rng = f"'{sheet}'!${col}${f}:${col}${l}"
    vals = [sheet, part, f"=COUNTA({rng})", f'=COUNTIF({rng},"완료")', f'=COUNTIF({rng},"진행중")', f'=COUNTIF({rng},"미착수")', f'=COUNTIF({rng},"해당없음")', f"=IFERROR((D{r}+G{r})/C{r},0)"]
    for i, v in enumerate(vals, 1):
        c = ws.cell(r, i, v); c.border = BOX; c.font = font()
    ws.cell(r, 1).hyperlink = f"#'{sheet}'!A1"; ws.cell(r, 1).font = font(color=BLUE, underline="single")
    ws.cell(r, 8).number_format = "0%"
    r += 1
ws.cell(r, 1, "합계").font = font(bold=True)
for cc in range(3, 8):
    col = get_column_letter(cc); ws.cell(r, cc, f"=SUM({col}5:{col}{r-1})").font = font(bold=True)
ws.cell(r, 8, f"=IFERROR((D{r}+G{r})/C{r},0)").number_format = "0%"; ws.cell(r, 8).font = font(bold=True)
for cc in range(1, 9): ws.cell(r, cc).border = BOX; ws.cell(r, cc).fill = PatternFill("solid", fgColor=PALE)
from openpyxl.formatting.rule import DataBarRule
ws.conditional_formatting.add(f"H5:H{r}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="5B9BD5"))
ws.freeze_panes = "A5"

for s in wb.worksheets:
    s.sheet_view.showGridLines = False
    s.sheet_properties.tabColor = {"안내": NAVY, "진행현황": NAVY}.get(s.title, BLUE if s.title.startswith(("공통", "사진")) else "6E7F96")
    s.page_setup.orientation = "landscape"; s.page_setup.fitToWidth = 1; s.page_setup.fitToHeight = 0
    s.sheet_properties.pageSetUpPr.fitToPage = True
wb.save(OUT)
print("saved", OUT)
