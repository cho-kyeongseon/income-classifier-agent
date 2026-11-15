#!/usr/bin/env python3
"""'/소득구분 업체명' 실행에 필요한 서류 리스트 엑셀을 만든다.

사용법:
    python scripts/make_required_docs_xlsx.py [출력.xlsx]
    (기본: 작업 폴더 바로 아래 소득구분_필요서류_리스트.xlsx)

시트
  1_실행 전 준비      순서대로 할 일 + 체크
  2_직원별 체크       직원마다 한 줄. 소득구분을 고르면 필요 없는 칸은 회색, 받은 서류는 ✔, 진행률 자동
  3_소득대장 항목     소득대장에 있어야 할 열 + 체크
  4_소득구분별 증빙   소득구분별 판단 기준·서류·질문·근거 + 체크
  (숨김) _기준        소득구분 × 서류 적용표

내용을 고칠 때는 이 파일의 표(PREP, LEDGER, EVIDENCE, DOCS, TYPES)를 고치고 다시 실행한다.
"""
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
CHECK = "✔"
WARNING = ("⚠ 주의할 점: 이 파일은 빈 양식이라 GitHub에 올라가요. 직원 이름을 채워서 쓸 때는 "
           "'다른 이름으로 저장'으로 따로 저장하세요(예: 해인커피_서류체크.xlsx). "
           "이 파일 그대로 저장하면 이름 같은 개인정보가 GitHub에 함께 올라갈 수 있어요.")
STATUS_LIST = f'"{CHECK},요청함,없음"'
STAFF_ROWS = 200  # 직원별 체크 시트에 미리 서식을 넣어 둘 줄 수

# ── 1. 실행 전 준비 ───────────────────────────────────────────
PREP = [
    ["순서", "준비할 것", "어떻게 / 어디에", "필수", "메모", "체크"],
    [1, "소득대장 파일 (.xlsx / .xls / .csv)", "작업 폴더의 data/raw/ 에 넣는다", "필수",
     "원본은 GitHub에 올라가지 않음(.gitignore)", ""],
    [2, "파일 이름 (업체명 안 넣어도 됨)",
     "/소득구분 업체명 이라고 시키면 결과는 그 업체명으로 만들어짐.\n"
     "자동으로 고르게 하려면: data/raw/<업체명>/ 폴더에 넣거나 파일 이름에 업체명을 넣는다", "-",
     "data/raw/에 여러 업체 파일이 섞여 있고 이름으로 구분이 안 되면, 어느 파일인지 Claude가 물어봄", ""],
    [3, "소득대장 항목 확인", "'3_소득대장 항목' 시트의 필수 항목이 있는지 본다", "필수",
     "열 이름은 달라도 됨(의미로 맞춤). 없는 항목은 [확인 필요]로 남고 보류됨", ""],
    [4, "직원별 서류 체크", "'2_직원별 체크' 시트에 직원을 한 줄씩 적고 받은 서류를 ✔", "권장",
     "없는 서류가 있어도 실행은 됨. 그 사람은 보류되고 받을 서류가 카톡 초안에 정리됨", ""],
    [5, "실행", "Claude 데스크톱 앱에서 이 폴더로 세션을 열고  /소득구분 업체명", "필수",
     "실행 PC에 Python 3 + openpyxl, pandas 필요", ""],
    [6, "결과 확인", "output/<업체명>_<실행시각>/<업체명>_소득구분_결과.xlsx", "-",
     "필요서류·기준·계산내역·카톡 초안은 같은 폴더의 01~04 폴더", ""],
]
NOTES = [
    "• '/소득구분 업체명'은 소득대장만 있어도 실행돼요. 서류가 부족한 사람은 '보류'로 남고, 받아야 할 서류가 01_필요서류와 카톡 초안에 정리돼요.",
    "• 대표님이 요청한 소득구분(예: '일용직으로', '기타소득으로')은 판단 근거가 아니에요. 사실관계와 서류로 판단해요.",
    "• 세율·근거는 2026-11-15 법령 확인 기준이에요. 법이 바뀌면 B가 다시 확인해요.",
    "• 체크: 칸을 누르면 오른쪽에 ▼가 생겨요. ✔(받음) / 요청함 / 없음 중에서 고르세요. 지우려면 Delete.",
]

# ── 3. 소득대장 항목 ──────────────────────────────────────────
LEDGER = [
    ["항목", "필수 여부", "어디에 쓰나", "없으면", "체크"],
    ["성명", "필수", "사람별로 묶기", "실행 불가", ""],
    ["주민등록번호(또는 생년월일)", "권장", "동명이인 구분, 내외국인·거주자 확인 (결과에는 뒷자리 가림)", "[확인 필요]", ""],
    ["지급일 (실제 지급한 날)", "필수", "원천징수 시점, 신고 월, 간이세액표 판 선택(2026.3.1 전후)", "귀속월로 대신하되 [확인 필요]", ""],
    ["귀속월", "급여는 필수", "어느 달 급여인지", "[확인 필요]", ""],
    ["지급액 (세전)", "필수", "세액 계산", "실행 불가", ""],
    ["지급 내용 / 명목", "필수", "소득구분 판단 (급여, 일당, 디자인 용역, 강의료 등)", "보류", ""],
    ["계약 형태", "권장", "정규직 / 아르바이트 / 프리랜서 / 1회성 구분", "[확인 필요]", ""],
    ["근무일 (날짜별 기록) 또는 근무일수", "일용은 필수", "일당 공제(1일 15만원) 계산, 같은 고용주 3개월 이상 계속 근무 여부", "일용 계산 보류", ""],
    ["급여 구성 (기본급·연장수당·식대·자가운전보조금 등)", "상용은 필수", "비과세 금액을 빼고 과세 월급여 계산", "상용 계산 보류", ""],
    ["4대보험 본인부담액", "상용은 권장", "급여 검토용(간이세액표에는 이미 반영됨)", "-", ""],
    ["지급 횟수 / 주기", "권장", "계속·반복성(사업소득 vs 기타소득)", "[확인 필요]", ""],
    ["사업자등록 여부", "프리랜서는 권장", "사업소득 판단, 세금계산서 대상 여부", "[확인 필요]", ""],
    ["비고 / 고객메모", "선택", "대표님 요청·특이사항 (예: '일용직으로 처리해 주세요')", "-", ""],
]

# ── 4. 소득구분별 증빙 ────────────────────────────────────────
EVIDENCE = [
    ["소득구분", "이렇게 판단해요", "받아야 할 서류", "꼭 확인할 질문", "근거", "체크"],
    ["공통 (모든 사람)", "-",
     "신분증 사본\n통장 사본\n외국인은 외국인등록증",
     "내국인/외국인인가?\n국내 거주자(국내 주소 또는 183일 이상 거소)인가?",
     "소득세법 제1조의2", ""],
    ["근로_상용 (정규직·계속 근무)", "고용관계가 있고 같은 고용주에게 3개월 이상 계속 고용",
     "근로계약서\n급여대장 / 급여명세서\n공제대상가족 수 (본인 포함)\n8세 이상 20세 이하 자녀 수\n원천징수 80%/120% 신청서 (신청한 경우만)",
     "공제대상가족 몇 명인가? (본인·배우자 각 1명)\n8~20세 자녀는 몇 명인가?\n상여금이 있나? (스크립트 미지원 → 보류)",
     "법 제20조, 제129조③\n시행령 제189조·별표2, 제194조", ""],
    ["  └ 식대 비과세", "식사를 따로 제공받지 않는 근로자의 월 20만원 이하 식사대",
     "현물 식사 제공 여부 확인",
     "회사에서 식사를 따로 주나?\n월 20만원을 넘는 부분은 과세",
     "법 제12조제3호러목", ""],
    ["  └ 자가운전보조금 비과세", "본인 차량을 직접 운전해 업무에 쓰고, 실제 여비 대신 받는 월 20만원 이내",
     "차량 등록증 (본인 명의 또는 본인 임차)\n회사 지급기준(사규)\n업무 사용 확인",
     "본인 명의 차량인가?\n업무에 쓰나?\n실제 출장비(유류비 등)를 따로 받나? → 받으면 비과세 안 됨",
     "시행령 제12조제3호", ""],
    ["근로_일용 (아르바이트·일당)", "일·시간 단위로 받고, 같은 고용주에게 3개월 이상 계속 고용되지 않음 (건설공사는 1년)",
     "근무기록 (출근부·작업일보, 날짜별)\n일당 내역\n일용 근로계약서",
     "같은 곳에서 3개월 이상 계속 일했나? → 그렇다면 상용일 수 있음\n일한 날짜와 일당은?",
     "시행령 제20조①\n법 제47조②, 제59조③, 제129조①4", ""],
    ["사업 (프리랜서 3.3%)", "고용관계 없이 독립적으로, 직업으로 계속·반복하는 인적용역",
     "용역(위촉)계약서\n다른 거래처 여부 / 지급 주기\n사업자등록증 (있는 경우)",
     "이 일을 직업으로 계속하나?\n다른 곳에서도 같은 일을 하나?\n사업자등록이 있나? (있으면 세금계산서 대상일 수 있음)",
     "법 제19조, 제129조①3\n시행령 제184조, 제149조의3", ""],
    ["기타 (20%)", "고용관계 없이 일시적으로 받은 대가",
     "지급 내용 증빙 (계약서·의뢰서)\n1회성임을 확인할 자료",
     "이번 한 번뿐인가?\n건별 기타소득금액 5만원 이하면 과세 안 함",
     "법 제21조, 제129조①6라, 제84조", ""],
    ["기타_필요경비60 (강연·원고·심사 등)", "일시적 강연료, 방송 해설·심사, 원고료·인세 등 → 필요경비 60%",
     "강의 의뢰서 / 원고 / 위촉 공문\n실제 경비 영수증 (60%를 넘는 경우만)",
     "강의·원고가 그 사람의 직업인가? → 직업이면 사업소득\n1회성인가?",
     "시행령 제87조제1호의2\n법 제21조①15·19", ""],
    ["이자 (14%)", "돈을 빌려주고 받은 이자", "금전대차 약정서",
     "영업으로 빌려준 것인가? (비영업대금이면 25%)", "법 제129조①1", ""],
    ["배당 (14%)", "법인의 이익 배당", "배당 결의서 (주주총회·이사회)",
     "출자공동사업자 배당인가? (25%)", "법 제129조①2", ""],
    ["특수관계인 (대표 가족 등)", "실제로 일했는지가 먼저",
     "출퇴근 기록\n업무 결과물\n급여 이체 내역\n4대보험 가입 내역",
     "실제로 근무했나? 입증되기 전에는 보류", "-", ""],
]

# ── 2. 직원별 체크: 소득구분 × 서류 ────────────────────────────
TYPES = ["근로_상용", "근로_일용", "사업", "기타", "기타_필요경비60", "이자", "배당", "모름"]
# (그룹, 서류 이름, 입력 방식, 적용 소득구분)   입력 방식: check=✔ 드롭다운, num=숫자
ALL = set(TYPES)
LABOR = {"근로_상용", "근로_일용", "모름"}
FREE = {"사업", "기타", "기타_필요경비60", "모름"}
DOCS = [
    ("공통", "신분증 사본", "check", ALL),
    ("공통", "통장 사본", "check", ALL),
    ("공통", "내외국인·거주자 확인", "check", ALL),
    ("근로", "근로계약서", "check", LABOR),
    ("상용", "급여대장·명세서", "check", {"근로_상용"}),
    ("상용", "공제대상가족 수\n(본인 포함, 숫자)", "num", {"근로_상용"}),
    ("상용", "8~20세 자녀 수\n(숫자, 없으면 0)", "num", {"근로_상용"}),
    ("상용", "식대: 식사 제공 여부", "check", {"근로_상용"}),
    ("상용", "자가운전: 본인 차량·사규·실비 미지급", "check", {"근로_상용"}),
    ("일용", "근무기록(날짜별)·일당", "check", {"근로_일용"}),
    ("사업·기타", "용역(위촉)계약서·의뢰서", "check", FREE),
    ("사업·기타", "계속·반복 여부(다른 거래처·주기)", "check", FREE),
    ("사업·기타", "사업자등록증(있으면)", "check", {"사업"}),
    ("사업·기타", "강의의뢰서·원고·위촉 공문", "check", {"기타_필요경비60"}),
    ("이자·배당", "금전대차 약정서", "check", {"이자"}),
    ("이자·배당", "배당 결의서", "check", {"배당"}),
]
SPECIAL = ("특수관계", "실제 근무 입증자료\n(출퇴근·업무·이체)")  # '대표 가족' = 예 일 때만


def build(path):
    wb = Workbook()
    head_fill = PatternFill("solid", fgColor="DDEBF7")
    group_fill = PatternFill("solid", fgColor="BDD7EE")
    must_fill = PatternFill("solid", fgColor="FCE4D6")
    sub_fill = PatternFill("solid", fgColor="F2F2F2")
    input_fill = PatternFill("solid", fgColor="FFFFFF")
    # 조건부 서식(dxf) 채우기는 bgColor로 줘야 엑셀에 색이 나온다
    na_fill = PatternFill(bgColor="D9D9D9")
    ok_fill = PatternFill(bgColor="C6EFCE")
    req_fill = PatternFill(bgColor="FFEB9C")
    no_fill = PatternFill(bgColor="FFC7CE")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def check_rules(ws, rng):
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{rng.split(":")[0]}="{CHECK}"'], fill=ok_fill))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{rng.split(":")[0]}="요청함"'], fill=req_fill))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{rng.split(":")[0]}="없음"'], fill=no_fill))

    def sheet(ws, title, rows, widths, must_col=None, check_col=None):
        ws.append([title])
        ws["A1"].font = Font(bold=True, size=13)
        ws.append([])
        for r in rows:
            ws.append(r)
        hdr = 3
        for c in ws[hdr]:
            c.font, c.fill, c.border, c.alignment = Font(bold=True), head_fill, border, center
        for r in ws.iter_rows(min_row=hdr + 1, max_row=ws.max_row):
            for c in r:
                c.border = border
                c.alignment = Alignment(vertical="top", wrap_text=True)
            if must_col is not None and str(r[must_col].value or "").startswith("필수"):
                r[must_col].fill, r[must_col].font = must_fill, Font(bold=True)
            if str(r[0].value or "").startswith("  └"):
                for c in r:
                    c.fill = sub_fill
        if check_col is not None:
            col = get_column_letter(check_col + 1)
            rng = f"{col}{hdr + 1}:{col}{ws.max_row}"
            dv = DataValidation(type="list", formula1=STATUS_LIST, allow_blank=True)
            dv.promptTitle, dv.prompt = "체크", "✔ 받음 / 요청함 / 없음"
            ws.add_data_validation(dv)
            dv.add(rng)
            for (cell,) in ws[rng]:
                cell.alignment = center
                cell.font = Font(bold=True, size=12)
            check_rules(ws, rng)
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = ws.cell(row=hdr + 1, column=1)
        ws.page_setup.orientation = "landscape"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0

    # 1. 실행 전 준비
    ws = wb.active
    ws.title = "1_실행 전 준비"
    sheet(ws, "/소득구분 업체명 — 실행 전 준비", PREP, [6, 30, 58, 8, 50, 8], must_col=3, check_col=5)
    ws["A2"] = WARNING
    ws.merge_cells("A2:F2")
    ws["A2"].font = Font(bold=True, color="C00000")
    ws["A2"].fill = PatternFill("solid", fgColor="FFF2CC")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 36
    ws.append([])
    for n in NOTES:
        ws.append([n])
        ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=6)
        ws.cell(row=ws.max_row, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[ws.max_row].height = 30

    # 숨김: 적용표
    ref = wb.create_sheet("_기준")
    ref.append(["소득구분"] + [d[1].split("\n")[0] for d in DOCS])
    for t in TYPES:
        ref.append([t] + [1 if t in d[3] else 0 for d in DOCS])
    last_doc_col = get_column_letter(1 + len(DOCS))
    wb.defined_names["TypeList"] = DefinedName("TypeList", attr_text=f"'_기준'!$A$2:$A${1 + len(TYPES)}")
    wb.defined_names["ApplyTable"] = DefinedName(
        "ApplyTable", attr_text=f"'_기준'!$B$2:${last_doc_col}${1 + len(TYPES)}")
    ref.sheet_state = "hidden"

    # 2. 직원별 체크
    ws = wb.create_sheet("2_직원별 체크", 1)
    ws["A1"] = "직원별 서류 체크"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"], ws["B2"] = "업체명", ""
    ws["A2"].font = Font(bold=True)
    ws["B2"].fill, ws["B2"].border = must_fill, border
    ws.merge_cells("B2:C2")
    ws["E2"] = ("사용법: 직원을 한 줄씩 적고(줄 추가는 그냥 아래 줄에 입력) 소득구분을 고르면, 필요 없는 칸은 회색이 돼요. "
                "받은 서류는 ▼에서 ✔, 숫자 칸은 숫자를 적어요. 진행률은 자동으로 계산돼요.")
    ws.merge_cells(start_row=2, start_column=5, end_row=2, end_column=5 + len(DOCS))
    ws["E2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 34
    ws["A3"] = WARNING
    ws["A3"].font = Font(bold=True, color="C00000")
    ws["A3"].fill = PatternFill("solid", fgColor="FFF2CC")
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[3].height = 30

    fixed = ["번호", "성명", "소득구분", "대표 가족\n(예/아니오)"]
    first_doc = len(fixed) + 1                       # E
    special_col = first_doc + len(DOCS)              # 특수관계 칸
    progress_col = special_col + 1
    memo_col = progress_col + 1
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=memo_col)
    G, H = 4, 5                                      # 그룹 줄, 제목 줄
    start, end = H + 1, H + STAFF_ROWS

    for i, name in enumerate(fixed, 1):
        ws.cell(row=G, column=i, value="기본")
        ws.cell(row=H, column=i, value=name)
    for j, (grp, name, _, _) in enumerate(DOCS):
        ws.cell(row=G, column=first_doc + j, value=grp)
        ws.cell(row=H, column=first_doc + j, value=name)
    ws.cell(row=G, column=special_col, value=SPECIAL[0])
    ws.cell(row=H, column=special_col, value=SPECIAL[1])
    ws.cell(row=G, column=progress_col, value="결과")
    ws.cell(row=H, column=progress_col, value="진행률")
    ws.cell(row=G, column=memo_col, value="결과")
    ws.cell(row=H, column=memo_col, value="메모")
    # 그룹 줄 병합
    col = 1
    while col <= memo_col:
        val = ws.cell(row=G, column=col).value
        end_col = col
        while end_col + 1 <= memo_col and ws.cell(row=G, column=end_col + 1).value == val:
            end_col += 1
        if end_col > col:
            ws.merge_cells(start_row=G, start_column=col, end_row=G, end_column=end_col)
        col = end_col + 1
    for r in (G, H):
        for c in range(1, memo_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font, cell.border, cell.alignment = Font(bold=True), border, center
            cell.fill = group_fill if r == G else head_fill
    ws.row_dimensions[H].height = 48

    # 입력 칸 서식
    for r in range(start, end + 1):
        for c in range(1, memo_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            cell.alignment = center if c != memo_col and c != 2 else Alignment(vertical="center", wrap_text=True)
            if c >= first_doc and c <= special_col:
                cell.font = Font(bold=True, size=12)
        ws.cell(row=r, column=1).value = f'=IF(B{r}="","",ROW()-{H})'

    L = get_column_letter
    fd, ld, sc, pc = L(first_doc), L(special_col - 1), L(special_col), L(progress_col)

    # 드롭다운
    dv_type = DataValidation(type="list", formula1="=TypeList", allow_blank=True)
    dv_type.promptTitle, dv_type.prompt = "소득구분", "모르면 '모름'"
    dv_fam = DataValidation(type="list", formula1='"예,아니오"', allow_blank=True)
    dv_chk = DataValidation(type="list", formula1=STATUS_LIST, allow_blank=True)
    dv_chk.promptTitle, dv_chk.prompt = "체크", "✔ 받음 / 요청함 / 없음"
    dv_num = DataValidation(type="whole", operator="between", formula1="0", formula2="30", allow_blank=True)
    dv_num.promptTitle, dv_num.prompt = "숫자", "숫자로 입력 (예: 3)"
    dv_num.error = "0~30 사이 숫자를 넣어 주세요"
    for dv in (dv_type, dv_fam, dv_chk, dv_num):
        ws.add_data_validation(dv)
    dv_type.add(f"C{start}:C{end}")
    dv_fam.add(f"D{start}:D{end}")
    for j, (_, _, kind, _) in enumerate(DOCS):
        rng = f"{L(first_doc + j)}{start}:{L(first_doc + j)}{end}"
        (dv_num if kind == "num" else dv_chk).add(rng)
    dv_chk.add(f"{sc}{start}:{sc}{end}")

    # 조건부 서식: 해당 없음(회색) → 받음/요청함/없음 색
    doc_rng = f"{fd}{start}:{ld}{end}"
    ws.conditional_formatting.add(doc_rng, FormulaRule(
        formula=[f'AND($C{start}<>"",INDEX(ApplyTable,MATCH($C{start},TypeList,0),COLUMN()-{first_doc - 1})=0)'],
        fill=na_fill, font=Font(color="A6A6A6"), stopIfTrue=True))
    ws.conditional_formatting.add(f"{sc}{start}:{sc}{end}", FormulaRule(
        formula=[f'$D{start}<>"예"'], fill=na_fill, font=Font(color="A6A6A6"), stopIfTrue=True))
    all_rng = f"{fd}{start}:{sc}{end}"
    ws.conditional_formatting.add(all_rng, FormulaRule(formula=[f"ISNUMBER({fd}{start})"], fill=ok_fill))
    check_rules(ws, all_rng)
    ws.conditional_formatting.add(f"{pc}{start}:{pc}{end}", FormulaRule(
        formula=[f"AND(ISNUMBER({pc}{start}),{pc}{start}>=1)"], fill=ok_fill, font=Font(bold=True)))

    # 진행률
    for r in range(start, end + 1):
        row_app = f"INDEX(ApplyTable,MATCH($C{r},TypeList,0),0)"
        done = f'SUMPRODUCT({row_app}*((({fd}{r}:{ld}{r}="{CHECK}")+ISNUMBER({fd}{r}:{ld}{r}))>0))'
        need = f"SUM({row_app})"
        sp_done, sp_need = f'IF($D{r}="예",--({sc}{r}="{CHECK}"),0)', f'IF($D{r}="예",1,0)'
        ws.cell(row=r, column=progress_col).value = (
            f'=IF($B{r}="","",IF($C{r}="","소득구분 선택",IFERROR(({done}+{sp_done})/({need}+{sp_need}),"")))')
        ws.cell(row=r, column=progress_col).number_format = "0%"

    widths = [6, 12, 16, 10] + [12] * len(DOCS) + [14, 9, 30]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[L(i)].width = w
    ws.freeze_panes = ws.cell(row=start, column=3)
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.print_title_rows = f"{G}:{H}"

    # 3, 4
    ws = wb.create_sheet("3_소득대장 항목")
    sheet(ws, "소득대장에 있어야 할 항목 (열 이름은 달라도 됨)", LEDGER, [40, 16, 62, 26, 8], must_col=1, check_col=4)
    ws = wb.create_sheet("4_소득구분별 증빙")
    sheet(ws, "소득구분별로 받아야 할 서류 (대표님께 요청할 것)", EVIDENCE, [30, 40, 42, 50, 30, 8], check_col=5)

    for s in wb.worksheets:
        s.sheet_view.tabSelected = False
    wb.active = 1  # 열면 직원별 체크부터
    wb.active.sheet_view.tabSelected = True
    wb.save(path)


def main(argv):
    out = Path(argv[1]) if len(argv) > 1 else ROOT / "소득구분_필요서류_리스트.xlsx"
    build(out)
    print(f"저장: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
