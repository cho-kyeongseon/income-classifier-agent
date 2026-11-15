#!/usr/bin/env python3
"""원천징수세액 계산기 (소득구분팀 C가 사용).

사용법:
    python scripts/calc_withholding.py <구분확정.csv> <rules.json> <출력폴더>

입력 CSV 열: 번호,성명,지급일,지급액,소득구분,근무일수,필요경비율,판단근거,상태
  상용 근로(근로_상용)용 추가 열(없어도 됨):
    비과세액        지급액 중 비과세 금액(식대·자가운전보조금 등). 비우면 0
    공제대상가족수  본인 포함. 비우면 계산 보류
    자녀수          공제대상가족 중 8세 이상 20세 이하 자녀 수. 비우면 0
    원천징수비율    80 / 100 / 120 (근로자 신청 시). 비우면 100
- 상태가 '확정'인 행만 계산한다. '보류'는 그대로 옮기고 세액은 비운다.
- 필요경비율 칸이 비어 있으면 rules.json의 expense_rate를 쓴다.
- 상용 근로는 지급일(원천징수 시점)로 간이세액표 판(자녀 공제액)을 고른다.

출력: <출력폴더>/계산결과.csv, 계산요약.json
      사람별 결과 엑셀(사람별 요약·지급건별 상세·소득구분별 합계·기준과 경고):
      출력폴더가 output/<업체명>_<실행ID>/03_계산 이면 실행 폴더 바로 아래 <업체명>_소득구분_결과.xlsx,
      아니면 <출력폴더>/소득구분_결과.xlsx
금액 계산은 Decimal로 하고, 단수 처리는 rules.json의 rounding을 따른다.
"""
import csv
import json
import re
import sys
from decimal import Decimal, ROUND_FLOOR
from pathlib import Path

D = Decimal
ROOT = Path(__file__).resolve().parent.parent
TAX_COLS = ["소득세", "지방소득세", "원천징수합계", "차인지급액", "비고"]
OPTIONAL_COLS = ["비과세액", "공제대상가족수", "자녀수", "원천징수비율"]


def to_dec(value, default="0"):
    text = str(value if value is not None else "").replace(",", "").strip()
    return D(text) if text else D(default)


def truncate(amount, unit):
    unit = D(unit)
    return (amount / unit).to_integral_value(rounding=ROUND_FLOOR) * unit


def pay_date(text):
    """'2025-01(귀속월...)', '2026.03.25' 같은 값에서 YYYY-MM-DD를 뽑는다. 일이 없으면 1일."""
    m = re.search(r"(\d{4})[-./](\d{1,2})(?:[-./](\d{1,2}))?", str(text or ""))
    if not m:
        return None
    return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3) or 1):02d}"


class SimpleTaxTable:
    """근로소득 간이세액표(시행령 별표2)."""

    def __init__(self, path):
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        self.rows = data["rows"]
        self.exact_10000 = data["exact_10000"]
        self.over_10000 = data["over_10000"]
        self.child = data["child_deduction"]

    def _col(self, values, family):
        """가족 1~11명은 해당 열, 11명 초과는 별표2 제4호 산식."""
        if family <= 11:
            return D(values[family - 1])
        v10, v11 = D(values[9]), D(values[10])
        return max(D(0), v11 - (v10 - v11) * (family - 11))

    def base_tax(self, monthly_won, family):
        k = monthly_won / D(1000)  # 천원
        if k < D(self.rows[0][0]):
            return D(0)
        if k == D(10000):
            return self._col(self.exact_10000, family)
        if k > D(10000):
            base = self._col(self.exact_10000, family)
            for b in self.over_10000:
                if b["upto"] is None or k <= D(b["upto"]):
                    excess = (k - D(b["base_over"])) * D(1000)
                    return base + D(b["add"]) + excess * D(str(b["factor"])) * D(str(b["rate"]))
        for lo, hi, values in self.rows:
            if D(lo) <= k < D(hi):
                return self._col(values, family)
        raise ValueError(f"간이세액표에서 월급여 {monthly_won}원 구간을 찾지 못했습니다")

    def child_rule(self, date):
        for c in self.child:
            if date >= c["from"] and (c["to"] is None or date <= c["to"]):
                return c
        return None

    def child_deduction(self, rule, children):
        if children <= 0:
            return D(0)
        if children == 1:
            return D(rule["one"])
        return D(rule["two"]) + D(rule["per_extra"]) * (children - 2)


class Calculator:
    def __init__(self, rules):
        self.rules = rules
        self.warnings = []
        self._warned = set()
        self._table = None

    def warn(self, msg):
        if msg not in self._warned:
            self._warned.add(msg)
            self.warnings.append(msg)

    def _rule(self, key, section=None):
        section = section or self.rules
        if key not in section:
            raise KeyError(f"rules.json에 '{key}' 항목이 없습니다")
        item = section[key]
        if not item.get("verified", False):
            self.warn(f"법령 미확인 기준 사용: {key} ({item.get('note', '')})")
        return item

    def table(self):
        if self._table is None:
            t = self._rule("근로_상용", self.rules["income_types"])
            path = Path(t.get("table_file", "reference/간이세액표.json"))
            if not path.is_absolute() and not path.exists():
                path = ROOT / path
            self._table = SimpleTaxTable(path)
        return self._table

    def regular_tax(self, row, amount, note):
        """상용 근로: 간이세액표. 계산할 수 없으면 (None, 이유)."""
        fam_text = str(row.get("공제대상가족수", "")).strip()
        if not fam_text:
            return None, "공제대상가족수 필요"
        family = int(to_dec(fam_text))
        if family < 1:
            return None, "공제대상가족수는 본인 포함 1명 이상"
        children = int(to_dec(row.get("자녀수"), "0"))
        ratio = to_dec(row.get("원천징수비율"), "100")
        if ratio not in (D(80), D(100), D(120)):
            return None, "원천징수비율은 80/100/120 중 하나"
        date = pay_date(row.get("지급일"))
        if not date:
            return None, "지급일(원천징수 시점) 필요"
        tbl = self.table()
        crule = tbl.child_rule(date)
        if crule is None:
            return None, f"{date} 지급분에 맞는 간이세액표 판 없음(2024.3.1 이후만 지원)"

        nontax_text = str(row.get("비과세액", "")).strip()
        nontax = to_dec(nontax_text, "0")
        if not nontax_text:
            note.append("비과세액 미입력(0으로 계산)")
        taxable = amount - nontax
        tax = tbl.base_tax(taxable, family) - tbl.child_deduction(crule, children)
        tax = max(tax, D(0)) * ratio / D(100)
        note.append(f"간이세액표 {crule['from'][:7]}~판, 과세급여 {int(taxable):,}원, 가족 {family}명, 자녀 {children}명"
                    + (f", {int(ratio)}%" if ratio != 100 else ""))
        return tax, None

    def income_tax(self, kind, amount, days, expense_rate):
        """소득구분별 소득세(단수 처리 전)를 돌려준다. 상용 근로는 regular_tax에서."""
        t = self._rule(kind, self.rules["income_types"])
        if kind == "근로_일용":
            days = days if days > 0 else D(1)
            taxable = amount - D(str(t["daily_deduction"])) * days
            if taxable <= 0:
                return D(0)
            return taxable * D(str(t["rate"])) * (D(1) - D(str(t["tax_credit_rate"])))
        if kind.startswith("기타"):
            rate = expense_rate if expense_rate is not None else D(str(t["expense_rate"]))
            other_income = amount * (D(1) - rate)
            if other_income <= D(str(t["min_taxable_income"])):
                return D(0)
            return other_income * D(str(t["rate"]))
        return amount * D(str(t["rate"]))

    def compute_row(self, row):
        kind = row["소득구분"].strip()
        amount = to_dec(row["지급액"])
        out = dict(row)
        held = dict(소득세="", 지방소득세="", 원천징수합계="", 차인지급액="")
        if row.get("상태", "").strip() != "확정" or kind in ("보류", ""):
            out.update(held, 비고="계산 보류")
            return out, None

        note = []
        rnd = self._rule("rounding")
        if kind == "근로_상용":
            raw, why = self.regular_tax(row, amount, note)
            if raw is None:
                out.update(held, 비고=f"계산 보류: {why}")
                return out, None
        else:
            days = to_dec(row.get("근무일수"), "0")
            exp_text = str(row.get("필요경비율", "")).strip()
            expense_rate = D(exp_text) if exp_text else None
            raw = self.income_tax(kind, amount, days, expense_rate)
        tax = truncate(raw, rnd["unit"])

        sae = self._rule("small_amount_exemption")
        if kind in sae["applies_to"] and D(0) < tax < D(str(sae["threshold"])):
            tax = D(0)
            note.append("소액부징수")

        local_rate = D(str(self._rule("local_tax_rate_of_income_tax")["value"]))
        local = truncate(tax * local_rate, rnd["unit"])
        total = tax + local
        out.update(
            소득세=int(tax), 지방소득세=int(local), 원천징수합계=int(total),
            차인지급액=int(amount - total), 비고=", ".join(note),
        )
        return out, (kind, amount, tax, local)


def person_summary(results):
    """사람별로 묶는다. 키: (번호, 성명)."""
    people = {}
    for r in results:
        key = (str(r.get("번호", "")).strip(), str(r.get("성명", "")).strip())
        p = people.setdefault(key, {
            "번호": key[0], "성명": key[1], "소득구분": [], "건수": 0, "확정건수": 0, "보류건수": 0,
            "총지급액": 0, "확정지급액": 0, "소득세": 0, "지방소득세": 0, "원천징수합계": 0, "차인지급액": 0,
            "보류사유": [],
        })
        kind = str(r.get("소득구분", "")).strip()
        if kind and kind not in p["소득구분"]:
            p["소득구분"].append(kind)
        amount = int(to_dec(r.get("지급액")))
        p["건수"] += 1
        p["총지급액"] += amount
        if r.get("소득세") == "":
            p["보류건수"] += 1
            reason = str(r.get("비고", "")).replace("계산 보류: ", "").replace("계산 보류", "").strip()
            basis = str(r.get("판단근거", "")).strip()
            for text in (reason, basis if not reason else ""):
                if text and text not in p["보류사유"]:
                    p["보류사유"].append(text)
        else:
            p["확정건수"] += 1
            p["확정지급액"] += amount
            for k in ("소득세", "지방소득세", "원천징수합계", "차인지급액"):
                p[k] += int(r[k])
    for p in people.values():
        p["상태"] = "확정" if p["보류건수"] == 0 else ("보류" if p["확정건수"] == 0 else "일부 보류")
    return list(people.values())


def write_excel(path, results, people, summary, rules):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    head_fill = PatternFill("solid", fgColor="DDEBF7")
    hold_fill = PatternFill("solid", fgColor="FFF2CC")
    total_fill = PatternFill("solid", fgColor="EDEDED")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    money = "#,##0"

    def sheet(ws, headers, rows, widths, money_cols=(), wrap_cols=(), hold_col=None):
        ws.append(headers)
        for c in ws[1]:
            c.font, c.fill, c.border = Font(bold=True), head_fill, border
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in rows:
            ws.append(row)
        for r in ws.iter_rows(min_row=2):
            is_hold = hold_col is not None and str(r[hold_col].value or "") in ("보류", "일부 보류", "계산 보류")
            for c in r:
                c.border = border
                c.alignment = Alignment(vertical="top", wrap_text=c.column - 1 in wrap_cols)
                if c.column - 1 in money_cols:
                    c.number_format = money
                if is_hold:
                    c.fill = hold_fill
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    # 1. 사람별 요약
    ws = wb.active
    ws.title = "사람별 요약"
    h = ["번호", "성명", "소득구분", "상태", "지급건수", "총지급액", "확정분 지급액", "소득세", "지방소득세",
         "원천징수 합계", "차인지급액(확정분)", "보류 건수", "보류 사유 / 필요한 것"]
    rows = [[p["번호"], p["성명"], ", ".join(p["소득구분"]), p["상태"], p["건수"], p["총지급액"], p["확정지급액"],
             p["소득세"], p["지방소득세"], p["원천징수합계"], p["차인지급액"], p["보류건수"], "\n".join(p["보류사유"])]
            for p in people]
    sheet(ws, h, rows, [6, 10, 18, 10, 9, 14, 14, 12, 12, 13, 15, 9, 70],
          money_cols=(5, 6, 7, 8, 9, 10), wrap_cols=(12,), hold_col=3)
    n = len(rows) + 1
    ws.append(["", "합계", "", "", f"=SUM(E2:E{n})", f"=SUM(F2:F{n})", f"=SUM(G2:G{n})", f"=SUM(H2:H{n})",
               f"=SUM(I2:I{n})", f"=SUM(J2:J{n})", f"=SUM(K2:K{n})", f"=SUM(L2:L{n})", ""])
    for c in ws[n + 1]:
        c.font, c.fill, c.border = Font(bold=True), total_fill, border
        if c.column in (6, 7, 8, 9, 10, 11):
            c.number_format = money

    # 2. 지급건별 상세
    ws = wb.create_sheet("지급건별 상세")
    base = [k for k in results[0].keys() if k not in TAX_COLS] if results else []
    h = base + TAX_COLS
    rows = [[r.get(k, "") for k in h] for r in results]
    for row in rows:
        for i, k in enumerate(h):
            if k in ("지급액", "비과세액") and str(row[i]).strip():
                row[i] = int(to_dec(row[i]))
    widths = [{"번호": 6, "성명": 10, "지급일": 22, "판단근거": 60, "비고": 40}.get(k, 12) for k in h]
    money_cols = [i for i, k in enumerate(h) if k in ("지급액", "비과세액", "소득세", "지방소득세", "원천징수합계", "차인지급액")]
    wrap_cols = [i for i, k in enumerate(h) if k in ("판단근거", "비고", "지급일")]
    sheet(ws, h, rows, widths, money_cols=money_cols, wrap_cols=wrap_cols,
          hold_col=h.index("상태") if "상태" in h else None)

    # 3. 소득구분별 합계
    ws = wb.create_sheet("소득구분별 합계")
    h = ["소득구분", "확정 건수", "지급액", "소득세", "지방소득세", "원천징수 합계"]
    rows = [[k, v["건수"], v["지급액"], v["소득세"], v["지방소득세"], v["소득세"] + v["지방소득세"]]
            for k, v in summary["소득구분별"].items()]
    rows.append(["합계", summary["계산건수"], summary["총지급액_계산분"], summary["소득세합계"],
                 summary["지방소득세합계"], summary["원천징수합계"]])
    sheet(ws, h, rows, [16, 10, 15, 13, 13, 14], money_cols=(2, 3, 4, 5))
    for c in ws[ws.max_row]:
        c.font, c.fill = Font(bold=True), total_fill
    ws.append([])
    ws.append([f"보류 {summary['보류건수']}건은 위 합계에 들어 있지 않습니다."])

    # 4. 기준과 경고
    ws = wb.create_sheet("기준과 경고")
    h = ["항목", "확인", "근거", "메모"]
    items = [(k, rules[k]) for k in ("rounding", "local_tax_rate_of_income_tax", "small_amount_exemption") if k in rules]
    items += list(rules.get("income_types", {}).items())
    rows = [[k, "확인" if v.get("verified") else "미확인", v.get("source", ""), v.get("note", "")] for k, v in items]
    rows += [["경고", "", w, ""] for w in summary["경고"]]
    sheet(ws, h, rows, [26, 8, 60, 60], wrap_cols=(2, 3))

    wb.save(path)


def excel_path(out_dir):
    """output/<업체명>_<YYYYMMDD_HHMM>/03_계산 → output/<업체명>_<YYYYMMDD_HHMM>/<업체명>_소득구분_결과.xlsx"""
    out = out_dir.resolve()
    if out.name == "03_계산":
        m = re.match(r"^(.+)_\d{8}_\d{4}$", out.parent.name)
        if m:
            return out.parent / f"{m.group(1)}_소득구분_결과.xlsx"
    return out_dir / "소득구분_결과.xlsx"


def main(argv):
    if len(argv) != 4:
        print(__doc__)
        return 2
    src, rules_path, out_dir = Path(argv[1]), Path(argv[2]), Path(argv[3])
    rules = json.loads(rules_path.read_text(encoding="utf-8"))
    calc = Calculator(rules)

    with src.open(encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))

    results, by_kind = [], {}
    tot = {"지급액": D(0), "소득세": D(0), "지방소득세": D(0)}
    held = []
    for row in rows:
        out, calc_data = calc.compute_row(row)
        results.append(out)
        if calc_data is None:
            held.append({"번호": row.get("번호"), "성명": row.get("성명"), "소득구분": row.get("소득구분"),
                         "사유": out.get("비고", "")})
            continue
        kind, amount, tax, local = calc_data
        k = by_kind.setdefault(kind, {"건수": 0, "지급액": D(0), "소득세": D(0), "지방소득세": D(0)})
        k["건수"] += 1
        for key, val in (("지급액", amount), ("소득세", tax), ("지방소득세", local)):
            k[key] += val
            tot[key] += val

    out_dir.mkdir(parents=True, exist_ok=True)
    fields = (list(rows[0].keys()) + TAX_COLS) if rows else []
    with (out_dir / "계산결과.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(results)

    def ints(d):
        return {k: (int(v) if isinstance(v, Decimal) else v) for k, v in d.items()}

    people = person_summary(results)
    summary = {
        "계산건수": sum(k["건수"] for k in by_kind.values()),
        "보류건수": len(held),
        "총지급액_계산분": int(tot["지급액"]),
        "소득세합계": int(tot["소득세"]),
        "지방소득세합계": int(tot["지방소득세"]),
        "원천징수합계": int(tot["소득세"] + tot["지방소득세"]),
        "차인지급합계_계산분": int(tot["지급액"] - tot["소득세"] - tot["지방소득세"]),
        "소득구분별": {k: ints(v) for k, v in by_kind.items()},
        "사람별": [{k: (", ".join(v) if isinstance(v, list) else v) for k, v in p.items()} for p in people],
        "보류": held,
        "경고": calc.warnings,
    }
    (out_dir / "계산요약.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    xlsx = excel_path(out_dir)
    write_excel(xlsx, results, people, summary, rules)

    # 화면에는 짧게: 합계와 사람별 한 줄
    print(f"계산 {summary['계산건수']}건 / 보류 {summary['보류건수']}건")
    print(f"원천징수 합계 {summary['원천징수합계']:,}원 (소득세 {summary['소득세합계']:,} + 지방소득세 {summary['지방소득세합계']:,})")
    for p in people:
        print(f"  {p['번호']} {p['성명']}: {', '.join(p['소득구분'])} / {p['상태']} / 원천징수 {p['원천징수합계']:,}원")
    print(f"엑셀: {xlsx}")
    for w_ in calc.warnings:
        print(f"[경고] {w_}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
