#!/usr/bin/env python3
"""원천징수세액 계산기 (소득구분팀 C가 사용).

사용법:
    python3 scripts/calc_withholding.py <구분확정.csv> <rules.json> <출력폴더>

입력 CSV 열: 번호,성명,지급일,지급액,소득구분,근무일수,필요경비율,판단근거,상태
- 상태가 '확정'인 행만 계산한다. '보류'는 그대로 옮기고 세액은 비운다.
- 필요경비율 칸이 비어 있으면 rules.json의 expense_rate를 쓴다.

출력: <출력폴더>/계산결과.csv, <출력폴더>/계산요약.json
금액 계산은 Decimal로 하고, 단수 처리는 rules.json의 rounding을 따른다.
"""
import csv
import json
import sys
from decimal import Decimal, ROUND_FLOOR
from pathlib import Path

D = Decimal


def to_dec(value, default="0"):
    text = str(value if value is not None else "").replace(",", "").strip()
    return D(text) if text else D(default)


def truncate(amount, unit):
    unit = D(unit)
    return (amount / unit).to_integral_value(rounding=ROUND_FLOOR) * unit


class Calculator:
    def __init__(self, rules):
        self.rules = rules
        self.warnings = []
        self._warned = set()

    def _rule(self, key, section=None):
        section = section or self.rules
        if key not in section:
            raise KeyError(f"rules.json에 '{key}' 항목이 없습니다")
        item = section[key]
        if not item.get("verified", False) and key not in self._warned:
            self._warned.add(key)
            self.warnings.append(f"법령 미확인 기준 사용: {key} ({item.get('note', '')})")
        return item

    def income_tax(self, kind, amount, days, expense_rate):
        """소득구분별 소득세(단수 처리 전)를 돌려준다."""
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
        if row.get("상태", "").strip() != "확정" or kind in ("보류", "근로_상용", ""):
            out.update(소득세="", 지방소득세="", 원천징수합계="", 차인지급액="", 비고="계산 보류")
            return out, None

        days = to_dec(row.get("근무일수"), "0")
        exp_text = str(row.get("필요경비율", "")).strip()
        expense_rate = D(exp_text) if exp_text else None

        rnd = self._rule("rounding")
        tax = truncate(self.income_tax(kind, amount, days, expense_rate), rnd["unit"])
        note = []

        sae = self._rule("small_amount_exemption")
        if kind in sae["applies_to"] and D(0) < tax < D(str(sae["threshold"])):
            tax = D(0)
            note.append("소액부징수")

        local_rate = D(str(self._rule("local_tax_rate_of_income_tax")["value"]))
        local = truncate(tax * local_rate, rnd["unit"])
        total = tax + local
        out.update(
            소득세=int(tax), 지방소득세=int(local), 원천징수합계=int(total),
            차인지급액=int(amount - total), 비고=",".join(note),
        )
        return out, (kind, amount, tax, local)


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
            held.append({"번호": row.get("번호"), "성명": row.get("성명"), "소득구분": row.get("소득구분")})
            continue
        kind, amount, tax, local = calc_data
        k = by_kind.setdefault(kind, {"건수": 0, "지급액": D(0), "소득세": D(0), "지방소득세": D(0)})
        k["건수"] += 1
        for key, val in (("지급액", amount), ("소득세", tax), ("지방소득세", local)):
            k[key] += val
            tot[key] += val

    out_dir.mkdir(parents=True, exist_ok=True)
    fields = list(rows[0].keys()) + ["소득세", "지방소득세", "원천징수합계", "차인지급액", "비고"] if rows else []
    with (out_dir / "계산결과.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(results)

    def ints(d):
        return {k: (int(v) if isinstance(v, Decimal) else v) for k, v in d.items()}

    summary = {
        "계산건수": sum(k["건수"] for k in by_kind.values()),
        "보류건수": len(held),
        "총지급액_계산분": int(tot["지급액"]),
        "소득세합계": int(tot["소득세"]),
        "지방소득세합계": int(tot["지방소득세"]),
        "원천징수합계": int(tot["소득세"] + tot["지방소득세"]),
        "차인지급합계_계산분": int(tot["지급액"] - tot["소득세"] - tot["지방소득세"]),
        "소득구분별": {k: ints(v) for k, v in by_kind.items()},
        "보류": held,
        "경고": calc.warnings,
    }
    (out_dir / "계산요약.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    for w_ in calc.warnings:
        print(f"[경고] {w_}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
