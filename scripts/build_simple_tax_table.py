#!/usr/bin/env python3
"""근로소득 간이세액표(소득세법 시행령 별표2)를 법제처 뷰어 텍스트 레이어에서 표로 만든다.

사용법:
    python scripts/build_simple_tax_table.py <xml폴더> <출력.json> [--dump 줄덤프.txt]

<xml폴더>: law.go.kr 별표 뷰어의 thumbnailxml/0.xml, 1.xml, ... 를 받아 둔 폴더
           (예: https://www.law.go.kr/viewer/BYL/LS/.../SKIN/<key>/thumbnailxml/<n>.xml)
출력 JSON: {"rows": [[이상, 미만, [가족1..11 세액]], ...]}  (단위: 월급여 천원, 세액 원, '-'는 0)
"""
import json
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

NUM = re.compile(r"^-$|^[0-9][0-9,]*$")


def page_lines(xml_path):
    """글자 단위 <text>를 줄(같은 t) 단위로 묶고, 가까운 글자끼리 토큰으로 합친다."""
    root = ET.parse(xml_path).getroot()
    chars = []
    for el in root.iter("text"):
        ch = (el.text or "")
        w = float(el.get("w", 0))
        if not ch.strip() or w <= 0:
            continue
        chars.append((float(el.get("t")), float(el.get("l")), w, ch))
    chars.sort()
    lines, cur, cur_t = [], [], None
    for t, l, w, ch in chars:
        if cur_t is None or abs(t - cur_t) > 4:
            if cur:
                lines.append(cur)
            cur, cur_t = [], t
        cur.append((l, w, ch))
    if cur:
        lines.append(cur)

    out = []
    for line in lines:
        line.sort()
        tokens, tok, prev_end = [], "", None
        for l, w, ch in line:
            if prev_end is not None and l - prev_end > 6:
                tokens.append(tok)
                tok = ""
            tok += ch
            prev_end = l + w
        if tok:
            tokens.append(tok)
        out.append(tokens)
    return out


def to_int(tok):
    return 0 if tok == "-" else int(tok.replace(",", ""))


def main(argv):
    if len(argv) < 3:
        print(__doc__)
        return 2
    xml_dir, out_path = Path(argv[1]), Path(argv[2])
    dump = Path(argv[argv.index("--dump") + 1]) if "--dump" in argv else None

    files = sorted(xml_dir.glob("*.xml"), key=lambda p: int(p.stem))
    rows, dump_lines = [], []
    for f in files:
        for tokens in page_lines(f):
            dump_lines.append(f"[{f.stem}] " + " | ".join(tokens))
            if len(tokens) == 13 and all(NUM.match(t) for t in tokens) and tokens[0] != "-":
                lo, hi = to_int(tokens[0]), to_int(tokens[1])
                rows.append([lo, hi, [to_int(t) for t in tokens[2:]]])

    if dump:
        dump.write_text("\n".join(dump_lines), encoding="utf-8")

    # 검증: 구간이 빈틈 없이 이어지는지
    problems = [f"{a[1]} != {b[0]}" for a, b in zip(rows, rows[1:]) if a[1] != b[0]]
    out_path.write_text(json.dumps({"rows": rows}, ensure_ascii=False), encoding="utf-8")
    print(f"행 {len(rows)}개: {rows[0][0]}~{rows[-1][1]}천원, 끊김 {len(problems)}건")
    for p in problems[:20]:
        print("  끊김:", p)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
