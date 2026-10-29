# -*- coding: utf-8 -*-
"""특수교육통계 제출자료 비교기

학교에서 제출한 엑셀(나이스 '제출자료 조회')과 특수교육통계 프로젝트가 만든 엑셀을
학교별로 짝지어 모든 시트, 모든 셀을 비교하고, 2분할 화면으로 볼 수 있는
HTML 보고서(인터넷 없이 열림)를 만든다.

- 파이썬 표준 라이브러리만 사용한다 (설치할 패키지 없음, 오프라인 실행).
- 숫자는 엑셀 파일에 저장된 글자 그대로를 Decimal 로 읽어 비교하므로
  부동소수점 오차가 없다 (예: 1.0 과 1 은 같고, 0.1+0.2 같은 오차도 없다).
- 파일에 들어 있는 값만 읽고 비교한다. 다른 자료를 추가하거나 추정하지 않는다.

사용법
  python compare.py                       # 기본 폴더(학교제출, 통계결과) 일괄 비교
  python compare.py -s 학교폴더 -t 통계폴더 -o 결과폴더
  python compare.py --pair 학교파일.xlsx 통계파일.xlsx
"""

import argparse
import csv
import datetime as _dt
import json
import os
import posixpath
import re
import sys
import unicodedata
import webbrowser
import zipfile
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation

VERSION = "1.1.0"
HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "viewer", "index.html")

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R_ID = "{%s}id" % NS["r"]


# ---------------------------------------------------------------- 엑셀 읽기

def col_to_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def num_to_col(n):
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


_REF = re.compile(r"^\$?([A-Z]+)\$?(\d+)$")


def parse_ref(ref):
    m = _REF.match(ref.upper())
    return int(m.group(2)), col_to_num(m.group(1))


def _text_of(el):
    """<si>/<is> 요소의 글자(서식 run 포함)를 이어 붙인다. 윗주(rPh)는 제외."""
    rph = set()
    for r in el.findall(".//m:rPh", NS):
        for t in r.iter("{%s}t" % NS["m"]):
            rph.add(id(t))
    return "".join(t.text or "" for t in el.iter("{%s}t" % NS["m"]) if id(t) not in rph)


def read_xlsx(path):
    """xlsx 파일을 읽어 시트 목록을 돌려준다.

    각 셀은 ('n', 원문숫자글자) / ('s', 글자) / ('b', 'TRUE'|'FALSE') /
    ('e', 오류값) / ('f', 수식) 중 하나다.
    """
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())

        def load(p):
            return ET.fromstring(z.read(p))

        rels = {}
        rels_path = "xl/_rels/workbook.xml.rels"
        if rels_path in names:
            for rel in load(rels_path).findall("rel:Relationship", NS):
                target = rel.get("Target")
                if target.startswith("/"):
                    target = target[1:]
                else:
                    target = posixpath.normpath(posixpath.join("xl", target))
                rels[rel.get("Id")] = target

        shared = []
        if "xl/sharedStrings.xml" in names:
            for si in load("xl/sharedStrings.xml").findall("m:si", NS):
                shared.append(_text_of(si))

        sheets = []
        wb = load("xl/workbook.xml")
        for sh in wb.findall("m:sheets/m:sheet", NS):
            name = sh.get("name")
            target = rels.get(sh.get(R_ID))
            if not target or target not in names:
                continue
            sheets.append(_read_sheet(name, load(target), shared))
        return sheets


def _read_sheet(name, root, shared):
    cells = {}
    max_r = max_c = 0
    for c in root.iterfind("m:sheetData/m:row/m:c", NS):
        ref = c.get("r")
        if not ref:
            continue  # r 없는 셀은 나이스/엑셀 출력에 없음
        r, col = parse_ref(ref)
        t = c.get("t", "n")
        v = c.find("m:v", NS)
        f = c.find("m:f", NS)
        val = None
        if t == "s":
            if v is not None and v.text is not None:
                val = ("s", shared[int(v.text)])
        elif t == "inlineStr":
            is_ = c.find("m:is", NS)
            if is_ is not None:
                val = ("s", _text_of(is_))
        elif t == "str":
            val = ("s", v.text or "") if v is not None else None
        elif t == "b":
            if v is not None:
                val = ("b", "TRUE" if v.text == "1" else "FALSE")
        elif t == "e":
            if v is not None:
                val = ("e", v.text)
        else:  # n (숫자) 또는 날짜 서식 숫자
            if v is not None and v.text is not None and v.text.strip() != "":
                val = ("n", v.text.strip())
        if val is None and f is not None:
            val = ("f", "=" + (f.text or ""))  # 계산값이 저장되지 않은 수식
        if val is None:
            continue
        if val[0] == "s" and val[1] == "":
            continue
        cells[(r, col)] = val
        max_r, max_c = max(max_r, r), max(max_c, col)

    merges = []
    for mc in root.iterfind("m:mergeCells/m:mergeCell", NS):
        a, _, b = mc.get("ref").partition(":")
        r1, c1 = parse_ref(a)
        r2, c2 = parse_ref(b or a)
        merges.append((r1, c1, r2, c2))
        max_r, max_c = max(max_r, r2), max(max_c, c2)

    widths = {}
    for col in root.iterfind("m:cols/m:col", NS):
        w = col.get("width")
        if w:
            for i in range(int(col.get("min")), min(int(col.get("max")), 200) + 1):
                widths[i] = float(w)
    return {"name": name, "cells": cells, "merges": merges,
            "max_r": max_r, "max_c": max_c, "widths": widths}


# ---------------------------------------------------------------- 값 정규화

_NUM_TEXT = re.compile(r"^[+-]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$")


def norm_text(s):
    s = unicodedata.normalize("NFC", s)
    return re.sub(r"\s+", " ", s).strip()


def to_decimal(s):
    try:
        d = Decimal(s)
    except InvalidOperation:
        return None
    return d if d.is_finite() else None  # 'NaN', 'Infinity' 같은 글자는 숫자로 보지 않는다


def classify(cell):
    """셀을 비교용 (종류, 값)으로 바꾼다.
    숫자처럼 생긴 글자('6', '1,234')는 숫자로 본다 — 나이스 출력이 일부 숫자를 글자로 저장하기 때문."""
    if cell is None:
        return ("empty", None)
    kind, raw = cell
    if kind == "n":
        d = to_decimal(raw)
        return ("num", d) if d is not None else ("text", raw)
    if kind == "s":
        t = norm_text(raw)
        if t == "":
            return ("empty", None)
        if _NUM_TEXT.match(t):
            return ("num", Decimal(t.replace(",", "")))
        return ("text", t)
    return (kind, raw)  # b, e, f


def fmt_decimal(d):
    if d == d.to_integral_value():
        return str(d.quantize(Decimal(1)))
    return format(d.normalize(), "f")


def fmt_delta(d):
    s = fmt_decimal(d)
    return s if s.startswith("-") else "+" + s


def display(cell):
    if cell is None:
        return ""
    kind, raw = cell
    if kind == "n":
        d = to_decimal(raw)
        return fmt_decimal(d) if d is not None else raw
    return raw


# ---------------------------------------------------------------- 머리글 찾기

def _merge_lookup(sheet):
    owner = {}
    for (r1, c1, r2, c2) in sheet["merges"]:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                owner[(r, c)] = (r1, c1)
    return owner


def _label_at(sheet, owner, r, c):
    anchor = owner.get((r, c), (r, c))
    k, v = classify(sheet["cells"].get(anchor))
    return v if k == "text" else None


def labels_for(sheet, owner, r, c):
    """셀 (r,c)의 행 제목(왼쪽 글자)과 열 제목(위쪽 글자)을 만든다."""
    row_parts, col_parts = [], []
    for cc in range(1, c):
        t = _label_at(sheet, owner, r, cc)
        if t and (not row_parts or row_parts[-1] != t):
            row_parts.append(t)
    for rr in range(1, r):
        t = _label_at(sheet, owner, rr, c)
        if t and (not col_parts or col_parts[-1] != t):
            col_parts.append(t)
    return " > ".join(row_parts), " > ".join(col_parts)


# ---------------------------------------------------------------- 비교

TYPE_NUM = "숫자 다름"
TYPE_BLANK = "한쪽만 값 있음"
TYPE_TEXT = "글자 다름"
TYPE_KIND = "형식 다름"
TYPE_SHEET = "시트 없음"


def compare_cells(a, b):
    """두 셀을 비교해 같으면 None, 다르면 (유형, 차이) 를 돌려준다."""
    ka, va = classify(a)
    kb, vb = classify(b)
    if ka == "empty" and kb == "empty":
        return None
    if ka == "num" and kb == "num":
        return None if va == vb else (TYPE_NUM, vb - va)
    if ka == "empty" or kb == "empty":
        return (TYPE_BLANK, None)
    if ka == kb:
        return None if va == vb else (TYPE_TEXT, None)
    return (TYPE_KIND, None)


def sheet_key(name):
    """'I-1.  특수학급 ...' -> 'I-1' 처럼 시트 번호를 뽑는다. 번호가 없으면 정규화한 이름."""
    m = re.match(r"^\s*([IVX]+-\d+)\s*\.", name)
    return m.group(1) if m else norm_text(name)


def pack_sheet(sheet):
    if sheet is None:
        return None
    return {
        "maxR": sheet["max_r"], "maxC": sheet["max_c"],
        "cells": {"%d,%d" % rc: display(v) for rc, v in sheet["cells"].items()},
        "kinds": {"%d,%d" % rc: v[0] for rc, v in sheet["cells"].items() if v[0] != "n"},
        "merges": [list(m) for m in sheet["merges"]],
        "widths": {str(k): v for k, v in sheet["widths"].items() if k <= sheet["max_c"]},
    }


def compare_books(school_sheets, stats_sheets):
    """시트 이름(번호)으로 짝지어 모든 셀을 비교한다."""
    left = {}
    for s in school_sheets:
        left.setdefault(sheet_key(s["name"]), s)
    right = {}
    for s in stats_sheets:
        right.setdefault(sheet_key(s["name"]), s)

    order = [sheet_key(s["name"]) for s in school_sheets]
    order += [k for k in (sheet_key(s["name"]) for s in stats_sheets) if k not in left]

    result = []
    for key in order:
        a, b = left.get(key), right.get(key)
        entry = {"key": key, "name": (a or b)["name"],
                 "left": pack_sheet(a), "right": pack_sheet(b), "diffs": []}
        if a is None or b is None:
            entry["diffs"].append({
                "r": 1, "c": 1, "addr": "-", "type": TYPE_SHEET,
                "lv": "" if a is None else "(있음)", "rv": "" if b is None else "(있음)",
                "delta": "", "rowLabel": "",
                "colLabel": "학교 제출 파일에 시트 없음" if a is None else "통계 결과 파일에 시트 없음",
            })
            result.append(entry)
            continue
        owner_a, owner_b = _merge_lookup(a), _merge_lookup(b)
        max_r, max_c = max(a["max_r"], b["max_r"]), max(a["max_c"], b["max_c"])
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                ca, cb = a["cells"].get((r, c)), b["cells"].get((r, c))
                if ca is None and cb is None:
                    continue
                res = compare_cells(ca, cb)
                if res is None:
                    continue
                typ, delta = res
                row_label, col_label = labels_for(a, owner_a, r, c)
                if not row_label and not col_label:
                    row_label, col_label = labels_for(b, owner_b, r, c)
                entry["diffs"].append({
                    "r": r, "c": c, "addr": "%s%d" % (num_to_col(c), r), "type": typ,
                    "lv": display(ca), "rv": display(cb),
                    "delta": fmt_delta(delta) if delta is not None else "",
                    "rowLabel": row_label, "colLabel": col_label,
                })
        result.append(entry)
    return result


# ---------------------------------------------------------------- 학교 짝짓기

_LEVELS = [("초등학교", "초"), ("중학교", "중"), ("고등학교", "고"), ("특수학교", "특")]
_NOISE = ["제출자료 조회", "제출자료조회", "특수교육통계", "제출자료", "비교결과"]


def school_name_from_file(path):
    stem = os.path.splitext(os.path.basename(path))[0]
    stem = unicodedata.normalize("NFC", stem)
    for w in _NOISE:
        stem = stem.replace(w, " ")
    stem = re.sub(r"\(.*?\)|\[.*?\]", " ", stem)
    stem = re.sub(r"\d{4,}", " ", stem)  # 날짜·연도
    parts = [p for p in re.split(r"[_\-\s]+", stem) if p]
    for p in parts:  # 학교급이 들어 있는 조각을 학교명으로
        if re.search(r"(초등학교|중학교|고등학교|학교|초|중|고)$", p):
            return p
    return "".join(parts)


def school_keys(name):
    k = re.sub(r"\s+", "", name)
    for long, short in _LEVELS:
        if k.endswith(long):
            k = k[: -len(long)] + short
    k1 = k
    k2 = re.sub(r"^(서울특별시|서울시|서울)", "", k)
    return k1, k2


def list_xlsx(folder):
    out = []
    for fn in sorted(os.listdir(folder)):
        if fn.startswith("~$") or fn.startswith("."):
            continue
        if fn.lower().endswith((".xlsx", ".xlsm")):
            out.append(os.path.join(folder, fn))
    return out


def load_mapping(path):
    """매칭표.csv: 학교제출파일명,통계결과파일명 (선택)"""
    mapping = {}
    if path and os.path.exists(path):
        with open(path, encoding="utf-8-sig") as fh:
            for row in csv.reader(fh):
                if len(row) >= 2 and row[0].strip() and not row[0].startswith("#"):
                    mapping[row[0].strip()] = row[1].strip()
    return mapping


def pair_files(school_files, stats_files, mapping=None):
    mapping = mapping or {}
    pairs, used = [], set()
    stats_by_name = {os.path.basename(p): p for p in stats_files}
    info = {p: school_keys(school_name_from_file(p)) for p in stats_files}

    for sp in school_files:
        base = os.path.basename(sp)
        name = school_name_from_file(sp)
        match = None
        if base in mapping and mapping[base] in stats_by_name:
            match = stats_by_name[mapping[base]]
        else:
            k1, k2 = school_keys(name)
            free = [p for p in stats_files if p not in used]
            for pick in (lambda q: info[q][0] == k1,
                         lambda q: info[q][1] == k2,
                         lambda q: info[q][1].endswith(k2) or k2.endswith(info[q][1])):
                cands = [q for q in free if pick(q)]
                if len(cands) == 1:
                    match = cands[0]
                    break
                if len(cands) > 1:
                    break  # 애매하면 짝짓지 않고 매칭표로 지정하게 한다
        if match:
            used.add(match)
        pairs.append((name, sp, match))
    leftover = [p for p in stats_files if p not in used]
    return pairs, leftover


# ---------------------------------------------------------------- 실행

def run(pairs, leftover_stats, out_dir, title):
    schools = []
    for name, sp, tp in pairs:
        item = {"name": name, "schoolFile": os.path.basename(sp),
                "statsFile": os.path.basename(tp) if tp else "", "sheets": [], "error": ""}
        try:
            if not tp:
                item["error"] = "짝이 되는 통계 결과 파일을 찾지 못했습니다."
            else:
                item["sheets"] = compare_books(read_xlsx(sp), read_xlsx(tp))
        except Exception as exc:  # 손상된 파일 등
            item["error"] = "파일을 읽지 못했습니다: %s" % exc
        item["diffCount"] = sum(len(s["diffs"]) for s in item["sheets"])
        schools.append(item)
        status = item["error"] or ("일치" if item["diffCount"] == 0 else "확인필요 %d건" % item["diffCount"])
        print("  - %-14s %s" % (name, status))

    for tp in leftover_stats:
        schools.append({"name": school_name_from_file(tp), "schoolFile": "",
                        "statsFile": os.path.basename(tp), "sheets": [], "diffCount": 0,
                        "error": "짝이 되는 학교 제출 파일을 찾지 못했습니다."})

    data = {"title": title, "version": VERSION,
            "generated": _dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "schools": schools}

    os.makedirs(out_dir, exist_ok=True)
    stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M")
    html_path = os.path.join(out_dir, "비교결과_%s.html" % stamp)
    json_path = os.path.join(out_dir, "비교결과_%s.json" % stamp)
    csv_path = os.path.join(out_dir, "확인필요목록_%s.csv" % stamp)

    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    with open(TEMPLATE, encoding="utf-8") as fh:
        template = fh.read()
    safe = payload.replace("</", "<\\/").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")
    html = template.replace("/*__DATA__*/null", safe, 1)
    with open(html_path, "w", encoding="utf-8") as fh:
        fh.write(html)
    with open(json_path, "w", encoding="utf-8") as fh:
        fh.write(payload)
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["학교", "시트", "셀", "행 제목", "열 제목", "학교 제출 값", "통계 결과 값", "차이(통계-학교)", "유형"])
        for s in schools:
            if s["error"]:
                w.writerow([s["name"], "", "", "", "", s["schoolFile"], s["statsFile"], "", s["error"]])
            for sh in s["sheets"]:
                for d in sh["diffs"]:
                    w.writerow([s["name"], sh["name"], d["addr"], d["rowLabel"], d["colLabel"],
                                d["lv"], d["rv"], d["delta"], d["type"]])
    return html_path, json_path, csv_path, data


def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    ap = argparse.ArgumentParser(description="특수교육통계 제출자료 비교기")
    ap.add_argument("-s", "--school", default=os.path.join(HERE, "학교제출"),
                    help="학교에서 제출한 엑셀 폴더 (기본: 학교제출)")
    ap.add_argument("-t", "--stats", default=os.path.join(HERE, "통계결과"),
                    help="특수교육통계 프로젝트 결과 엑셀 폴더 (기본: 통계결과)")
    ap.add_argument("-o", "--out", default=os.path.join(HERE, "비교결과"),
                    help="결과를 저장할 폴더 (기본: 비교결과)")
    ap.add_argument("-m", "--map", default=None,
                    help="학교 짝 지정 CSV (학교제출파일명,통계결과파일명). 기본: 학교제출 폴더의 매칭표.csv")
    ap.add_argument("--pair", nargs=2, metavar=("학교파일", "통계파일"), help="파일 두 개만 비교")
    ap.add_argument("--title", default="특수교육통계 제출자료 비교")
    ap.add_argument("--open", action="store_true", help="끝나면 보고서를 브라우저로 연다")
    args = ap.parse_args(argv)

    if args.pair:
        sp, tp = args.pair
        pairs, leftover = [(school_name_from_file(sp), sp, tp)], []
    else:
        for folder in (args.school, args.stats):
            if not os.path.isdir(folder):
                os.makedirs(folder, exist_ok=True)
                print("폴더를 만들었습니다: %s  (여기에 엑셀 파일을 넣고 다시 실행하세요)" % folder)
        sf, tf = list_xlsx(args.school), list_xlsx(args.stats)
        for folder in (args.school, args.stats):
            old = [f for f in os.listdir(folder) if f.lower().endswith(".xls")]
            if old:
                print("주의: .xls(옛 형식) 파일 %d개는 건너뜁니다. 엑셀에서 .xlsx로 다시 저장해 주세요: %s"
                      % (len(old), ", ".join(old[:5])))
        if not sf or not tf:
            print("비교할 파일이 없습니다. 학교제출 %d개, 통계결과 %d개" % (len(sf), len(tf)))
            return 1
        mapping = load_mapping(args.map or os.path.join(args.school, "매칭표.csv"))
        pairs, leftover = pair_files(sf, tf, mapping)
        print("학교 제출 %d개, 통계 결과 %d개를 비교합니다." % (len(sf), len(tf)))

    html_path, json_path, csv_path, data = run(pairs, leftover, args.out, args.title)
    total = sum(s["diffCount"] for s in data["schools"])
    bad = sum(1 for s in data["schools"] if s["diffCount"] or s["error"])
    print("\n완료: %d개교 중 %d개교 확인 필요, 확인필요 셀 %d건" % (len(data["schools"]), bad, total))
    print("보고서: %s" % html_path)
    print("목록(CSV): %s" % csv_path)
    if args.open:
        webbrowser.open("file:///" + os.path.abspath(html_path).replace(os.sep, "/").lstrip("/"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
