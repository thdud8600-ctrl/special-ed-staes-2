# -*- coding: utf-8 -*-
"""python -m unittest discover tests"""
import os
import sys
import tempfile
import unittest
import zipfile
from decimal import Decimal

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import compare  # noqa: E402

M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def make_xlsx(path, sheets):
    """sheets: [(이름, {"A1": 값, ...}, ["A1:B1", ...])] 로 최소 xlsx 를 만든다."""
    strings = []

    def cell(ref, v):
        if isinstance(v, str):
            strings.append(v)
            return '<c r="%s" t="s"><v>%d</v></c>' % (ref, len(strings) - 1)
        return '<c r="%s"><v>%s</v></c>' % (ref, v)

    with zipfile.ZipFile(path, "w") as z:
        rels, entries = [], []
        for i, (name, cells, merges) in enumerate(sheets, 1):
            rows = {}
            for ref, v in cells.items():
                rows.setdefault(compare.parse_ref(ref)[0], []).append(cell(ref, v))
            data = "".join('<row r="%d">%s</row>' % (r, "".join(cs)) for r, cs in sorted(rows.items()))
            mc = "".join('<mergeCell ref="%s"/>' % m for m in merges)
            z.writestr("xl/worksheets/sheet%d.xml" % i,
                       '<worksheet xmlns="%s"><sheetData>%s</sheetData>%s</worksheet>'
                       % (M, data, "<mergeCells>%s</mergeCells>" % mc if mc else ""))
            rels.append('<Relationship Id="rId%d" Target="worksheets/sheet%d.xml" Type="x"/>' % (i, i))
            entries.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (name, i, i))
        z.writestr("xl/workbook.xml",
                   '<workbook xmlns="%s" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                   "<sheets>%s</sheets></workbook>" % (M, "".join(entries)))
        z.writestr("xl/_rels/workbook.xml.rels",
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>'
                   % "".join(rels))
        z.writestr("xl/sharedStrings.xml",
                   '<sst xmlns="%s">%s</sst>' % (M, "".join("<si><t>%s</t></si>" % s for s in strings)))


class CellCompare(unittest.TestCase):
    def test_numbers_exact(self):
        self.assertIsNone(compare.compare_cells(("n", "1.0"), ("n", "1")))
        self.assertIsNone(compare.compare_cells(("s", "6"), ("n", "6")))  # 나이스가 글자로 저장한 숫자
        self.assertEqual(compare.compare_cells(("n", "7.0"), ("n", "9"))[1], Decimal(2))
        self.assertEqual(compare.compare_cells(("n", "0.3"), ("n", "0.30000000000000004"))[0], compare.TYPE_NUM)

    def test_blank_vs_zero_is_flagged(self):
        self.assertEqual(compare.compare_cells(None, ("n", "0"))[0], compare.TYPE_BLANK)

    def test_text_whitespace_ignored(self):
        self.assertIsNone(compare.compare_cells(("s", "장애\r\n유형"), ("s", "장애 유형")))
        self.assertEqual(compare.compare_cells(("s", "총계"), ("s", "합계"))[0], compare.TYPE_TEXT)


class Pairing(unittest.TestCase):
    def test_names(self):
        self.assertEqual(compare.school_name_from_file("제출자료 조회_서울동답초등학교.xlsx"), "서울동답초등학교")
        self.assertEqual(compare.school_name_from_file("동답초_특수교육통계_제출자료.xlsx"), "동답초")

    def test_pair(self):
        school = ["/a/제출자료 조회_서울동답초등학교.xlsx", "/a/제출자료 조회_서울가나초등학교.xlsx"]
        stats = ["/b/가나초_특수교육통계_제출자료.xlsx", "/b/동답초_특수교육통계_제출자료.xlsx", "/b/다라초_특수교육통계_제출자료.xlsx"]
        pairs, left = compare.pair_files(school, stats)
        self.assertEqual([os.path.basename(p[2]) for p in pairs],
                         ["동답초_특수교육통계_제출자료.xlsx", "가나초_특수교육통계_제출자료.xlsx"])
        self.assertEqual([os.path.basename(p) for p in left], ["다라초_특수교육통계_제출자료.xlsx"])


class Workbooks(unittest.TestCase):
    def test_end_to_end(self):
        with tempfile.TemporaryDirectory() as d:
            a, b = os.path.join(d, "a.xlsx"), os.path.join(d, "b.xlsx")
            head = {"A1": "장애유형", "B1": "전체", "C1": "여"}
            make_xlsx(a, [("I-1.  학급  학생 수", dict(head, A2="총계", B2=7, C2="2"), ["A1:A1"]),
                          ("I-2.  학년", {"A1": 1}, [])])
            make_xlsx(b, [("I-1.  학급  학생 수", dict(head, A2="총계", B2=9, C2=2), ["A1:A1"])])
            res = compare.compare_books(compare.read_xlsx(a), compare.read_xlsx(b))
            d1 = res[0]["diffs"]
            self.assertEqual(len(d1), 1)
            self.assertEqual((d1[0]["addr"], d1[0]["lv"], d1[0]["rv"], d1[0]["delta"]), ("B2", "7", "9", "+2"))
            self.assertEqual((d1[0]["rowLabel"], d1[0]["colLabel"]), ("총계", "전체"))
            self.assertEqual(res[1]["diffs"][0]["type"], compare.TYPE_SHEET)


if __name__ == "__main__":
    unittest.main()
