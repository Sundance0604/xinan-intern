# -*- coding: utf-8 -*-
"""直接解析 xlsx 底层 XML 的读取器。

系统导出的「交易查询」等表格 styles.xml 常损坏，openpyxl/pandas 读不了。
这里用 zipfile + ElementTree 读 sheet XML 和 sharedStrings，全程不碰 styles.xml。

用法:
    from xlsx_reader import read_sheet
    header, rows = read_sheet(path)
"""
import re
import zipfile
from xml.etree import ElementTree as ET

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def read_sheet(path):
    """解析 xlsx 首个工作表，返回 (header, rows)，每行为字符串列表。"""
    z = zipfile.ZipFile(path)

    sst = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall(f"{{{MAIN_NS}}}si"):
            sst.append("".join(t.text or "" for t in si.iter(f"{{{MAIN_NS}}}t")))

    sheets = sorted(n for n in z.namelist()
                    if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n))
    if not sheets:
        raise RuntimeError("未找到工作表 xl/worksheets/sheetN.xml")
    root = ET.fromstring(z.read(sheets[0]))

    out_rows = []
    max_col = 0
    for row in root.iter(f"{{{MAIN_NS}}}row"):
        cells = {}
        for c in row.findall(f"{{{MAIN_NS}}}c"):
            ref = c.get("r") or ""
            m = re.match(r"([A-Z]+)\d+", ref)
            if not m:
                continue
            idx = _col_index(m.group(1))
            max_col = max(max_col, idx)
            t = c.get("t")
            v = c.find(f"{{{MAIN_NS}}}v")
            if t == "inlineStr":
                is_el = c.find(f"{{{MAIN_NS}}}is")
                val = "".join(x.text or "" for x in is_el.iter(f"{{{MAIN_NS}}}t")) if is_el is not None else ""
            elif v is None:
                val = ""
            elif t == "s":
                val = sst[int(v.text)] if v.text else ""
            else:
                val = v.text or ""
            cells[idx] = val
        if cells:
            out_rows.append([cells.get(i, "") for i in range(1, max_col + 1)])

    return (out_rows[0] if out_rows else []), (out_rows[1:] if out_rows else [])
