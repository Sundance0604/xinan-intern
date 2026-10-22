# -*- coding: utf-8 -*-
"""把 extract_qq_docs.js 输出的 JSON 转成 xlsx（每个 sheet 一个工作表）。

用法: python json_to_xlsx.py [输入json] [输出xlsx]
默认: 读 系统临时目录/qq_all_sheets.json
      写 本地数据目录/持仓表_YYYYMMDD.xlsx（当天一份，次日即成为「前一日持仓表」）

数值处理:
  - 腾讯文档已算好的公式结果本身就是数字，直接按数字写。
  - 下列「数值列」里的文本也转成数字（百分比列存为小数并套百分比格式）：
    票面、买入收益率、买入数量、结算金额、当日中债估值、资本利得、每日票息、剩余规模、
    部门本周资金成本、中低/AA/AA+ 相关、票息总额、每日融资收益率、加权平均收益率；
    第二个 sheet 的场外/场内/RTGS/资金占用净增加/资金占用余额/资金成本率/计划内资金使用成本。
  - 其余列（简称、代码、日期、担保、主办人、舆情等）保持文本；代码保留前导 0。
"""
import json, sys, datetime, os, tempfile
from openpyxl import Workbook
from copy import copy as _copy

IN = sys.argv[1] if len(sys.argv) > 1 else os.path.join(tempfile.gettempdir(), 'qq_all_sheets.json')
# 只保留按日一份：今天的表就是明天的「前一日持仓表」
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.environ.get('SWSC_DATA_DIR', os.path.join(os.path.dirname(os.path.dirname(__file__)), 'local_data')), f'持仓表-{datetime.date.today():%Y-%m-%d}.xlsx')


def norm_header(s):
    return (str(s).replace('\n', '').replace('\r', '').replace(' ', '')
            .replace('（', '(').replace('）', ')').strip())


# 需要按数字存储的列（表头归一化后匹配）
NUMERIC_HEADERS = {norm_header(x) for x in [
    "票面", "买入收益率", "买入数量（万）", "结算金额（元）", "当日中债估值（%）",
    "资本利得", "每日票息", "剩余规模", "部门本周资金成本",
    "中低评级合计（合计不超15亿）", "AA持仓数量", "AA结算金额（事业部不超5亿）",
    "AA+持仓数量", "AA+结算金额", "票息总额", "每日融资收益率", "加权平均收益率",
    # 第二个 sheet：资金占用积数统计表
    "场外（银行间）", "场内交易债券（除RTGS)", "RTGS", "资金占用净增加",
    "资金占用余额", "资金成本率", "计划内资金使用成本",
]}
# 百分比列：存为小数（3.28% -> 0.0328）并套百分比格式，Excel 里仍显示 3.28%
PERCENT_HEADERS = {norm_header(x) for x in [
    "票面", "买入收益率", "当日中债估值（%）", "每日融资收益率", "加权平均收益率",
    "资金成本率", "部门本周资金成本",
]}

BLANKS = {"", "-", "--", "—", "/"}


def to_number(val):
    """把文本转成数字。返回 (number, is_percent) 或 None（不是数字）。"""
    if isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        return float(val), False
    s = str(val).strip().replace(',', '').replace('，', '')
    if s in BLANKS:
        return None
    is_pct = s.endswith('%') or s.endswith('％')
    if is_pct:
        s = s[:-1].strip()
    try:
        n = float(s)
    except ValueError:
        return None
    return (n / 100.0 if is_pct else n), is_pct


def header_row_index(rows):
    """找到表头行：前 5 行里非空文本最多的那行。"""
    best_i, best_n = 0, -1
    for i, row in enumerate(rows[:5]):
        n = sum(1 for v in row if isinstance(v, str) and v.strip())
        if n > best_n:
            best_i, best_n = i, n
    return best_i


data = json.load(open(IN, encoding='utf-8'))

wb = Workbook()
wb.remove(wb.active)

for sh in data['sheets']:
    ws = wb.create_sheet(title=sh['name'][:31])
    rows = sh['rows']
    hr = header_row_index(rows)
    headers = {c: norm_header(v) for c, v in enumerate(rows[hr]) if isinstance(v, str) and v.strip()}

    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c)
            hdr = headers.get(c - 1)
            if r == hr + 1:                           # 表头行：文字按文本，个别数值格按数字
                cell.value = '' if val is None else val
                if isinstance(val, str):
                    cell.data_type = 's'
            elif isinstance(val, bool):
                cell.value = str(val)
                cell.data_type = 's'
            elif isinstance(val, (int, float)):
                cell.value = val                      # 公式计算值：数字
                if hdr in PERCENT_HEADERS:
                    cell.number_format = '0.0000%'
            elif hdr in NUMERIC_HEADERS:
                num = to_number(val)
                if num is None:
                    # 数值列里的 "-" 等占位符：写空单元格，不写文本
                    cell.value = None
                else:
                    cell.value = num[0]
                    cell.number_format = '0.0000%' if (num[1] or hdr in PERCENT_HEADERS) else 'General'
            else:
                cell.value = '' if val is None else val
                cell.data_type = 's'                  # 文本（代码等）保留前导 0，避免 "=" 变公式

    # 表头加粗
    for c in range(1, sh['colCount'] + 1):
        _f = _copy(ws.cell(row=hr + 1, column=c).font)
        _f.bold = True
        ws.cell(row=hr + 1, column=c).font = _f

try:
    os.makedirs(os.path.dirname(os.path.abspath(OUT)), exist_ok=True)
    wb.save(OUT)
    print('saved:', OUT)
except PermissionError as e:
    sys.exit(f'写入失败（文件可能正在 Excel/WPS 中打开，请先关闭后重试）: {e.filename}')
print('sheets:', [(s['name'], s['rowCount'], s['colCount']) for s in data['sheets']])
