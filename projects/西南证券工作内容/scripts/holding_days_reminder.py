#!/usr/bin/env python3
"""持仓工作日提醒：读取持仓表第一个 sheet，找出持仓工作日 >= 13 天的债券（李新鹏不限），
按主办人分组生成 txt。

用法:
  python holding_days_reminder.py [持仓表.xlsx] [--days 13] [--out 输出.txt] [--date YYYY-MM-DD]

默认:
  持仓表 = 本地数据目录/持仓表_YYYYMMDD.xlsx（取最新一份，即当日下午抓取生成的当日表）
  输出   = 本地数据目录/持仓工作日提醒_YYYY-MM-DD.txt
  date   = 今天（等价于公式里的 TODAY()）

规则:
  - 持仓工作日 = NETWORKDAYS(买入结算日, date, 节假日)。节假日从 L 列公式
    =NETWORKDAYS(...,$A$lo:$A$hi) 里读列 A 那段日期（中秋国庆假等）。
  - 输出"天数" = 持仓工作日 + 1。
  - 部门券 -> 叶总（"叶总早上好"）；李新鹏 -> 不设天数门槛、开头不加姓名（"早上好"）；
    其他人 -> 姓+老师（"X老师早上好"）。
"""
import sys, os, re, glob, datetime, argparse

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

DESK = os.environ.get('SWSC_DATA_DIR', os.path.join(os.path.dirname(os.path.dirname(__file__)), 'local_data'))
OUT_DIR = DESK


def latest_holdings():
    """当日的持仓表（持仓表-YYYY-MM-DD.xlsx，取日期最新一份）。"""
    best = None
    for f in glob.glob(os.path.join(DESK, "持仓表[-_]*.xlsx")):
        base = os.path.basename(f)
        m = re.search(r"持仓表[-_](\d{4})[-_]?(\d{2})[-_]?(\d{2})", base)
        if m:
            key = (m.group(1) + m.group(2) + m.group(3),
                   1 if re.match(r"持仓表-\d{4}-\d{2}-\d{2}", base) else 0)
            if not best or key > best[0]:
                best = (key, f)
    if best:
        return best[1]
    legacy = os.path.join(DESK, "销售交易部持仓表.xlsx")
    return legacy if os.path.exists(legacy) else os.path.join(DESK, "持仓表-0000-00-00.xlsx")

# 主办人 -> (称呼前缀, 天数门槛)。门槛 None 表示不设门槛（全部发送）。
OWNER_RULES = {
    "部门券": ("叶总", 13),
    "李新鹏": ("", None),
}
DEFAULT_THRESHOLD = 13


def parse_date(s):
    if s is None:
        return None
    if isinstance(s, datetime.datetime):
        return s.date()
    if isinstance(s, datetime.date):
        return s
    s = str(s).strip()
    if not s:
        return None
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y%m%d",
                "%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    m = re.match(r"(\d{4})/(\d{1,2})/(\d{1,2})", s)
    if m:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def networkdays(start, end, holidays):
    """等价 Excel NETWORKDAYS(start, end, holidays)：含首尾，排除周末和节假日。"""
    if end < start:
        return 0
    hd = set(holidays)
    n, d = 0, start
    while d <= end:
        if d.weekday() < 5 and d not in hd:
            n += 1
        d += datetime.timedelta(days=1)
    return n


def cell_number(v):
    """单元格里的数值（新抓取的表里公式已存为计算值）；公式文本/空值返回 None。"""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip() if v is not None else ""
    if not s or s.startswith("="):
        return None
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return None


def find_col(header, *names):
    for i, h in enumerate(header):
        hs = str(h or "").strip()
        for n in names:
            if hs == n or hs.replace("(", "（").replace(")", "）") == n:
                return i
    return None


def parse_holidays(ws, formula_col_idx, a_col_idx, code_col_idx=None):
    """从持仓工作日公式里解析节假日范围（$A$lo:$A$hi）并读出日期。
    新抓取的表里公式已存为计算值，此时改为：扫描列 A 中持仓区之后的日期（即节假日清单）。"""
    for row in ws.iter_rows(min_row=2):
        v = row[formula_col_idx].value
        if not v:
            continue
        m = re.search(r"\$A\$(\d+):\$A\$(\d+)", str(v))
        if m:
            lo, hi = int(m.group(1)), int(m.group(2))
            return [d for r in range(lo, hi + 1)
                    for d in [parse_date(ws.cell(row=r, column=a_col_idx + 1).value)] if d]
    # 回退：找持仓区最后一行（代码列非空），其后的日期即节假日
    if code_col_idx is not None:
        last = 1
        for row in ws.iter_rows(min_row=2):
            if str(row[code_col_idx].value or "").strip():
                last = row[0].row
        return [d for r in range(last + 1, ws.max_row + 1)
                for d in [parse_date(ws.cell(row=r, column=a_col_idx + 1).value)] if d]
    return []


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("hold", nargs="?", default=latest_holdings())
    ap.add_argument("--days", type=int, default=DEFAULT_THRESHOLD)
    ap.add_argument("--out", default=None)
    ap.add_argument("--date", default=None)
    args = ap.parse_args()

    import openpyxl
    wb = openpyxl.load_workbook(args.hold, data_only=True)
    ws = wb.worksheets[0]
    header = [c.value for c in ws[1]]

    idx = {}
    for name, keys in {
        "简称": ("简称",),
        "代码": ("代码",),
        "买入结算日": ("买入结算日",),
        "持仓工作日": ("持仓工作日（参考）", "持仓工作日"),
        "主办人": ("主办人",),
    }.items():
        idx[name] = find_col(header, *keys)
        if idx[name] is None:
            sys.exit(f"表头找不到列：{name}，现有表头：{[str(h) for h in header]}")

    today = parse_date(args.date) if args.date else datetime.date.today()
    holidays = parse_holidays(ws, idx["持仓工作日"], idx["简称"], idx["代码"])

    # 收集债券行：优先用表里已算好的「持仓工作日」（新抓取的表公式已转计算值）；
    # 没有数值时再用 买入结算日 + 节假日 重算。
    bonds = []  # (出现顺序, 简称, 代码, 主办人, 持仓工作日)
    for r in ws.iter_rows(min_row=2):
        a = str(r[idx["简称"]].value or "").strip()
        code = str(r[idx["代码"]].value or "").strip()
        if not code or a.startswith("【"):
            continue
        d = parse_date(r[idx["买入结算日"]].value)
        if d is None:
            # 买入结算日为空时，NETWORKDAYS 会从 1900 年起算得出无意义的巨大值，直接跳过
            continue
        wd = cell_number(r[idx["持仓工作日"]].value)
        if wd is None:
            wd = networkdays(d, today, holidays)
        wd = int(round(wd))
        owner = str(r[idx["主办人"]].value or "").strip()
        bonds.append((len(bonds), a, code, owner, wd))

    # 分组，应用天数门槛
    groups, owner_order = {}, {}
    for order, a, code, owner, wd in bonds:
        prefix, thr = OWNER_RULES.get(owner, (None, args.days))
        if prefix is None:
            prefix = (owner[0] + "老师") if owner else ""
        if thr is not None and wd < thr:
            continue
        if owner not in groups:
            groups[owner] = []
            owner_order[owner] = order
        groups[owner].append((a, code, wd))

    if not groups:
        print("今日无持仓工作日达标的债券。")
        return

    lines = []
    for owner in sorted(groups, key=lambda o: owner_order[o]):
        prefix, _ = OWNER_RULES.get(owner, (None, args.days))
        if prefix is None:
            prefix = (owner[0] + "老师") if owner else ""
        items = sorted(groups[owner], key=lambda x: -x[2])  # 天数多的排前面
        lines.append(f"{prefix}早上好，请查看持仓工作日：")
        for a, code, wd in items:
            lines.append(f"{a}    {code}    {wd + 1}天")
        lines.append("")

    lines.append(f"—— {today.strftime('%Y-%m-%d')}")   # 文件末尾标注当日日期
    text = "\n".join(lines).rstrip() + "\n"
    out = args.out or os.path.join(OUT_DIR, f"持仓工作日提醒_{today.strftime('%Y-%m-%d')}.txt")
    with open(out, "w", encoding="utf-8") as f:
        f.write(text)
    print(text)
    print(f"[已写入] {out}")
    print(f"[统计] 表内债券 {len(bonds)} 只，达标 {sum(len(v) for v in groups.values())} 只，"
          f"节假日 {len(holidays)} 天。")


if __name__ == "__main__":
    main()
