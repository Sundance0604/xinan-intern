#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""三张表自动化：结算表(盈亏情况) / 分类表(每日结算分类) / 盈利表(当日盈利情况)。

用法:
  python pnl_tables.py <当日交易查询.xlsx> [选项]

选项:
  --prev 前一日持仓表.xlsx   默认自动找 本地数据目录/持仓表_*.xlsx 里日期 < 今天的最新一份
  --hold 当日持仓表.xlsx     默认 本地数据目录/销售交易部持仓表.xlsx
  --chai 拆借成本(元)        同业拆借成本，用户口头给（周四非零），默认 0
  --date YYYY-MM-DD          计算日，默认今天
  --json                     只打印计算 JSON，不生成 xlsx

计算口径（均已用 1015 样例反向验证）:
  - 交易量(亿元) = 非「同业拆借」行的结算金额(元)之和 / 1e8
  - 结算金额差 = 卖出结算金额 - 买入结算金额；买入价按「前一日持仓表」FIFO 匹配
    （同一代码多行，先买入结算日、再主办人，累加买入数量直到卖出数量）
  - 总结算金额差 = 各笔结算金额差之和 + 拆借成本
  - 票息收入(万元) = 持仓表「票息总额」合计
  - 分类表不计同业拆借：撮合 = 债券结算差额(未加拆借成本) - 代缴。
    结算表和盈利表仍计入拆借成本。用户确认口径见项目《成本核验注意事项.md》。
  - 日均占资(亿元) = 资金占用积数统计表「资金占用余额」自 1/1 到当日均值（余额=年初结转+F 逐日累加）
"""
import sys, os, re, glob, json, argparse, datetime
from decimal import Decimal, ROUND_HALF_UP
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_reader import read_sheet                  # 交易查询：底层 XML 解析（styles 损坏也能读）
import openpyxl
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

DESK = os.environ.get('SWSC_DATA_DIR', os.path.join(os.path.dirname(os.path.dirname(__file__)), 'local_data'))


def latest_holdings(before=None):
    """持仓表-YYYY-MM-DD.xlsx：取日期 <= before（默认不限）的最新一份。"""
    cands = []
    for f in glob.glob(os.path.join(DESK, "持仓表[-_]*.xlsx")):
        base = os.path.basename(f)
        m = re.search(r"持仓表[-_](\d{4})[-_]?(\d{2})[-_]?(\d{2})", base)
        if m:
            new_fmt = 1 if re.match(r"持仓表-\d{4}-\d{2}-\d{2}", base) else 0
            cands.append((m.group(1) + m.group(2) + m.group(3), new_fmt, f))
    if cands:
        if before:
            key = before.strftime("%Y%m%d")
            prior = [c for c in cands if c[0] <= key]
            if prior:
                return max(prior)[2]
        return max(cands)[2]
    legacy = os.path.join(DESK, "销售交易部持仓表.xlsx")
    return legacy if os.path.exists(legacy) else None


TRADE_COLS = ("交易状态", "交易品种", "交易方向", "交易日期", "代码", "名称", "交易面额(万)", "结算金额(元)",
              "托管市场", "交易编号", "交易时间", "录入员", "到期结算方式")
SUFFIX = {"上交所": ".SH", "深交所": ".SZ", "银行间": ".IB", "中债登": ".IB", "清算所": ".IB", "北交所": ".BJ"}


def num(v):
    if v is None:
        return 0.0
    s = str(v).strip().replace(",", "").replace("，", "")
    if s in ("", "-", "--", "—"):
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


def cell_number(v):
    """数值单元格（新抓取的表公式已转计算值）；公式文本/空值返回 None。"""
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


def parse_trade_time(v, default_date):
    """交易时间 '2026-10-15 13:43:17' → datetime；只有时分秒时补当天。"""
    s = str(v).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
        try:
            return datetime.datetime.strptime(s, fmt)
        except ValueError:
            pass
    m = re.search(r"(\d{1,2}):(\d{2}):(\d{2})", s)
    if m:
        return datetime.datetime.combine(
            default_date, datetime.time(int(m.group(1)), int(m.group(2)), int(m.group(3))))
    return datetime.datetime.combine(default_date, datetime.time(0, 0))


def parse_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    s = str(v).strip()
    m = re.match(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", s)
    if m:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


# ---------- 交易查询 ----------
def load_trades(path):
    header, rows = read_sheet(path)
    idx = {}
    for i, h in enumerate(header):
        if str(h).strip() in TRADE_COLS and str(h).strip() not in idx:
            idx[str(h).strip()] = i
    def g(r, k):
        i = idx.get(k)
        return (r[i] if i is not None and i < len(r) else "")
    out = []
    for r in rows:
        prod = str(g(r, "交易品种")).strip()
        code = str(g(r, "代码")).strip()
        if prod == "同业拆借":
            continue
        if not re.fullmatch(r"\d+", code):      # 合计行 / 空行
            continue
        settle = str(g(r, "到期结算方式")).strip()
        out.append({
            "交易编号": str(g(r, "交易编号")).strip(),
            "交易状态": str(g(r, "交易状态")).strip(),
            "品种": prod, "方向": str(g(r, "交易方向")).strip(), "日期": str(g(r, "交易日期")).strip(),
            "交易时间": str(g(r, "交易时间")).strip(),
            "代码": code, "名称": str(g(r, "名称")).strip(),
            "面额万": num(g(r, "交易面额(万)")), "结算金额": num(g(r, "结算金额(元)")),
            "市场": str(g(r, "托管市场")).strip(),
            "代码带后缀": code + SUFFIX.get(str(g(r, "托管市场")).strip(), ""),
            "录入员": str(g(r, "录入员")).strip(),
            "担保": "担保" if ("担保交收" in settle and "非担保" not in settle) else "非担保",
        })
    return out


_BUY_CACHE = {}


def buy_trades_on(date, trade_dir):
    """某交易日交易查询里的全部买入记录。"""
    if date in _BUY_CACHE:
        return _BUY_CACHE[date]
    mmdd = date.strftime("%m%d")
    files = sorted(glob.glob(os.path.join(trade_dir, f"{mmdd}交易查询*.xlsx")))
    buys = []
    if files:
        for t in load_trades(files[0]):
            if "买入" in t["方向"] and "卖出" not in t["方向"]:
                buys.append(t)
    _BUY_CACHE[date] = buys
    return buys


def resolve_lot_cost(lot, trade_dir):
    """一批买入的结算金额：优先用「买入结算日」当天交易查询里的实际结算金额
    （ima 笔记口径：找买入日期的交易查询做差）；找不到时回退持仓表金额。"""
    if lot.get("买入来源") == "当日买入" or not lot.get("买入结算日"):
        return lot["结算金额"], "交易查询"
    buys = [t for t in buy_trades_on(lot["买入结算日"], trade_dir)
            if t["代码"] == lot["代码"].split(".")[0]]
    if not buys:
        return lot["结算金额"], "持仓表(无当日交易查询)"
    need = lot["买入数量万"]
    exact = [t for t in buys if abs(t["面额万"] - need) < 1e-6]
    if exact:
        best = min(exact, key=lambda t: abs(t["结算金额"] - lot["结算金额"]))
        return best["结算金额"], "交易查询"
    total, rem = 0.0, need
    for t in sorted(buys, key=lambda x: x["交易时间"]):
        if rem <= 1e-9:
            break
        take = min(t["面额万"], rem)
        total += t["结算金额"] * (take / t["面额万"]) if t["面额万"] else 0.0
        rem -= take
    if rem <= 1e-9:
        return total, "交易查询"
    return lot["结算金额"], "持仓表(交易查询不足)"


def parse_pct(v):
    """把 '3.04%' / 3.04 / 0.0304 统一成小数。"""
    s = str(v).strip().replace("％", "%")
    if s.endswith("%"):
        return num(s[:-1]) / 100.0
    n = num(s)
    return n / 100.0 if n > 1 else n


# ---------- 持仓表 ----------
def load_holdings(path):
    """读持仓表第一个 sheet，返回 持仓区/代投代缴区 的行 + 汇总。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    header = {str(c.value).strip().split("\n")[0]: c.column for c in ws[1] if c.value}
    def col(*names):
        for n in names:
            for h, i in header.items():
                if h == n or h.replace("(", "（").replace(")", "）") == n:
                    return i
        return None
    cA, cB, cH, cI, cJ, cK, cN, cE, cR = (col("简称"), col("代码"), col("担保/非担保"),
                                            col("买入数量（万）"), col("结算金额（元）"),
                                            col("买入结算日"), col("主办人"), col("买入收益率"),
                                            col("每日票息"))
    def cell(r, c):
        return r[c - 1] if c else None

    rows = {"持仓": [], "代投代缴": []}
    section = None
    for r in ws.iter_rows(min_row=2):
        a = str(r[0].value).strip() if r[0].value is not None else ""
        if a == "【持仓】":
            section = "持仓"; continue
        if a == "【代投代缴】":
            section = "代投代缴"; continue
        if section is None:
            continue
        code = str(cell(r, cB).value).strip() if cell(r, cB) and cell(r, cB).value else ""
        if not code:
            if section == "代投代缴":
                section = None
            continue
        yld = parse_pct(cell(r, cE).value) if cE else 0.0
        qty = num(cell(r, cI).value) if cI else 0.0
        # 每日票息：优先用表里算好的 R 列计算值，缺失时按 收益率×数量/365×10000 重算
        px = cell_number(cell(r, cR).value) if cR else None
        if px is None:
            px = yld * qty * 10000 / 365.0
        buy_date = parse_date(cell(r, cK).value) if cK else None
        rows[section].append({
            "简称": str(cell(r, cA).value).strip() if cA and cell(r, cA).value else "",
            "代码": code, "担保": str(cell(r, cH).value).strip() if cH else "",
            "买入数量万": qty,
            "结算金额": num(cell(r, cJ).value) if cJ else 0.0,
            "买入结算日": buy_date,
            "主办人": str(cell(r, cN).value).strip() if cN else "",
            "每日票息": px,   # 元/天
            "买入来源": "持仓记录", "区块": section,
            "排序时间": datetime.datetime.combine(buy_date, datetime.time.min) if buy_date
                        else datetime.datetime(1900, 1, 1),
            "交易编号": "",
        })
    all_rows = rows["持仓"] + rows["代投代缴"]
    piaoxi_daily = sum(x["每日票息"] for x in all_rows) / 10000.0   # 仅在票息总额缺失时兜底

    # 票息收入 = 持仓表「票息总额」汇总值（【持仓】汇总行的 AA 列，公式 SUMPRODUCT(票面,数量)/365）
    cAA = col("票息总额")
    piaoxi_total = None
    if cAA:
        for r in ws.iter_rows(min_row=1):
            if str(r[0].value).strip() == "【持仓】":
                piaoxi_total = cell_number(r[cAA - 1].value)
                break

    # 汇总值：优先读表里已算好的【持仓敞口】/【拆借】/【自有资金】（K 列为数值列）
    def summary_value(label):
        for r in ws.iter_rows():
            for c in r:
                if isinstance(c.value, str) and c.value.strip() == label:
                    n = cell_number(r[10].value) if len(r) > 10 else None
                    if n is not None:
                        return n
        return None
    changkou = summary_value("【持仓敞口】")
    chai = summary_value("【拆借】")
    ziyou = summary_value("【自有资金】")
    if changkou is None:
        changkou = sum(x["结算金额"] for x in all_rows)
    if chai is None:
        chai = 5e8
    if ziyou is None:
        ziyou = changkou - chai

    return rows, {"票息收入万元": piaoxi_total if piaoxi_total is not None else piaoxi_daily,
                  "自有资金": ziyou, "持仓敞口": changkou, "拆借额度": chai}


# ---------- 资金占用积数统计表（第二 sheet）----------
def avg_fund_balance(path, target):
    """资金占用余额均值（元）。持仓文件没有第二个 sheet/对应表头时返回 None。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    if len(wb.worksheets) < 2:
        return None
    ws = wb.worksheets[1]
    hdr = {str(c.value).strip(): c.column for c in ws[2]}
    need = ("场外（银行间）", "场内交易债券（除RTGS)", "RTGS")
    if not all(k in hdr for k in need):
        return None
    cA, cB, cC, cD = 1, hdr[need[0]], hdr[need[1]], hdr[need[2]]
    # 资金占用余额 G 列：优先用表里已算好的数值；缺失时按 年初结转 + ΣF 累加
    start = num(ws.cell(row=6, column=7).value)
    vals, cur = [], start
    for r in range(6, ws.max_row + 1):
        d = parse_date(ws.cell(row=r, column=cA).value)
        if d is None:
            continue
        if d > target:
            break
        if d < datetime.date(2026, 1, 1):
            continue
        g = cell_number(ws.cell(row=r, column=7).value)
        if g is None:
            F = num(ws.cell(row=r, column=cB).value) + num(ws.cell(row=r, column=cC).value) + num(ws.cell(row=r, column=cD).value)
            cur += F
            g = cur
        else:
            cur = g
        vals.append(g)
    return sum(vals) / len(vals) if vals else 0.0


# ---------- FIFO 匹配买入价 ----------
def candidate_lots(rows, code):
    """该债券在前一日持仓表 + 当日买入里的全部可用买入批次（按 FIFO 顺序）。"""
    return sorted([x for x in rows if x["代码"].split(".")[0] == code.split(".")[0]],
                  key=lambda x: (x.get("排序时间") or datetime.datetime(1900, 1, 1),
                                 x.get("买入来源") == "当日买入",
                                 x.get("主办人", "")))


def fifo_buy_amount(cands, sell_qty_wan):
    """对给定的买入批次按 FIFO 累加买入结算金额。

    返回 (分摊买入结算金额, 匹配明细列表, 是否未匹配完)。
    匹配明细逐条给出：哪一批买入、买入结算日、主办人、该批数量/结算金额、
    本次消耗数量、分摊到的买入金额、匹配后该批剩余数量。
    """
    need, total, details = sell_qty_wan, 0.0, []
    for x in cands:
        if need <= 0:
            break
        if x["买入数量万"] <= 0:
            continue
        take = min(x["买入数量万"], need)
        cost = x.get("成本", x["结算金额"])
        part = float((Decimal(str(cost)) * Decimal(str(take)) / Decimal(str(x["买入数量万"]))).quantize(Decimal('.01'), rounding=ROUND_HALF_UP))
        details.append({
            "买入来源": x.get("买入来源", "持仓记录"), "区块": x.get("区块", ""),
            "交易编号": x.get("交易编号", ""),
            "买入简称": x["简称"], "代码": x["代码"],
            "买入结算日": x["买入结算日"].strftime("%Y-%m-%d") if x["买入结算日"] else "",
            "主办人": x["主办人"], "担保": x["担保"],
            "该批买入数量万": x["买入数量万"],
            "该批买入结算金额": round(cost, 2),
            "本次匹配数量万": take,
            "本次分摊买入金额": round(part, 2),
            "匹配后剩余数量万": round(x["买入数量万"] - take, 4),
        })
        total += part
        need -= take
    return round(total, 2), details, (need > 1e-9)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("trade")
    ap.add_argument("--trade-statuses", nargs='+', default=['成交确认'],
                    help="当日纳入的交易状态，默认仅成交确认；其他状态必须先经业务确认")
    ap.add_argument("--coupons", default=None, help="人工确认的当日付息JSON：date、source、payments（代码→元）")
    ap.add_argument("--prev", default=None)
    ap.add_argument("--hold", default=None)
    ap.add_argument("--prev-profit", default=None, help="上一交易日盈利表，读取其最新数据页")
    ap.add_argument("--prev-classification", default=None, help="上一交易日分类表，保留历史周页")
    ap.add_argument("--replay-costs", default=None, help="用户批准的历史FIFO重放配置JSON；不设置则不自动重放")
    ap.add_argument("--chai", type=float, default=0.0)
    ap.add_argument("--avg-zc", type=float, default=None,
                    help="本年度日均占资(亿元)；持仓文件没有资金占用积数表时手动传入")
    ap.add_argument("--trades-dir", default=os.path.join(DESK, '交易记录'),
                    help="历史交易查询目录；默认本地数据目录下的交易记录")
    ap.add_argument("--hosts", default=None,
                    help='当日买入的主办人映射 JSON：{"债券代码": "主办人"}；主办人只能来自用户当天给的文本')
    ap.add_argument("--fifo", action="store_true",
                    help="多笔买入时自动按 FIFO 匹配；默认不自动，先列出候选让用户选择")
    ap.add_argument("--choices", default=None,
                    help='用户选择使用哪笔买入的 JSON：{"债券代码": {"买入结算日":"YYYY-MM-DD","主办人":"X"}}')
    ap.add_argument("--daijiao-choices", default=None,
                    help='债券同时出现在持仓区和代投代缴区时，指定是否算代缴：{"债券代码": true/false}')
    ap.add_argument("--tax-month", type=float, default=None, help="本月截至目前应缴税费(万元)")
    ap.add_argument("--tax-year", type=float, default=None, help="本年累计税费(万元)")
    ap.add_argument("--date", default=None)
    ap.add_argument("--outdir", default=DESK, help="输出目录；可用SWSC_DATA_DIR设置默认值")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    today = parse_date(args.date) or datetime.date.today()
    replay_lots, replay_audits = {}, {}
    if args.replay_costs:
        from pathlib import Path
        from replay_trade_lots import replay
        config_path = Path(args.replay_costs).resolve()
        config = json.loads(config_path.read_text(encoding='utf-8'))
        for code, rule in config.items():
            if rule.get('method') != 'fifo':
                sys.exit('历史重放配置必须明确 method=fifo，并由用户确认')
            cutoff = parse_date(rule['cutoff'])
            if cutoff is None or cutoff >= today:
                sys.exit('历史重放截止日必须早于本次计算日')
            files = [str((config_path.parent / f).resolve()) for f in rule['files']]
            remaining, audit = replay(files, str(code), cutoff.isoformat(),
                                      set(rule['statuses']), rule['expected_quantity'], rule.get('opening', ()))
            replay_lots[str(code)] = remaining
            replay_audits[str(code)] = {'remaining': remaining, 'audit': audit,
                                       'cutoff': cutoff.isoformat(), 'statuses': rule['statuses']}
    # 当日买入的主办人：只能来自用户当天给的文本（代码→主办人），不得用录入员代替
    host_map = {}
    if args.hosts:
        try:
            raw = json.load(open(args.hosts, encoding="utf-8"))
            if isinstance(raw, dict):
                host_map = {str(k).split(".")[0].strip(): str(v).strip() for k, v in raw.items()}
        except Exception as e:
            sys.exit(f"读取主办人映射失败: {e}")
    # 用户对「多笔买入」的选择：{"代码": {"买入结算日": "...", "主办人": "..."}} 或 {"代码": "fifo"}
    choices = {}
    if args.choices:
        try:
            raw = json.load(open(args.choices, encoding="utf-8"))
            if isinstance(raw, dict):
                choices = {str(k).split(".")[0].strip(): v for k, v in raw.items()}
        except Exception as e:
            sys.exit(f"读取选择文件失败: {e}")
    daijiao_choices = {}
    if args.daijiao_choices:
        try:
            raw = json.load(open(args.daijiao_choices, encoding="utf-8"))
            if isinstance(raw, dict):
                daijiao_choices = {str(k).split(".")[0].strip(): bool(v) for k, v in raw.items()}
        except Exception as e:
            sys.exit(f"读取代缴选择文件失败: {e}")
    all_trades = load_trades(args.trade)
    trades = [t for t in all_trades if t['交易状态'] in args.trade_statuses]
    if not trades:
        sys.exit('指定状态范围内没有交易，请核对交易状态列与--trade-statuses')
    sells = [t for t in trades if "卖出" in t["方向"]]
    buys = [t for t in trades if "买入" in t["方向"] and "卖出" not in t["方向"]]

    # 1) 交易量
    trade_amt = sum(t["结算金额"] for t in trades)            # 元
    trade_vol_yi = trade_amt / 1e8                              # 亿元

    # 2) 前一日持仓表：日期 < 今天的最新一份（持仓表_YYYYMMDD.xlsx）
    if not args.prev:
        snaps = []
        for s in glob.glob(os.path.join(DESK, "持仓表[-_]*.xlsx")):
            m = re.search(r"持仓表[-_](\d{4})[-_]?(\d{2})[-_]?(\d{2})", os.path.basename(s))
            if m:
                d = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                if d < today:
                    snaps.append(s)
        args.prev = max(snaps) if snaps else latest_holdings()
    hold_rows, hold_sum = load_holdings(args.prev)
    all_hold = hold_rows["持仓"] + hold_rows["代投代缴"]
    daijiao_codes = {x["代码"].split(".")[0] for x in hold_rows["代投代缴"]}
    hold_codes = {x["代码"].split(".")[0] for x in hold_rows["持仓"]}

    # 当日买入也作为 FIFO 批次（用于「当日买、当日卖」）
    same_day_sig = {}
    for x in all_hold:
        if x["买入结算日"] == today and x["买入数量万"] > 0:
            sig = (x["代码"].split(".")[0], round(x["买入数量万"], 4), round(x["结算金额"], 2))
            same_day_sig[sig] = same_day_sig.get(sig, 0) + 1
    for t in buys:
        sig = (t["代码"], round(t["面额万"], 4), round(t["结算金额"], 2))
        if same_day_sig.get(sig, 0) > 0:      # 持仓表里已有这批当日买入，避免重复
            same_day_sig[sig] -= 1
            continue
        all_hold.append({
            "简称": t["名称"], "代码": t["代码带后缀"], "担保": t["担保"],
            "买入数量万": t["面额万"], "结算金额": t["结算金额"],
            "买入结算日": today, "主办人": host_map.get(t["代码"], ""), "每日票息": 0.0,
            "买入来源": "当日买入", "排序时间": parse_trade_time(t["交易时间"], today),
            "交易编号": t["交易编号"],
        })

    # 每批买入的成本：优先按买入日交易查询的结算金额（ima 口径）
    for x in all_hold:
        x["成本"], x["成本来源"] = resolve_lot_cost(x, args.trades_dir)

    # 3) 结算金额差（按代码汇总卖出）
    sell_by_code = {}
    for t in sells:
        k = t["代码"]
        sell_by_code.setdefault(k, {"名称": t["名称"], "面额万": 0.0, "结算金额": 0.0, "代码带后缀": t["代码带后缀"]})
        sell_by_code[k]["面额万"] += t["面额万"]
        sell_by_code[k]["结算金额"] += t["结算金额"]

    diffs = []          # 每只券的结算金额差
    pending = []        # 多笔买入、需用户先选择
    daijiao_pending = []  # 同时在【持仓】区和【代投代缴】区、需用户判断是否代缴
    for code, s in sell_by_code.items():
        # ---- 代缴归属：只在一个区块 → 按该区块；两个区块都有 → 让用户判断 ----
        in_dj, in_hold = code in daijiao_codes, code in hold_codes
        if in_dj and not in_hold:
            is_daijiao = True
        elif in_hold and not in_dj:
            is_daijiao = False
        elif in_dj and in_hold:                 # 两个区块都有 → 让用户判断
            if code in daijiao_choices:
                is_daijiao = daijiao_choices[code]
            else:
                is_daijiao = None
                daijiao_pending.append({
                    "代码": code, "代码带后缀": s["代码带后缀"], "名称": s["名称"],
                    "面额万": s["面额万"], "卖出结算金额": s["结算金额"],
                    "持仓区有": in_hold, "代投代缴区有": in_dj,
                })
        else:                                    # 两边都没有（如当日买入卖出）→ 非代缴
            is_daijiao = False

        cands = candidate_lots(all_hold, code)
        replay_approved = code in replay_lots
        if replay_approved:
            original = cands[0] if cands else {'主办人': '', '担保': '非担保', '区块': ''}
            cands = [{**original, '简称': s['名称'], '代码': s['代码带后缀'],
                      '买入数量万': float(x['quantity']), '结算金额': float(x['cost']),
                      '成本': float(x['cost']), '成本来源': '经用户确认的历史FIFO重放',
                      '买入结算日': None, '买入来源': '历史重放剩余批次',
                      '交易编号': x['trade_id']} for x in replay_lots[code]]
        choice = choices.get(code)
        auto = replay_approved or args.fifo or choice == "fifo"
        if choice and choice != "fifo":          # 用户指定了用哪笔（可多笔）
            want = choice if isinstance(choice, list) else [choice]
            picked = []
            for w in want:
                for x in cands:
                    bd = x["买入结算日"].strftime("%Y-%m-%d") if x["买入结算日"] else ""
                    if ((not w.get("买入结算日") or w["买入结算日"] == bd)
                            and (not w.get("主办人") or w.get("主办人") == x["主办人"])
                            and (not w.get("交易编号") or w["交易编号"] == x.get("交易编号", ""))):
                        if x not in picked:
                            picked.append(x)
            if picked:
                cands = picked
                auto = True

        cand_view = [{
            "买入来源": x.get("买入来源", ""), "交易编号": x.get("交易编号", ""),
            "买入结算日": x["买入结算日"].strftime("%Y-%m-%d") if x["买入结算日"] else "",
            "主办人": x["主办人"], "担保": x["担保"],
            "买入数量万": x["买入数量万"], "结算金额": round(x["结算金额"], 2),
        } for x in cands]

        if len(cands) > 1 and not auto:
            # 多笔买入：先提醒用户选择，不自动 FIFO
            pending.append({"代码": code, "代码带后缀": s["代码带后缀"], "名称": s["名称"],
                            "面额万": s["面额万"], "卖出结算金额": s["结算金额"],
                            "候选买入": cand_view})
            diffs.append({"代码": code, "代码带后缀": s["代码带后缀"], "名称": s["名称"],
                          "面额万": s["面额万"], "卖出结算金额": s["结算金额"],
                          "买入结算金额": None, "结算金额差": None,
                          "代投代缴": is_daijiao, "未匹配完": False,
                          "需人工确认": True})
            continue

        buy_amt, match_details, short = fifo_buy_amount(cands, s["面额万"])
        diff = round(s["结算金额"] - buy_amt, 2)
        diffs.append({"代码": code, "代码带后缀": s["代码带后缀"], "名称": s["名称"], "面额万": s["面额万"],
                      "卖出结算金额": s["结算金额"], "买入结算金额": buy_amt,
                      "结算金额差": diff, "代投代缴": is_daijiao,
                      "未匹配完": short, "需人工确认": False})

    coupon_audit = []
    if args.coupons:
        from coupon_adjustments import apply_coupons
        coupon_audit = apply_coupons(diffs, args.coupons, today)
    total_diff = sum(d["结算金额差"] for d in diffs if d["结算金额差"] is not None)
    zong = total_diff + args.chai                          # 总结算金额差
    daijiao = sum(d["结算金额差"] for d in diffs
                  if d["代投代缴"] is True and d["结算金额差"] is not None)
    cuohe = total_diff - daijiao                           # 分类表排除拆借成本

    # 4) 当日持仓表：票息收入 / 时点占资 / 日终持仓（当日下午 15:30 抓取生成的当日表）
    if not args.hold:
        args.hold = latest_holdings(today)
    _, hold_sum2 = load_holdings(args.hold) if args.hold and os.path.exists(args.hold) else (None, {"票息收入万元": 0, "自有资金": None, "持仓敞口": None})
    piaoxi = hold_sum2.get("票息收入万元") or 0.0           # 万元（票息总额）
    # 日均占资
    if args.avg_zc is not None:
        avg_zc = args.avg_zc
    else:
        raw = avg_fund_balance(args.hold, today)
        avg_zc = raw / 1e8 if raw is not None else None      # 亿元

    result = {
        "纳入交易状态": args.trade_statuses,
        "排除的交易编号": [t['交易编号'] for t in all_trades if t['交易状态'] not in args.trade_statuses],
        "人工付息核对": coupon_audit,
        "历史重放核对": replay_audits,
        "日期": today.strftime("%Y-%m-%d"),
        "交易量（亿元）": round(trade_vol_yi, 10),
        "总结算金额差（元）": round(zong, 2),
        "当日收益（万元）": round(zong / 10000, 10),
        "票息收入（万元）": round(piaoxi, 10),
        "撮合（万元）": round(cuohe / 10000, 10),
        "代缴（万元）": round(daijiao / 10000, 10),
        "日均占资（亿元）": round(avg_zc, 10) if avg_zc is not None else None,
        "时点占资（亿元）": hold_sum2["自有资金"] / 1e8 if hold_sum2["自有资金"] else None,
        "当日日终持仓（亿元）": hold_sum2["持仓敞口"] / 1e8 if hold_sum2["持仓敞口"] else None,
        "拆借成本（元）": args.chai,
        "需人工确认": pending,
        "有需人工确认项": bool(pending),
        "代缴待确认": daijiao_pending,
        "有代缴待确认项": bool(daijiao_pending),
        "逐笔结算金额差": [
            {**d, "结算金额差": (round(d["结算金额差"], 2) if d["结算金额差"] is not None else None)}
            for d in diffs],
        "前一日持仓表": args.prev,
    }
    print(json.dumps(result, ensure_ascii=False, indent=1))
    if not args.json:
        if pending:
            print_pending(pending)
            print("⚠ 有需人工确认项，总结算金额差暂不完整；你确认后我再重跑。")
        if daijiao_pending:
            print_daijiao_pending(daijiao_pending)
            print("⚠ 有代缴归属待确认项，本次不生成分类表；你确认后我再重跑。")
        if avg_zc is None:
            print("[注意] 持仓文件没有资金占用积数统计表，日均占资留空；请用 --avg-zc 传入（亿元）。")
        missing = sorted({t["名称"] for t in buys if not host_map.get(t["代码"])})
        if missing:
            print("\n[注意] 以下当日买入缺少主办人（主办人只能来自用户当天给的文本，"
                  "请补充后通过 --hosts 传入）：" + "、".join(missing))
        gen_xlsx(result, args, today, skip_class=bool(daijiao_pending))


def print_pending(pending):
    """多笔买入的卖出：先列候选，交由用户选择，不自动 FIFO。"""
    lines = ["", "⚠ 【需要你确认：以下债券卖出时存在多笔买入】"]
    for p in pending:
        lines.append(f"{p['名称']} {p['代码带后缀']}  卖出 {p['面额万']:.0f}万  "
                     f"卖出结算金额 {p['卖出结算金额']:,.2f}")
        for i, c in enumerate(p["候选买入"], 1):
            tag = c["买入来源"] + (f" #{c['交易编号']}" if c["交易编号"] else "")
            lines.append(f"  {i}. [{tag}] {c['买入结算日']} {c['主办人'] or '（主办人未知）'} "
                         f"{c['担保']} 买入 {c['买入数量万']:.0f}万 / 结算 {c['结算金额']:,.2f}")
    lines.append("请告诉我要用哪一笔（或哪几笔），或者要删除/修改哪一行；确认后我再按你的选择计算。")
    print("\n".join(lines))


def find_prev(pattern, before=None):
    """在 本地数据目录 和 本地数据目录/表单 找文件名 MMDD 早于 before 的最新文件；找不到则取最新。"""
    cands = []
    for base in (DESK, os.path.join(DESK, "表单")):
        for f in glob.glob(os.path.join(base, pattern)):
            m = re.match(r"(\d{2})(\d{2})", os.path.basename(f))
            mmdd = int(m.group(1)) * 100 + int(m.group(2)) if m else -1
            cands.append((mmdd, f))
    if not cands:
        return None
    if before is not None:
        b = before.month * 100 + before.day
        prior = [c for c in cands if 0 < c[0] < b]
        if prior:
            return max(prior)[1]
    return max(cands)[1]


def print_daijiao_pending(items):
    """同时在【持仓】区和【代投代缴】区的债券：让用户判断这笔卖出是否算代缴。"""
    lines = ["", "⚠ 【需要你确认：以下债券同时出现在【持仓】区和【代投代缴】区】"]
    for p in items:
        lines.append(f"{p['名称']} {p['代码带后缀']}  卖出 {p['面额万']:.0f}万  "
                     f"卖出结算金额 {p['卖出结算金额']:,.2f}")
        lines.append("  该债券在两个区块都有买入记录，请确认这笔卖出算不算「代缴」。")
    lines.append('确认后用 --daijiao-choices \'{"债券代码": true/false}\' 传入后重跑。')
    print("\n".join(lines))


def gen_xlsx(result, args, today, skip_class=False):
    from template_reports import write_reports
    write_reports(result, args, today, find_prev, parse_date, skip_class)


if __name__ == "__main__":
    main()
