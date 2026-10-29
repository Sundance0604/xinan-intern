#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
西南证券持仓表 —— 买入记录处理脚本

职责：
    接收"已经处理好的字段"（结构化 JSON 输入），
    计算并生成持仓表最终 Excel。

输入（input.json）：
    {
      "settle_date": "2026-10-12",          # 买入结算日（当天）
      "bonds": [
        {
          "code": "032690011.IB",           # 债券代码
          "name": "26无锡城南PPN001",        # 简称
          "quantity_wan": 2000,             # 买入数量（万）
          "net_price": 100.4150,            # 买入净价
          "host": "郑一凡",                  # 主办人
          # ---- 以下由 DM 接口提供（基础信息）----
          "coupon": 1.97,                   # 票面利率 %
          "remaining_tenor": "4.89Y",       # 剩余年限
          "rating": "AA+",                  # 主体评级
          "guarantee": "非担保",             # 担保/非担保
          "remaining_size_wan": 36000,      # 剩余规模（万元）
          "inte_start_date": "2026-09-02",  # 起息日（算应计利息用）
          "inte_calcu_cls": "A/A",          # 计息基准
          # ---- 以下由 DM 接口提供（估值，T-1 或结算日）----
          "valuation_net_price": 100.6481,  # 中债估值净价
          "valuation_yield": 1.8299,        # 中债估值收益率 %（百分比）
          # ---- 买入收益率（可选，不填则由净价反解）----
          "buy_yield": 1.5,                 # 用户直接给的买入收益率（百分比，如 1.5）
          "next_pay_date": "2027-09-02",    # 下次付息日（反解 YTM 用）
          "matu_pay_date": "2030-09-02",    # 到期兑付日（反解 YTM 用）
          "option_date": "",                # 行权日（回售/永续债，反解"行权"口径）
          "pay_freq": 1                     # 每年付息次数（1=按年，2=半年）
        }
      ]
    }

输出：
    持仓表格式的 Excel（列序与「销售交易部持仓表.xlsx」一致 A~S 业务列）
    不用管的列（舆情、内评、卖出结算日等）留空。

口径约定（对齐「输出模板.xlsx」，用户 2026-10-27 明确）：
    表头在第 4 行（前 3 行留白），共 30 列 A~AD（A~S 业务列 + T~AD 汇总列）。
    「率」类字段（票面 D、买入收益率 E、中债估值 P）存**小数**（0.042 = 4.2%），
    配合百分比格式显示（D 用 0.00%，E/P 用 0.0000%）。

公式（已验证）：
    面值 = 数量(万) * 10000 元
    应计利息(每百元) = 票面% * 天数 / 365 * 100   （计息基准 A/365F）
    结算金额 J = 全价 * 面值 / 100
    资本利得 Q = (估值净价 - 买入净价) * 面值 / 100
    每日票息 R = 票面% * 面值 / 365
    买入收益率 E = 用户给的值，或由买入净价反解 YTM（二分法）
        - 有 buy_yield 直接用；无则用 next_pay_date + matu_pay_date 反解
        - 回售/永续债给 option_date，按"行权"口径反解
"""

import json
import sys
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


def parse_date(s):
    return datetime.strptime(str(s).strip(), "%Y-%m-%d")


def round_half_up(value, ndigits):
    """四舍五入（非银行家舍入），避免浮点边界导致的尾数偏差。"""
    if value is None:
        return None
    return float(Decimal(str(value)).quantize(Decimal('1.' + '0' * ndigits), rounding=ROUND_HALF_UP))


def calc_accrued_interest(coupon_pct, inte_start, settle_date, inte_calcu_cls="A/365F"):
    """
    计算每百元面值的应计利息。
    计息基准：
      A/365F  实际天数 / 365
      A/A     实际天数 / 实际天数（年按365近似）
    简化：统一用 实际天数 / 365，误差 <0.05%（此前已验证）。
    """
    days = (settle_date - inte_start).days
    if days < 0:
        days = 0
    return coupon_pct * days / 365.0  # 每百元面值应计利息（元）


def _cashflow_times(settle_date, next_pay_date, matu_date, pay_freq=1, option_date=None):
    """
    生成现金流的时间点（年化），用于反解买入收益率。
    - pay_freq: 每年付息次数（1=按年，2=半年）
    - option_date: 若给定（回售/行权日），现金流只到该日（含回售价），对应"行权"口径
    返回：[(时间点_年, 现金流_元/百元), ...]，最后一个是本金+票息
    """
    period_days = 365.0 / pay_freq
    end = parse_date(option_date) if option_date else parse_date(matu_date)

    times = []
    d = parse_date(next_pay_date)
    while d <= end:
        t = (d - settle_date).days / 365.0
        if t > 0:
            times.append(t)
        d = d + timedelta(days=int(round(period_days)))
    if not times:
        return []
    return times


def _bond_dirty_price(yield_rate, coupon_pct, times, pay_freq=1):
    """
    给定年化收益率 yield_rate（小数），计算全价（元/百元）。
    口径（与交易系统一致，已用 23锦生01 行权 1.317% 锚定验证）：
      - 最后一笔现金流期限 <= 1 年：单利 full = CF / (1 + y*t)
      - 最后一笔现金流期限 > 1 年：复利 full = CF / (1 + y/pay_freq)^(t*pay_freq)
    （国内债券定价标准约定：剩余期限<=1年单利，>1年复利）
    """
    last_t = times[-1]
    if last_t <= 1.0:
        # 单利：每笔现金流按 (1 + y*t) 折现
        price = 0.0
        for t in times:
            price += coupon_pct / (1 + yield_rate * t)
        price += 100.0 / (1 + yield_rate * last_t)
        return price
    # 复利
    coupon_per_period = coupon_pct / pay_freq
    period_rate = yield_rate / pay_freq
    price = 0.0
    for t in times:
        price += coupon_per_period / (1 + period_rate) ** (t * pay_freq)
    price += 100.0 / (1 + period_rate) ** (times[-1] * pay_freq)
    return price


def calc_buy_yield(full_price, coupon_pct, settle_date, next_pay_date, matu_date,
                   pay_freq=1, option_date=None):
    """
    由买入全价反解买入收益率（YTM）。
    二分法求解，区间 [-0.5, 0.5]（-50% ~ 50%），迭代 200 次，精度 <1e-8。
    返回百分比数值（如 2.5 = 2.5%）；无法求解时返回 None。
    """
    times = _cashflow_times(settle_date, next_pay_date, matu_date, pay_freq, option_date)
    if not times:
        return None

    lo, hi = -0.5, 0.5
    f_lo = _bond_dirty_price(lo, coupon_pct, times, pay_freq) - full_price
    f_hi = _bond_dirty_price(hi, coupon_pct, times, pay_freq) - full_price
    if f_lo * f_hi > 0:
        # 全价不在收益率区间内（异常票息/价格），返回 None
        return None

    for _ in range(200):
        mid = (lo + hi) / 2.0
        f_mid = _bond_dirty_price(mid, coupon_pct, times, pay_freq) - full_price
        if f_mid == 0.0:
            return mid * 100.0
        if f_lo * f_mid < 0:
            hi = mid
        else:
            lo = mid
            f_lo = f_mid
    return (lo + hi) / 2.0 * 100.0


def process_bond(b, settle_date):
    """处理单只债券，返回持仓表一行（A~S 列）。"""
    code = b["code"]
    name = b.get("name", "")
    qty_wan = float(b["quantity_wan"])           # 万
    _net = b.get("net_price")
    net_price = float(_net) if (_net is not None and _net != "") else None   # 买入净价（元/百元面值），缺失则 None
    host = b.get("host", "")

    coupon = float(b.get("coupon", 0))           # 票面 %
    remaining_tenor = b.get("remaining_tenor", "")
    rating = b.get("rating", "")
    guarantee = b.get("guarantee", "")
    remaining_size_wan = b.get("remaining_size_wan")   # 万元
    inte_start_date = b.get("inte_start_date", "")
    inte_calcu_cls = b.get("inte_calcu_cls", "A/365F")
    _vnp = b.get("valuation_net_price")
    valuation_net_price = float(_vnp) if (_vnp is not None and _vnp != "") else None
    _vy = b.get("valuation_yield")
    valuation_yield = float(_vy) if (_vy is not None and _vy != "") else None
    # DM 提供的日间应计利息（每百元面值），优先使用（更准确）
    day_accrued_interest = b.get("day_accrued_interest")

    # ---- 买入收益率相关字段（反解 YTM 用，均可选）----
    buy_yield = b.get("buy_yield")          # 用户直接给的买入收益率（小数或百分数，见下）
    next_pay_date = b.get("next_pay_date", "")      # 下次付息日（DM: 可推）
    matu_pay_date = b.get("matu_pay_date", "")      # 到期兑付日（DM: matu_pay_date）
    option_date = b.get("option_date", "")          # 行权日（DM: next_option_date，回售/永续用）
    pay_freq = int(b.get("pay_freq", 1))            # 每年付息次数（1=按年，2=半年）

    # 面值（元）
    face_value = qty_wan * 10000.0

    # 应计利息（每百元）—— 口径：按计息基准区分含头/不含头
    #   A/365F（交易所 SH/SZ）：天数 = (结算日 - 上次付息日) + 1（含头）
    #       例：09-23付息 → 10-15结算 = 23天
    #   A/A（银行间 IB）：天数 = (结算日 - 上次付息日)（不含头）
    #       例：08-25付息 → 10-15结算 = 51天
    #   两者差一天，根因是计息基准的起算约定不同（已用同事真实数据验证）
    last_pay_date = b.get("last_pay_date", "")       # 上次付息日（DM: year_pay_date 推算）
    if last_pay_date:
        days = (settle_date - parse_date(last_pay_date)).days
        if inte_calcu_cls and inte_calcu_cls.startswith("A/365"):
            days += 1   # 交易所 A/365F 含头
        # 银行间 A/A 不含头，days 不变
        accrued_100 = coupon * days / 365.0
    elif day_accrued_interest is not None:
        # 兜底：DM 估值接口的日间应计利息
        accrued_100 = float(day_accrued_interest)
    elif inte_start_date:
        # 再兜底：从起息日粗算（仅适用于当年新发债，跨年付息会有误差）
        accrued_100 = calc_accrued_interest(coupon, parse_date(inte_start_date), settle_date, inte_calcu_cls)
    else:
        accrued_100 = 0.0

    # 全价 = 净价 + 应计利息（净价缺失时，全价/结算金额/资本利得/收益率一并留空）
    if net_price is not None:
        full_price = net_price + accrued_100
        settle_amount = full_price * face_value / 100.0
        # 资本利得：估值净价缺失（如一级缴款未上市）时留空
        capital_gain = (valuation_net_price - net_price) * face_value / 100.0 if valuation_net_price is not None else None
    else:
        full_price = None
        settle_amount = None
        capital_gain = None

    # 买入收益率 E（统一百分比口径，如 2.4621 = 2.4621%）：
    #   1) 用户直接给了 buy_yield -> 用它（百分比，如 1.5）
    #   2) 没给 -> 用买入净价反解 YTM（返回百分比）
    buy_yield_pct = None
    if buy_yield is not None and float(buy_yield) != 0:
        buy_yield_pct = float(buy_yield)   # 百分比，直接使用
    elif net_price is not None:
        # 反解：需要 next_pay_date + matu_pay_date（或 option_date）
        if next_pay_date and (matu_pay_date or option_date):
            buy_yield_pct = calc_buy_yield(
                full_price, coupon, settle_date, next_pay_date,
                matu_pay_date, pay_freq, option_date or None
            )

    # 买入收益率 E 统一保留 4 位小数（百分比），避免满精度导致每日票息尾数不干净
    if buy_yield_pct is not None:
        buy_yield_pct = round_half_up(float(buy_yield_pct), 4)

    # 每日票息 R = 买入收益率(小数，4位) × 数量(万) × 10000 / 365
    # 注意：原文件公式是 E*I/365*10000（E=买入收益率，I=买入数量万），
    # 用的是「买入收益率」，不是「票面利率」。两者含义不同。
    # 用四舍五入后的 E（4位）计算，保证 E 列与 R 列自洽。
    daily_coupon = None
    if buy_yield_pct is not None:
        daily_coupon = buy_yield_pct / 100.0 * face_value / 365.0

    # 剩余规模（亿） = 万元 / 10000
    remaining_size_yi = None
    if remaining_size_wan is not None:
        remaining_size_yi = float(remaining_size_wan) / 10000.0

    # 持仓工作日（参考）：结算日到今天的交易日数（这里简化为自然日占位，正式由交易日历算）
    # 脚本只对"已处理字段"服务，持仓工作日由外部交易日历提供时再填入，此处留空
    holding_days = ""

    # 组装 A~S 列
    row = {
        "A": name,                          # 简称
        "B": code,                          # 代码
        "C": remaining_tenor,               # 剩余年限
        "D": coupon,                        # 票面（百分比，如 4.2 = 4.2%）
        "E": buy_yield_pct,                 # 买入收益率（百分比；未给则用净价反解）
        "F": rating,                        # 主体/债项评级
        "G": "",                            # 舆情（留空）
        "H": guarantee,                     # 担保/非担保
        "I": qty_wan,                       # 买入数量（万）
        "J": round(settle_amount, 2) if settle_amount is not None else "",       # 结算金额（元）
        "K": settle_date,                   # 买入结算日
        "L": holding_days,                  # 持仓工作日（留空，待交易日历）
        "M": "",                            # 卖出结算日（留空）
        "N": host,                          # 主办人
        "O": "",                            # 内评（留空）
        "P": valuation_yield,               # 当日中债估值（%）（百分比，如 2.3188）
        "Q": round(capital_gain, 2) if capital_gain is not None else "",        # 资本利得
        "R": round(daily_coupon, 4) if daily_coupon is not None else "",        # 每日票息
        "S": remaining_size_yi,             # 剩余规模（亿）
    }
    return row


def _set_rate(ws, coord, pct_value, border, fmt):
    """率类字段：百分比数值 -> 小数（/100），配合百分比格式显示。"""
    cell = ws[coord]
    if pct_value is None or pct_value == "":
        cell.value = None
    else:
        cell.value = float(pct_value) / 100.0
        cell.number_format = fmt
    cell.border = border


def _set_num(ws, coord, value, border, fmt):
    """数值字段：写值 + 数字格式。"""
    cell = ws[coord]
    if value is None or value == "":
        cell.value = None
    else:
        cell.value = value
        cell.number_format = fmt
    cell.border = border


def _write_row(ws, r_idx, row, border):
    """写一行数据，处理各列数字格式与公式（对齐「输出模板.xlsx」）。"""
    # 文本/空列
    for col in ["A", "B", "C", "F", "G", "H", "M", "N", "O"]:
        val = row.get(col, "")
        cell = ws[f"{col}{r_idx}"]
        cell.value = val if val not in ("", None) else None
        cell.border = border

    # 率类（票面 D、买入收益率 E、中债估值 P）：存小数
    _set_rate(ws, f"D{r_idx}", row.get("D"), border, "0.00%")
    _set_rate(ws, f"E{r_idx}", row.get("E"), border, "0.0000%")
    _set_rate(ws, f"P{r_idx}", row.get("P"), border, "0.0000%")

    # 数值列
    _set_num(ws, f"I{r_idx}", row.get("I"), border, "0")           # 买入数量（万，整数）
    _set_num(ws, f"J{r_idx}", row.get("J"), border, "#,##0.00")    # 结算金额
    _set_num(ws, f"Q{r_idx}", row.get("Q"), border, "#,##0.00")    # 资本利得
    _set_num(ws, f"S{r_idx}", row.get("S"), border, "0.00")        # 剩余规模（亿）

    # 买入结算日 K：日期 + 中文日期格式
    kc = ws[f"K{r_idx}"]
    kc.value = row.get("K")
    kc.number_format = 'm"月"d"日"'
    kc.border = border

    # 持仓工作日 L：公式（交易日数，排除周末；节假日区域 A137:A142 预留）
    lc = ws[f"L{r_idx}"]
    lc.value = f'=NETWORKDAYS(K{r_idx},TODAY(),$A$137:$A$142)'
    lc.border = border

    # 每日票息 R：公式 =E*I/365*10000（E 为小数，仅当 E 有值时写）
    rc = ws[f"R{r_idx}"]
    if row.get("E") is not None:
        rc.value = f'=E{r_idx}*I{r_idx}/365*10000'
        rc.number_format = "0.0000"
    rc.border = border

    # T~AD 汇总列：留空（用户后续填/加公式），仅加边框
    for i in range(20, 31):  # 20=T ... 30=AD
        ws[f"{get_column_letter(i)}{r_idx}"].border = border


def build_workbook(rows, settle_date, out_path):
    """生成持仓表 Excel（格式对齐「输出模板.xlsx」）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "买入记录"

    # 30 列表头（A~AD）
    headers = [
        "简称", "代码", "剩余年限", "票面", "买入收益率", "主体/债项评级", "舆情",
        "担保/非担保", "买入数量（万）", "结算金额（元）", "买入结算日", "持仓工作日（参考）",
        "卖出结算日", "主办人", "内评", "当日中债估值（%）", "资本利得", "每日票息", "剩余规模",
        "具体舆情", "部门本周资金成本", "中低评级合计\n（合计不超15亿）", "AA持仓数量",
        "AA结算金额\n（事业部不超5亿）", "AA+持仓数量", "AA+结算金额", "票息总额",
        "每日融资收益率", "加权平均收益率", "持仓收益率",
    ]

    # 表头样式
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 表头写在第 4 行（前 3 行留白，对齐模板）
    HEADER_ROW = 4
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=HEADER_ROW, column=i, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center
        c.border = border

    # 列宽（对齐模板 A~S；T~AD 用默认 14）
    widths = [
        42.71, 42.71, 32.50, 29.71, 20.71, 25.71, 35.20, 16.18, 24.28, 28.32,
        34.26, 56.51, 32.50, 24.28, 20.23, 16.18, 33.71, 25.71, 20.71, 20.71,
        14, 14, 14, 14, 14, 14, 14, 14, 14, 14,
    ]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # 数据行（从第 5 行开始）
    for r_idx, row in enumerate(rows, start=HEADER_ROW + 1):
        _write_row(ws, r_idx, row, border)

    wb.save(out_path)
    return out_path


def main():
    if len(sys.argv) < 2:
        print("用法: python bond_process.py input.json [output.xlsx]")
        sys.exit(1)

    in_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else "持仓表_买入记录.xlsx"

    with open(in_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    settle_date = parse_date(data.get("settle_date", "2026-10-12"))
    bonds = data.get("bonds", [])

    rows = [process_bond(b, settle_date) for b in bonds]

    build_workbook(rows, settle_date, out_path)
    print(f"已生成 {len(rows)} 条买入记录 -> {out_path}")


if __name__ == "__main__":
    main()
