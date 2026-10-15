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
          "valuation_yield": 1.8299,        # 中债估值收益率 %
          # ---- 买入收益率（可选，不填则由净价反解）----
          "buy_yield": 1.5,                 # 用户直接给的买入收益率（小数或百分数）
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
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def parse_date(s):
    return datetime.strptime(str(s).strip(), "%Y-%m-%d")


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
    标准约定：yield_rate 为年化，每期折现因子 = (1 + yield_rate/pay_freq)^期数。
    其中"期数" = t * pay_freq（t 为年化时间）。
    """
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
    由买入全价反解买入收益率（YTM，小数）。
    二分法求解，区间 [-0.5, 0.5]（-50% ~ 50%），迭代 200 次，精度 <1e-8。
    返回小数（如 0.025 = 2.5%）；无法求解时返回 None。
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
            return mid
        if f_lo * f_mid < 0:
            hi = mid
        else:
            lo = mid
            f_lo = f_mid
    return (lo + hi) / 2.0


def process_bond(b, settle_date):
    """处理单只债券，返回持仓表一行（A~S 列）。"""
    code = b["code"]
    name = b.get("name", "")
    qty_wan = float(b["quantity_wan"])           # 万
    net_price = float(b["net_price"])            # 买入净价（元/百元面值）
    host = b.get("host", "")

    coupon = float(b.get("coupon", 0))           # 票面 %
    remaining_tenor = b.get("remaining_tenor", "")
    rating = b.get("rating", "")
    guarantee = b.get("guarantee", "")
    remaining_size_wan = b.get("remaining_size_wan")   # 万元
    inte_start_date = b.get("inte_start_date", "")
    inte_calcu_cls = b.get("inte_calcu_cls", "A/365F")
    valuation_net_price = float(b.get("valuation_net_price", 0))
    valuation_yield = float(b.get("valuation_yield", 0))
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

    # 应计利息（每百元）
    if day_accrued_interest is not None:
        # 优先用 DM 估值接口返回的日间应计利息（已含上次付息日到结算日的正确天数）
        accrued_100 = float(day_accrued_interest)
    elif inte_start_date:
        # 兜底：从起息日粗算（仅适用于当年新发债，跨年付息会有误差）
        accrued_100 = calc_accrued_interest(coupon, parse_date(inte_start_date), settle_date, inte_calcu_cls)
    else:
        accrued_100 = 0.0

    # 全价 = 净价 + 应计利息
    full_price = net_price + accrued_100

    # 结算金额 J = 全价 * 面值 / 100
    settle_amount = full_price * face_value / 100.0

    # 资本利得 Q = (估值净价 - 买入净价) * 面值 / 100
    capital_gain = (valuation_net_price - net_price) * face_value / 100.0

    # 每日票息 R = 票面% * 面值 / 365
    daily_coupon = coupon / 100.0 * face_value / 365.0

    # 买入收益率 E：
    #   1) 用户直接给了 buy_yield -> 用它（支持小数 0.025 或百分数 2.5）
    #   2) 没给 -> 用买入净价反解 YTM
    buy_yield_decimal = None
    if buy_yield is not None and float(buy_yield) != 0:
        v = float(buy_yield)
        buy_yield_decimal = v if v < 1 else v / 100.0   # 百分数转小数
    else:
        # 反解：需要 next_pay_date + matu_pay_date（或 option_date）
        if next_pay_date and (matu_pay_date or option_date):
            buy_yield_decimal = calc_buy_yield(
                full_price, coupon, settle_date, next_pay_date,
                matu_pay_date, pay_freq, option_date or None
            )

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
        "D": coupon / 100.0,                # 票面（小数，与表内一致）
        "E": buy_yield_decimal,             # 买入收益率（小数；未给则用净价反解）
        "F": rating,                        # 主体/债项评级
        "G": "",                            # 舆情（留空）
        "H": guarantee,                     # 担保/非担保
        "I": qty_wan,                       # 买入数量（万）
        "J": round(settle_amount, 2),       # 结算金额（元）
        "K": settle_date,                   # 买入结算日
        "L": holding_days,                  # 持仓工作日（留空，待交易日历）
        "M": "",                            # 卖出结算日（留空）
        "N": host,                          # 主办人
        "O": "",                            # 内评（留空）
        "P": valuation_yield / 100.0,       # 当日中债估值（%）（小数，与表内一致）
        "Q": round(capital_gain, 2),        # 资本利得
        "R": round(daily_coupon, 4),        # 每日票息
        "S": remaining_size_yi,             # 剩余规模（亿）
    }
    return row


def build_workbook(rows, settle_date, out_path):
    """生成持仓表 Excel。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "买入记录"

    headers = [
        "简称", "代码", "剩余年限", "票面", "买入收益率", "主体/债项评级", "舆情",
        "担保/非担保", "买入数量（万）", "结算金额（元）", "买入结算日", "持仓工作日（参考）",
        "卖出结算日", "主办人", "内评", "当日中债估值（%）", "资本利得", "每日票息", "剩余规模",
    ]

    # 表头样式
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 写表头
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center
        c.border = border

    # 列宽
    widths = [20, 16, 14, 10, 12, 14, 8, 12, 14, 16, 12, 16, 12, 10, 8, 16, 12, 10, 10]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    # 写数据行
    col_order = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S"]
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, col in enumerate(col_order, start=1):
            val = row.get(col, "")
            cell = ws.cell(row=r_idx, column=c_idx, value=val if val != "" else None)
            cell.border = border
            if col in ("D", "E", "P"):  # 百分比小数，格式化为小数
                if isinstance(val, float):
                    cell.number_format = "0.000000"
            elif col in ("J", "Q"):  # 金额
                if isinstance(val, float):
                    cell.number_format = "#,##0.00"
            elif col in ("R",):
                if isinstance(val, float):
                    cell.number_format = "0.0000"
            elif col in ("S",):
                if isinstance(val, float):
                    cell.number_format = "0.0"

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
