#!/usr/bin/env python3
"""完整核对：早盘持仓表 + 晚间持仓表 + 交易查询表 -> 逐笔核对结果（JSON）。

用法: python reconcile.py 早盘.xlsx 晚间.xlsx 交易查询.xlsx
只核对行的新增/删除和所在区块，不核对数值。
"""
import sys, json
from collections import defaultdict
from diff_holdings import load as load_hold, diff
from load_trades import load as load_trades


def base(c):
    return str(c).split(".")[0]


def q(x):
    return float(str(x).replace(",", ""))


def run(morning, evening, trade_file):
    old, new = load_hold(morning), load_hold(evening)
    d = diff(old, new)
    trades = load_trades(trade_file)["交易"]
    old_codes = {base(x["key"][0]) for s in old for x in old[s]}

    added = {s: list(d[s]["新增"]) for s in d}
    removed = {s: list(d[s]["删除"]) for s in d}
    used_add, used_rem = set(), set()
    results = []

    def res(trade, expect, actual, verdict):
        results.append({"交易": trade, "期望": expect, "实际": actual, "结论": verdict})

    def find_add(code, qty, date=None, sections=("持仓", "代投代缴")):
        for s in sections:
            for i, a in enumerate(added[s]):
                if (s, i) in used_add:
                    continue
                if base(a["代码"]) == code and q(a["买入数量(万)"]) == qty and (date is None or a["买入结算日"] == date):
                    used_add.add((s, i)); return s, a
        return None, None

    # ---- 买入类 ----
    for t in trades:
        if t["方向"] not in ("买入", "分销买入"):
            continue
        name = f'{t["名称"]} {t["数量(万)"]:.0f}万 {t["方向"]}'
        code, qty = t["代码"], t["数量(万)"]
        if t["对冲"]:
            s, a = find_add(code, qty, t["日期"])
            res(name, "同日买卖对冲，晚间表不应有该行", f'{s}区出现新增行(第{a["晚间表行号"]}行)' if a else "无新增行",
                "❌ 多余" if a else "✅ 正确")
        elif t["一级代投代缴"]:
            s, a = find_add(code, qty)
            if a and s == "代投代缴":
                res(name, "新增在代投代缴区", f'代投代缴区新增(第{a["晚间表行号"]}行)', "✅ 正确")
            elif a:
                res(name, "新增在代投代缴区", f'错放在{s}区(第{a["晚间表行号"]}行)', "❌ 位置错误")
            else:
                res(name, "新增在代投代缴区", "未新增", "❌ 遗漏")
        else:
            kind = ("加仓" if code in old_codes else "二级买入") + ("(担保买入)" if t["担保"] == "担保" else "")
            s, a = find_add(code, qty, t["日期"])
            if a and s == "持仓":
                note = "；担保买入占位，次日补全" if t["担保"] == "担保" else ""
                note += "；原行未改" if code in old_codes else ""
                res(name + f" [{kind}]", "持仓区新增一行", f'持仓区新增(第{a["晚间表行号"]}行){note}', "✅ 正确")
            elif a:
                res(name + f" [{kind}]", "持仓区新增一行", f'错放在{s}区(第{a["晚间表行号"]}行)', "❌ 位置错误")
            else:
                res(name + f" [{kind}]", "持仓区新增一行", "未新增", "❌ 遗漏")

    # ---- 卖出类：按 代码+担保 汇总 ----
    sells = defaultdict(float)
    for t in trades:
        if t["方向"] == "卖出" and not t["对冲"]:
            sells[(t["代码"], t["担保"], t["名称"])] += t["数量(万)"]
    sell_keys = {(c, g) for (c, g, _) in sells}
    for (code, guar, nm), total in sells.items():
        rem = [(s, i, r) for s in removed for i, r in enumerate(removed[s])
               if base(r["代码"]) == code and r["担保"] == guar]
        rem_qty = sum(q(r["买入数量(万)"]) for _, _, r in rem)
        for s, i, _ in rem:
            used_rem.add((s, i))
        where = "、".join(f'{s}区第{r["早盘表行号"]}行({r["买入数量(万)"]}万)' for s, _, r in rem) or "无"
        name = f"{nm} 合计{total:.0f}万 卖出" + ("(担保卖出)" if guar == "担保" else "")
        if guar == "担保":
            if rem:
                res(name, "当天保留原行，不删", f"已删除：{where}", "❌ 提前删除")
            else:
                res(name, "当天保留原行，不删", "原行仍在", "✅ 正确(次日需删除)")
        else:
            if abs(rem_qty - total) < 1e-6:
                res(name, f"删除合计{total:.0f}万", f"已删除：{where}", "✅ 正确")
            elif rem_qty < total:
                res(name, f"删除合计{total:.0f}万", f"仅删除{rem_qty:.0f}万：{where}", "❌ 遗漏(少删)")
            else:
                res(name, f"删除合计{total:.0f}万", f"删除了{rem_qty:.0f}万：{where}", "❌ 多余(多删)")
    # 担保卖出但删除了同代码的担保行，已在上面处理；对"卖出代码+担保"之外的删除视为多余

    extra = []
    for s in added:
        for i, a in enumerate(added[s]):
            if (s, i) not in used_add:
                extra.append(f'{s}区新增 第{a["晚间表行号"]}行 {a["简称"]} {a["买入数量(万)"]}万')
    for s in removed:
        for i, r in enumerate(removed[s]):
            if (s, i) not in used_rem:
                extra.append(f'{s}区删除 早盘表第{r["早盘表行号"]}行 {r["简称"]} {r["买入数量(万)"]}万')
    return {"逐笔": results, "无对应交易的变动": extra,
            "统计": {"总数": len(results), "正确": sum(r["结论"].startswith("✅") for r in results),
                   "有问题": sum(r["结论"].startswith("❌") for r in results)}}


if __name__ == "__main__":
    if len(sys.argv) < 4:
        sys.exit(__doc__)
    print(json.dumps(run(*sys.argv[1:4]), ensure_ascii=False, indent=1))
