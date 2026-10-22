#!/usr/bin/env python3
"""读取交易查询导出表（186列，openpyxl常读不了其样式），输出规范化的交易清单 JSON。

用法: python load_trades.py 交易查询.xlsx
处理:
  - openpyxl/pandas 读取失败时，自动用 LibreOffice 转 csv 再读
  - 同业拆借行、末尾合计行直接忽略
  - 标注: 担保/非担保、规范化代码(加 .SH/.SZ/.IB 后缀)、同日同量买卖对冲、一级代投代缴(交易品种=分销买卖/方向=分销买入/备注含一级缴款)、交易状态
"""
import sys, json, re, subprocess, tempfile, os, glob
import pandas as pd

IGNORE_PRODUCTS = {"同业拆借"}
SUFFIX = {"上交所": ".SH", "深交所": ".SZ", "银行间": ".IB", "中债登": ".IB", "清算所": ".IB"}


def read_any(path):
    try:
        return pd.read_excel(path, dtype=str)
    except Exception:
        d = tempfile.mkdtemp()
        subprocess.run(["soffice", "--headless", "--convert-to", "csv", "--outdir", d, path],
                       check=True, capture_output=True)
        return pd.read_csv(glob.glob(os.path.join(d, "*.csv"))[0], dtype=str)


def num(s):
    try:
        return float(str(s).replace(",", ""))
    except Exception:
        return None


def load(path):
    df = read_any(path)
    out, ignored = [], {"同业拆借": 0, "合计行/空行": 0}
    for _, r in df.iterrows():
        prod = r.get("交易品种")
        if not isinstance(r.get("交易编号"), str) or not re.fullmatch(r"\d{5,}", r["交易编号"].strip()) or pd.isna(prod):
            ignored["合计行/空行"] += 1; continue
        if prod in IGNORE_PRODUCTS:
            ignored["同业拆借"] += 1; continue
        settle = str(r.get("到期结算方式") or "")
        code = str(r.get("代码") or "").strip()
        mkt = str(r.get("托管市场") or "").strip()
        issuer = str(r.get("发行人/参考实体") or "")
        out.append({
            "交易编号": r["交易编号"].strip(), "状态": r.get("交易状态"), "品种": prod, "业务品种": r.get("业务品种"),
            "日期": r.get("交易日期"), "方向": r.get("交易方向"), "代码": code,
            "规范代码": code + SUFFIX.get(mkt, ""), "名称": r.get("名称"),
            "数量(万)": num(r.get("交易面额(万)")), "结算金额": num(r.get("结算金额(元)")),
            "担保": "担保" if ("担保交收" in settle and "非担保" not in settle) else "非担保",
            "到期结算方式": settle, "到期收益率": r.get("到期收益率(%)"), "行权收益率": r.get("行权收益率(%)"),
            "发行人": issuer,
            "一级代投代缴": (prod == "分销买卖" or "分销" in str(r.get("交易方向") or "")
                             or "一级" in str(r.get("备注") or "")),
            "备注": r.get("备注") if isinstance(r.get("备注"), str) else "",
        })
    # 同日同代码、买卖数量相同 => 对冲（晚间表预期无此行）
    for t in out:
        t["对冲"] = False
    for b in out:
        if b["方向"] == "买入" and not b["一级代投代缴"] and not b["对冲"]:
            for s in out:
                if (s["方向"] == "卖出" and not s["对冲"] and s["代码"] == b["代码"]
                        and s["数量(万)"] == b["数量(万)"] and s["日期"] == b["日期"]):
                    b["对冲"] = s["对冲"] = True; break
    return {"交易": out, "忽略": ignored}


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    print(json.dumps(load(sys.argv[1]), ensure_ascii=False, indent=1))
