# 债券买入记录处理脚本使用说明

## 脚本职责

`bond_process.py` 只对**已经处理好的字段**提供服务：接收结构化的 JSON 输入（含债券代码、数量、净价、主办人，以及 DM 接口拉取的市场数据），计算结算金额、资本利得、每日票息，生成持仓表格式的 Excel。

**脚本本身不调用 DM 接口**——DM 数据由对话侧拉取后整理进 JSON。

## 使用方式

```bash
python bond_process.py input.json [输出.xlsx]
```

本项目的代码、测试输入和示例结果已集中到 `projects/bond_purchase/`。
从仓库根目录运行示例：

```powershell
python projects/bond_purchase/bond_process.py projects/bond_purchase/test_data/input_sample.json projects/bond_purchase/examples/sample_purchase.xlsx
```

完整项目背景见 [OVERVIEW.md](OVERVIEW.md)。

`test_data/input_sample.json` 和 `examples/sample_purchase.xlsx` 为合成示例。原始真实输入和工作簿迁移到本地 `local_data/` 并排除提交。

## 输入格式（input.json）

```json
{
  "settle_date": "2026-10-12",
  "bonds": [
    {
      "code": "032690011.IB",           // 债券代码（必填）
      "name": "26无锡城南PPN001",        // 简称
      "quantity_wan": 2000,             // 买入数量（万）（必填）
      "net_price": 100.4150,            // 买入净价（必填）
      "host": "郑一凡",                  // 主办人（必填）
      "coupon": 1.97,                   // 票面利率 %（DM: iss_coup_rate）
      "remaining_tenor": "4.89Y",       // 剩余年限（DM: remaining_tenor）
      "rating": "AA+",                  // 主体评级（DM: issuer_rating）
      "guarantee": "非担保",             // 担保/非担保（DM: guarantor_name 是否为空）
      "remaining_size_wan": 36000,      // 剩余规模万元（DM: new_size）
      "inte_start_date": "2026-09-02",  // 起息日（兜底算应计利息用）
      "inte_calcu_cls": "A/A",          // 计息基准（DM: inte_calcu_cls，决定应计利息含头/不含头）
      "last_pay_date": "2026-09-02",    // 上次付息日（算应计利息用，DM: year_pay_date 推算）
      "valuation_net_price": 100.6481,  // 中债估值净价（DM: valuation_net_price）
      "valuation_yield": 1.8299,        // 中债估值收益率 %（百分比）
      "day_accrued_interest": 0.2159,   // 日间应计利息 元/百元（DM: day_valuation_interest，推荐）
      // ---- 买入收益率（可选，不填则由净价反解）----
      "buy_yield": 1.5,                 // 用户直接给的买入收益率（百分比，如 1.5 = 1.5%）
      "next_pay_date": "2027-09-02",    // 下次付息日（反解 YTM 用）
      "matu_pay_date": "2030-09-02",    // 到期兑付日（反解 YTM 用）
      "option_date": "",                // 行权日（回售/永续债，反解"行权"口径）
      "pay_freq": 1                     // 每年付息次数（1=按年，2=半年）
    }
  ]
}
```

> **口径约定**：「率」类字段（票面 D、买入收益率 E、中债估值 P）统一用**百分比数值**，即 `4.2` 表示 4.2%（不再用 `0.042` 小数）。

## 计算公式（已验证）

| 字段 | 公式 |
|---|---|
| 面值 | 数量(万) × 10000 元 |
| 结算金额 J | (买入净价 + 应计利息) × 面值 / 100 |
| 资本利得 Q | (中债估值净价 − 买入净价) × 面值 / 100 |
| 每日票息 R | 买入收益率(小数) × 面值 / 365（= 买入收益率% × 数量万 × 10000 / 365，对应原表 `=E*I/365*10000`） |
| 买入收益率 E | 用户给的 `buy_yield`，或由买入净价反解 YTM（二分法） |

### 买入收益率的计算规则

- **给了 `buy_yield`** → 直接用（百分比，如 `1.5`）。
- **没给 `buy_yield`** → 用买入全价（净价 + 应计利息）反解到期收益率（YTM）：
  - 需要 `next_pay_date` + `matu_pay_date`（普通债）；回售/永续债再给 `option_date`，按「行权」口径反解。
  - 缺这些字段则留空。
  - 反解方法：二分法迭代（区间 -50%~50%，精度 <1e-8），票息频率按 `pay_freq` 折算。

> 口径与「估值取行权」一致：有回售/永续的债券，买入收益率也应取「行权」口径。

## 关键注意事项

1. **应计利息按计息基准区分「含头/不含头」**（关键，搞混差一天）：
   - **A/365F（交易所 SH/SZ）**：天数 = (结算日 − 上次付息日) **+ 1**（含头）。如 09-23 付息 → 10-15 结算 = 23 天。
   - **A/A（银行间 IB）**：天数 = (结算日 − 上次付息日)（不含头）。如 08-25 付息 → 10-15 结算 = 51 天。
   - 输入 JSON 提供 `last_pay_date` + `inte_calcu_cls` 即可，脚本自动判断。
   - ⚠️ 真实案例：银行间债统一用含头会多算一天，结算金额差 643~1561 元。
   - ⚠️ DM 的 `day_valuation_interest` 是 T-1 口径，仅作兜底。
   - ⚠️ 不要从起息日自己算——跨年付息的债券会算错（如 24大横琴MTN005 起息 2024 年，若从起息日算会多 7 元/百元）。

2. **估值类型选择**：有回售权/永续的债券，估值净价和收益率要取 `valuation_type=行权` 那条（与表内 P 列口径一致）。

3. **留空的列**：舆情(G)、买入收益率(E)、持仓工作日(L)、卖出结算日(M)、内评(O) 等，脚本统一留空。

4. **持仓工作日**：脚本暂留空，正式需要 DM `get_trade_dates` 交易日历计算（结算日→今天的交易日数）。

## 数据来源映射（DM 接口 → 字段）

| 持仓表字段 | DM 接口 | DM 字段 |
|---|---|---|
| 票面 D | bond_basic_info | iss_coup_rate |
| 评级 F | bond_basic_info | issuer_rating |
| 担保 H | bond_basic_info | guarantor_name（空=非担保） |
| 剩余年限 C | bond_basic_info | remaining_tenor |
| 剩余规模 S | bond_basic_info | new_size（万元→亿） |
| 中债估值 P | bond_valuation | valuation_yield（行权优先） |
| 应计利息 | bond_valuation | day_valuation_interest |
| 估值净价 | bond_valuation | valuation_net_price |
