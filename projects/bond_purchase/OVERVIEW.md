# xinan-intern — 西南证券固收交易岗自动化工具

本项目面向**西南证券固定收益（债券）交易岗**的日常持仓管理流程，通过接入 **DM（大智慧）金融数据 MCP 接口**，将债券持仓表中"可自动化的市场数据字段"（票面利率、剩余规模、评级、中债估值、资本利得等）实现批量拉取与自动计算，替代人工在 iFind / 万得中逐条查询、刷卷、手工录入的重复劳动。

> 仓库发布合成测试数据；真实交易数据保留在本地 local_data/，不提交。

---

## 目录结构

```text
.
├── bond_process.py
├── README.md
├── OVERVIEW.md
├── requirements.txt
├── test_data/
│   ├── input_sample.json
│   └── （合成 JSON 示例）
└── examples/
    └── sample_purchase.xlsx
```

---

## 项目在做什么

### 业务背景

固收交易岗每天需要维护一张「持仓表」，记录每一只持仓债券的：

- **市场数据**：票面利率、剩余年限、主体/债项评级、是否城投、剩余规模、当日中债估值
- **交易数据**：买入数量、买入净价、结算金额、买入收益率、主办人、买入结算日
- **计算结果**：资本利得、每日票息、持仓工作日

其中「市场数据」可以从 DM 接口直接获取，「计算结果」可以由公式算出，只有「交易数据」（买了多少、什么价、谁买的）必须来自人工/内部系统。

### 本项目解决的痛点

| 环节 | 传统做法 | 本项目做法 |
|------|---------|-----------|
| 中债估值更新 | 打开 iFind 刷卷模板，逐条复制 | DM 接口批量拉取，自动选「行权/到期」收益率 |
| 票面/规模/评级 | 万得逐条查 | DM `bond_basic_info` 批量返回 |
| 结算金额 | 手工计算 | 脚本按公式自动算 |
| 资本利得 | 手工计算 | 脚本按公式自动算 |
| 偏离度提醒 | 手工算 | DM 实时行情直接返回买卖收益率-中债/中证 |

---

## 快速开始

### 1. 环境准备

- Python 3.x（需安装 `openpyxl`）
- 可选：conda 环境

```bash
pip install openpyxl
```

### 2. 配置 DM MCP

本项目依赖 DM（大智慧数据服务）的 MCP 接口获取债券市场数据。配置方式：在 WorkBuddy 的 `~/.workbuddy/mcp.json` 中写入以下 HTTP 服务（AppKey / AppSecret 需向数据服务方申请）：

```json
{
  "mcpServers": {
    "dm-mcp-service-bond-analysis": {
      "type": "http",
      "url": "https://gapi-ext.innodealing.com/dm-mcp-service/publicapi/mcp/v1/bond-analysis/",
      "headers": { "x-app-key": "<你的AppKey>", "x-app-secret": "<你的AppSecret>" }
    },
    "dm-mcp-service-bond-market": {
      "type": "http",
      "url": "https://gapi-ext.innodealing.com/dm-mcp-service/publicapi/mcp/v1/bond-market/",
      "headers": { "x-app-key": "<你的AppKey>", "x-app-secret": "<你的AppSecret>" }
    }
  }
}
```

共 15 个服务（bond-analysis / bond-primary / bond-market / offshore-bond / company / fund / futures / fx / derivatives / money-market / bill / macro / stock / index / tools），按需配置。配置后需在 WorkBuddy 连接器管理页「信任」这些服务才会生效。

### 3. 运行脚本

```bash
python bond_process.py test_data/input_sample.json [输出.xlsx]
```

脚本只对「已经处理好的字段」服务：输入结构化 JSON（含债券代码、数量、净价、主办人 + DM 拉取的市场数据），输出持仓表格式的 Excel。

---

## 核心脚本使用说明

详见 [`README.md`](README.md)。

### 输入格式（input.json）

```json
{
  "settle_date": "2026-10-12",
  "bonds": [
    {
      "code": "032690011.IB",           // 债券代码
      "name": "26无锡城南PPN001",        // 简称
      "quantity_wan": 2000,             // 买入数量（万）
      "net_price": 100.4150,            // 买入净价
      "host": "郑一凡",                  // 主办人
      "coupon": 1.97,                   // 票面利率 %（DM）
      "remaining_tenor": "4.89Y",       // 剩余年限（DM）
      "rating": "AA+",                  // 主体评级（DM）
      "guarantee": "非担保",             // 担保/非担保（DM）
      "remaining_size_wan": 36000,      // 剩余规模万元（DM）
      "valuation_net_price": 100.6481,  // 中债估值净价（DM）
      "valuation_yield": 1.8299,        // 中债估值收益率 %（DM）
      "day_accrued_interest": 0.2159    // 日间应计利息（DM，推荐）
    }
  ]
}
```

### 计算公式（已通过真实数据验证）

| 字段 | 公式 |
|------|------|
| 面值 | 数量(万) × 10000 元 |
| 结算金额 J | (买入净价 + 应计利息) × 面值 / 100 |
| 资本利得 Q | (中债估值净价 − 买入净价) × 面值 / 100 |
| 每日票息 R | 票面% × 面值 / 365 |

---

## 数据来源映射（DM 接口 → 持仓表字段）

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
| 交易日历 | tools | get_trade_dates（含调休） |

---

## 关键规则与注意事项

1. **估值类型选择**：含回售权/永续的债券，估值要取 `valuation_type=行权` 那条（与持仓表 P 列口径一致），否则与「到期」收益率会差 0.3%+。

2. **应计利息必须用 DM 的 `day_valuation_interest`**（日间应计利息，已含「上次付息日→结算日」的正确天数）。⚠️ 不要从起息日自己算——跨年付息的老债会算错（如某 2024 年发行债券，从起息日算会多 7 元/百元，导致结算金额错 56.8 万元）。

3. **估值数据是 T-1**：调用 `bond_valuation` 时需显式传日期，当天（T 日）查询返回空，需查上一交易日。

4. **代码查询偶尔返回空**：用债券简称（`sec_short_name_list`）查询比代码（`security_id_list`）更稳定。

5. **留空的字段**：舆情、内评、卖出结算日等，脚本统一留空（依赖内部系统/人工）。

6. **持仓工作日**：脚本暂留空，需用 DM `get_trade_dates` 交易日历计算（注意调休日仅银行间开市）。

---

## 完整工作流（每日操作）

1. 拿到当天的买入清单（微信/内部系统导出的未处理字段）
2. 整理成结构化输入（代码、数量、净价、主办人）
3. 调用 DM 接口批量拉取市场数据（票面、估值、应计利息等）
4. 运行 `bond_process.py` 生成最终 Excel
5. 核对结果，粘贴回持仓表

---

## 免责声明

本项目为工作流程自动化工具，所有计算结果仅供参考，正式记账请以内部系统为准。DM 接口的可用性、数据权限以数据服务方为准。
