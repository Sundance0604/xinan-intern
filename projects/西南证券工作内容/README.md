# 销售交易日常工作

本模块串联买入记录、腾讯文档持仓快照、持仓对账、三张报表和持仓天数提醒。Python使用本机`pavane`环境；腾讯文档导出另需Node.js和Chrome。

## 目录

| 位置 | 用途 | 是否推送 |
|---|---|---|
| `scripts/*.py`、`scripts/*.js` | 可复用计算与读取代码 | 是 |
| `scripts/bond_purchase/` | 原工作流配套的买入记录计算器 | 是 |
| `tests/` | 合成数据测试，不含真实交易 | 是 |
| `成本核验注意事项.md` | 历史成本、同业拆借、报表精度约定 | 是 |
| `config.example.json` | 腾讯文档配置模板 | 是 |
| `config.local.json` | 真实文档链接与工作表ID | 否 |
| `scripts/data/` | 现有日期数据与历史交易，保留原路径 | 否 |
| `local_data/` | 新数据的默认位置、内部案例归档 | 否 |
| 日期目录的`output/` | 三张报表；`_support/`保存日志、备份和预览 | 否 |
| 本地`SKILL.md`、`功能清单.md` | 含内部文档链接及联系人信息的原手册 | 否 |

## 准备环境

在本目录执行：

```powershell
conda run -n pavane python -m pip install -r requirements.txt
npm --prefix scripts ci
Copy-Item config.example.json config.local.json
```

编辑`config.local.json`填写有权访问的文档链接、两个工作表ID和Chrome路径。Node.js 24已在本地验证；使用其他版本时需满足`puppeteer-core`的要求。配置和业务数据已加入Git忽略规则。

## 五个工作板块

| 工作 | 入口 | 输入 → 输出 |
|---|---|---|
| 买入记录 | `scripts/bond_purchase/bond_process.py` | 已准备的交易及市场数据JSON → 买入Excel；脚本不查询DM |
| 持仓快照 | `extract_qq_docs.js`、`json_to_xlsx.py` | 腾讯文档 → JSON → 两个工作表的Excel |
| 持仓核对 | `daily_trade_records.py` | 前日和当日持仓 → 买卖变化TXT |
| 三张表 | `pnl_tables.py` | 交易、持仓、历史成本、前日报表 → 结算/分类/盈利Excel |
| 持仓提醒 | `holding_days_reminder.py` | 当日持仓 → 分主办人的持仓天数TXT |

`run_daily.bat`只运行抓取、导出、提醒。三张表需要输入交易和确认成本，不会由该入口自动生成；当前不新增定时任务。

## 抓取与导出

```powershell
node scripts/extract_qq_docs.js --config config.local.json --out local_data/snapshot.json
conda run -n pavane python scripts/json_to_xlsx.py local_data/snapshot.json local_data/持仓表.xlsx
```

导出依赖网页内部数据模型；页面结构变化或登录权限变化时可能失败。空工作表会报错。成功后核对两个工作表的行数、票息、持仓敞口及资金余额；公式未取得计算值时不要直接用于盈利计算。

## 三张表

将示例日期和路径替换为本次实际输入；拆借成本以负数传入。

```powershell
conda run -n pavane python scripts/pnl_tables.py local_data/交易查询.xlsx --date YYYY-MM-DD --prev local_data/前日持仓.xlsx --hold local_data/当日持仓.xlsx --trades-dir local_data/交易记录 --prev-profit local_data/前日盈利表.xlsx --prev-classification local_data/前日分类表.xlsx --chai -100 --outdir local_data/output
```

先加`--json`检查匹配，再生成Excel。多批买入需经用户确认，可通过`--choices`和`--daijiao-choices`指定；主办人只认用户提供的交易文本。历史持仓金额与买入金额相同，也仍需检查后续卖出是否已消耗该批次。

只有用户批准FIFO和交易状态范围后，才使用`--replay-costs 本地配置.json`。配置按代码组织，文件路径相对于配置文件：

```json
{
  "123456": {
    "method": "fifo",
    "files": ["历史交易.xlsx"],
    "cutoff": "2026-01-05",
    "statuses": ["成交确认"],
    "expected_quantity": "1000",
    "opening": []
  }
}
```

`opening`为空仅适用于已确认覆盖完整买入期间的记录；有期初库存时传入`trade_id`、`quantity`、`cost`。缺口、重复编号冲突和剩余数量不符会报错。配置不代表以后日期自动获准同一FIFO或状态范围。

**只有分类表排除同业拆借。** 核对关系为：`(代缴+撮合)×10000+拆借成本=结算表总额`。盈利表引用含拆借成本的当日收益。税费未提供时沿用前日并提示。缺前日累计或日均占资时不填造0。

报表全部显示两位小数，底层计算精度保留。证券代码保持文本，收益率显示两位百分数，不能通过提前舍入原始数据达到显示要求。

## 测试

```powershell
conda run -n pavane python -m unittest discover -s tests -v
```

完整的共享核验规则见[成本核验注意事项](成本核验注意事项.md)。真实案例留在本地归档。
