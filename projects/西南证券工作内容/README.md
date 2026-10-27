# 销售交易日常工作

本模块串联买入记录、腾讯文档持仓快照、持仓对账、三张报表和持仓天数提醒。Python使用本机`pavane`环境；腾讯文档导出另需Node.js和Chrome。

**新同事从[交接使用说明](交接使用说明.md)开始**：包含安装命令、输入清单、每天的操作步骤、担保录入、实收付息、税费、审核和常见问题。三张空白模板位于`templates/`，随代码一起交接。

## 目录

| 位置 | 用途 | 是否推送 |
|---|---|---|
| `scripts/*.py`、`scripts/*.js` | 可复用计算与读取代码 | 是 |
| `scripts/bond_purchase/` | 原工作流配套的买入记录计算器 | 是 |
| `tests/` | 合成数据测试，不含真实交易 | 是 |
| `templates/` | 用户确认的三张空白报表模板，不含真实交易数据 | 是 |
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

### 当日担保买入的处理顺序

当日担保买入先算好结算金额与资本利得，保存到本地审计文件；持仓表J/Q当日留空并加“次日按原值补录、待王老师审核”的备注，重算受影响汇总后再生成第三张盈利表。次日按已保存的当日值补录，不能改用次日价格。按交易交收方式识别担保，不要只相信持仓标签；当日已有部分卖出的，只排除剩余持仓部分。

不要将J留空视为买入成本0，也不要取消真实的当日卖出收益。票息和资金积数表继续按各自口径取数。持仓快照缓存的汇总不会因本地清空单元格自动改变，须实际重算后再取数。完整步骤及审核要求见[工作流程注意事项](成本核验注意事项.md)。

### 固定模板和输入

生成代码默认套用[固定模板](templates/README.md)，保留用户统一的字体、列宽、边框、合并区域及盈利口径说明。模板提供样式，`--prev-profit`和`--prev-classification`提供真实历史数据，不能互相替代。盈利表日期统一为A1报告日、B1年末、C1年初，年化收益率引用`A1-C1`。

将示例日期和路径替换为本次实际输入；拆借成本以负数传入。

```powershell
conda run -n pavane python scripts/pnl_tables.py local_data/交易查询.xlsx --date YYYY-MM-DD --prev local_data/前日持仓.xlsx --hold local_data/当日持仓.xlsx --trades-dir local_data/交易记录 --prev-profit local_data/前日盈利表.xlsx --prev-classification local_data/前日分类表.xlsx --chai -100 --outdir local_data/output
```

先加`--json`检查匹配，再生成Excel。多批买入需经用户确认，可通过`--choices`和`--daijiao-choices`指定；主办人只认用户提供的交易文本。历史持仓金额与买入金额相同，也仍需检查后续卖出是否已消耗该批次。

当日交易默认仅纳入“成交确认”；其他状态经本次业务确认后，用`--trade-statuses 成交确认 交易执行中 审批通过`显式选择。确认的实收付息可用`--coupons 本地付息配置.json`补计，配置格式及防重复规则见交接手册。付息不是拆借成本，也不能与持仓估算票息重复加到收益。脚本填表后需用Excel/WPS重算、保存，确保次日能读取公式缓存。

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

盈亏表同名同代码债券的多笔卖出按结算日汇总为一行，结算金额差求和；逐笔成本和批次消耗记录保留在本地`_support/`。沿用用户最新确认的格式，汇总后更新求和范围及合并区域。详见[同名债券合并列示规则](成本核验注意事项.md)。

## 测试

```powershell
conda run -n pavane python -m unittest discover -s tests -v
```

完整的共享核验规则见[成本核验注意事项](成本核验注意事项.md)。真实案例留在本地归档。
