# FIFO 卖出盈利统计

Python 直接读取原始 Excel 的 OOXML，兼容 WPS 非标准样式。文件名和工作表名可任意，按列名识别数据。输出包括逐债券结算差、逐笔 FIFO 明细、JSON 异常日志和 CSV 汇总。

## 安装与 GUI

```powershell
python -m pip install -r requirements.txt
python fifo_sell_profit.py
```

窗口中选择交易查询文件、持仓记录文件，再点击“生成汇总和日志”。日期留空时从唯一卖出交易日识别；存在多个卖出日期时必须指定日期。Python 标准桌面安装自带 tkinter；没有图形环境时使用终端模式。

## 终端模式

```powershell
python fifo_sell_profit.py "任意交易文件名.xlsx" "任意持仓文件名.xlsx"
python fifo_sell_profit.py --trades "交易查询.xlsx" --holdings "持仓.xlsx" --date 2026-09-24
python fifo_sell_profit.py --trades "交易查询.xlsx" --holdings "持仓.xlsx" --confirmed-only
```

支持 `--gui` 打开窗口，`--holding-only` 仅使用持仓成本，`--output-dir` 指定其他输出目录。默认输出在脚本所在目录，与 Python 脚本同目录：

```text
YYYY-MM-DD_FIFO盈利统计.xlsx
YYYY-MM-DD_FIFO日志.json
YYYY-MM-DD_FIFO盈利汇总.csv
```

同一天重复运行会更新结果，不修改源 Excel。退出码 0 表示成功，2 表示有卖出成本缺口（报告和日志仍会生成），1 表示读取或运行失败，并写入 `运行错误.json`。

## 统计口径

- **卖出结算差**：默认纳入指定交易日“成交确认”和“交易执行中”的买入、卖出。先用持仓记录，再用卖出前的当日买入按 FIFO 匹配。买入记录按买入结算日、原行号排序；当日交易按交易时间、原行号排序。持仓数量视为卖出前可用余额，代投代缴记录也纳入。
- **成本分摊**：结算金额按匹配面额比例分摊。使用 Decimal，累计成本四舍五入到分，确保整批卖完后的分摊成本等于该批原始结算金额。
- **跨批次日志**：单笔卖出超过首条可用买入剩余数量时记录异常及拆分明细。买入总量不足时盈利为空，不将缺失成本当零。同一批次被多笔卖出消耗也记录在 JSON。
- **所有交易面额合计（亿元）**：读取交易查询的“交易面额(元)”，所有交易行相加后除以 100,000,000。包括买入、卖出、分销、同业拆借及其他日期，独立于盈利的日期/状态筛选；表尾合计行只核作源数据，不重复计入。
- **票息总额（亿元）**：读取持仓记录“票息总额”的唯一有效值，按用户指定的亿元口径原值展示，不额外换算、不叠加到卖出结算差。若该值是公式，工作簿必须保存计算缓存；缺失或多个值会报错。
- **当日债券买卖结算金额合计（亿元）**：保留原指标，仅计算盈利口径内的债券买卖结算金额，与“所有交易面额合计”不同。

JSON 保留源文件哈希、工作表、真实行号、各买入批次成本和剩余数量。债券代码按文本保留前导零。

## 测试数据与验证

`test_data/sample_trades.xlsx` 和 `test_data/sample_holdings.xlsx` 为合成数据，不包含真实机构和人员信息。覆盖多笔卖出、跨批次、执行中买入、不同交易日期、同业拆借、合计行排除和票息读取。

```powershell
python test_data/generate_samples.py
python fifo_sell_profit.py test_data/sample_trades.xlsx test_data/sample_holdings.xlsx
python -m unittest discover -s tests -v
```

示例结果：卖出结算差 55,000.00 元，所有交易面额 0.041 亿元，票息总额 0.0000025 亿元，1 次跨批次事件。真实回归数据仅存在本地 `local_data/`，不存在时相应测试跳过，合成测试仍完整运行。
