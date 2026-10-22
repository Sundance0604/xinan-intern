# 西南证券债券交易自动化工具

仓库按业务模块组织，代码、合成测试数据和真实本地数据分开存放。

> 👉 **使用 Agent 的同事请看**：[《使用 Agent 的协作规范：默认日期、数据处理原则与方法》](AGENT_GUIDE.md) —— 统一日期约定、字段口径和标准流程，避免算错。

| 项目 | 入口 | 用途 |
|---|---|---|
| [债券买入记录处理](projects/bond_purchase/README.md) | `projects/bond_purchase/bond_process.py` | 读取结构化 JSON，生成持仓表买入记录 |
| [FIFO 卖出盈利统计](projects/bond_profit/README.md) | `projects/bond_profit/fifo_sell_profit.py` | GUI 或终端读取两份 Excel，统计卖出结算差、所有交易面额和票息总额 |
| [销售交易日常工作](projects/西南证券工作内容/README.md) | `projects/西南证券工作内容/scripts/pnl_tables.py` | 腾讯文档持仓、对账、三张报表及提醒；历史成本需核验剩余批次 |
| [城市信用债曲线与批量查询](projects/城市信用债曲线_完整离线包/README.md) | `projects/城市信用债曲线_完整离线包/启动批量查询.cmd` | 导入 HTML 更新估值快照，使用 DeepSeek 解析批量 OFR 文本，查看城市曲线并导出利差报告 |

```text
projects/
├── 西南证券工作内容/
│   ├── README.md                  # 五个板块与运行命令
│   ├── 成本核验注意事项.md         # 成本、拆借分类与两位小数约定
│   ├── scripts/                   # 通用程序；data/保留本地不推送
│   ├── tests/                     # 合成数据测试
│   └── config.example.json        # 真实配置使用忽略的config.local.json
├── bond_purchase/
│   ├── bond_process.py
│   ├── README.md
│   ├── OVERVIEW.md
│   ├── requirements.txt
│   ├── test_data/
│   └── examples/
└── bond_profit/
    ├── fifo_sell_profit.py
    ├── README.md
    ├── requirements.txt
    ├── tests/
    └── test_data/                  # 合成测试工作簿
```

Python 3.10+；安装两个项目共同依赖：

```powershell
python -m pip install -r projects/bond_profit/requirements.txt
```

FIFO GUI（需要带 tkinter 的桌面 Python）：

```powershell
python projects/bond_profit/fifo_sell_profit.py
```

FIFO 终端示例（文件名不影响识别；默认也纳入交易执行中）：

```powershell
python projects/bond_profit/fifo_sell_profit.py projects/bond_profit/test_data/sample_trades.xlsx projects/bond_profit/test_data/sample_holdings.xlsx
```

汇总 Excel、JSON 日志及 CSV 默认输出在 `fifo_sell_profit.py` 所在目录。新项目的真实输入放在本地 `projects/bond_profit/local_data/`，运行结果也保留本地，不提交到公开仓库。

运行测试：

```powershell
python -m unittest discover -s projects/bond_profit/tests -v
```

买入与卖出项目的真实输入均保留在各自本地 `local_data/` 目录，公开仓库仅提供合成交易测试数据。市场数据准备、字段映射和原工作流见买入项目 [项目概览](projects/bond_purchase/OVERVIEW.md)。

销售交易日常模块兼容现有`projects/西南证券工作内容/scripts/data/日期/output/`路径。真实交易、持仓、内部手册、浏览器配置及生成报表均留在本地；Git只收录代码、配置模板、通用说明和合成测试。报表统一显示两位小数，证券代码、日期和底层计算精度保留。

城市信用债项目使用 Node.js 18+，随包提供 2026-09-15 市场数据快照及原始备份；可以导入新的兼容 HTML，并在本地持久保存。个人 DeepSeek 密钥与 `runtime/` 运行快照不提交。首次使用请参考该项目 README 和无密钥配置示例。
