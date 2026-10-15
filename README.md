# 西南证券债券交易自动化工具

仓库包含两个独立项目，代码、测试数据和说明分别存放。

> 👉 **使用 Agent 的同事请看**：[《使用 Agent 的协作规范：默认日期、数据处理原则与方法》](AGENT_GUIDE.md) —— 统一日期约定、字段口径和标准流程，避免算错。

| 项目 | 入口 | 用途 |
|---|---|---|
| [债券买入记录处理](projects/bond_purchase/README.md) | `projects/bond_purchase/bond_process.py` | 读取结构化 JSON，生成持仓表买入记录 |
| [FIFO 卖出盈利统计](projects/bond_profit/README.md) | `projects/bond_profit/fifo_sell_profit.py` | GUI 或终端读取两份 Excel，统计卖出结算差、所有交易面额和票息总额 |

```text
projects/
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

两个项目的真实输入均保留在各自本地 `local_data/` 目录，公开仓库仅提供合成测试数据。市场数据准备、字段映射和原工作流见买入项目 [项目概览](projects/bond_purchase/OVERVIEW.md)。
