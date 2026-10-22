# 西南证券债券交易自动化工具

仓库包含三个独立项目，代码、测试数据和说明分别存放。

> 👉 **使用 Agent 的同事请看**：[《使用 Agent 的协作规范：默认日期、数据处理原则与方法》](AGENT_GUIDE.md) —— 统一日期约定、字段口径和标准流程，避免算错。

| 项目 | 入口 | 用途 |
|---|---|---|
| [债券买入记录处理](projects/bond_purchase/README.md) | `projects/bond_purchase/bond_process.py` | 读取结构化 JSON，生成持仓表买入记录 |
| [FIFO 卖出盈利统计](projects/bond_profit/README.md) | `projects/bond_profit/fifo_sell_profit.py` | GUI 或终端读取两份 Excel，统计卖出结算差、所有交易面额和票息总额 |
| [城市信用债曲线与批量查询](projects/城市信用债曲线_完整离线包/README.md) | `projects/城市信用债曲线_完整离线包/启动批量查询.cmd` | 导入 HTML 更新估值快照，使用 DeepSeek 解析批量 OFR 文本，查看城市曲线并导出利差报告 |

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

买入与卖出项目的真实输入均保留在各自本地 `local_data/` 目录，公开仓库仅提供合成交易测试数据。市场数据准备、字段映射和原工作流见买入项目 [项目概览](projects/bond_purchase/OVERVIEW.md)。

城市信用债项目使用 Node.js 18+，随包提供 2026-09-15 市场数据快照及原始备份；可以导入新的兼容 HTML，并在本地持久保存。个人 DeepSeek 密钥与 `runtime/` 运行快照不提交。首次使用请参考该项目 README 和无密钥配置示例。
