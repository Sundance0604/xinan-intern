"""生成合成测试数据，金额与名称均为虚构：python test_data/generate_samples.py"""
from datetime import date
from pathlib import Path
from openpyxl import Workbook


def generate(directory=None):
    directory = Path(directory or Path(__file__).resolve().parent)
    directory.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = '交易查询'
    ws.append(['交易编号', '交易状态', '交易日期', '交易方向', '代码', '名称',
               '交易面额(万)', '交易面额(元)', '结算金额(元)', '交易时间'])
    trades = [('T01', '成交确认', '2026-09-24', '买入', '000001', '示例债券甲', 100, 1000000, '09:00:00'),
              ('T02', '成交确认', '2026-09-24', '卖出', '000001', '示例债券甲', 60, 610000, '10:00:00'),
              ('T03', '交易执行中', '2026-09-24', '卖出', '000001', '示例债券甲', 40, 410000, '11:00:00'),
              ('T04', '交易执行中', '2026-09-24', '卖出', '000002', '示例债券乙', 30, 330000, '12:00:00'),
              ('T05', '交易执行中', '2026-09-24', '买入', '000003', '示例债券丙', 100, 990000, '09:30:00'),
              ('T06', '成交确认', '2026-09-24', '卖出', '000003', '示例债券丙', 50, 500000, '13:00:00'),
              ('T07', '成交确认', '2026-09-17', '拆入', 'ORD_TEST_1', '示例同业拆借', 10, 100000, '09:00:00'),
              ('T08', '成交确认', '2026-09-24', '拆入', 'ORD_TEST_2', '示例同业拆借', 20, 200000, '09:00:00')]
    for key, status, day, direction, code, name, qty, amount, time in trades:
        ws.append([key, status, day, direction, code, name, qty, qty*10000, amount, f'{day} {time}'])
    ws.append(['8', '', '', '', '', '', 410, 4100000, 4140000, ''])  # 表尾合计，不参与再汇总。
    wb.save(directory/'sample_trades.xlsx')
    wb = Workbook()
    ws = wb.active
    ws.title = '持仓记录'
    ws.append(['简称', '代码', '买入数量（万）', '结算金额（元）', '买入结算日', '票息总额'])
    ws.append(['【持仓】', '', '', '', '', .0000025])
    ws.append(['示例债券乙', '000002.IB', 20, 200000, date(2026, 9, 1), ''])
    ws.append(['示例债券乙', '000002.IB', 30, 300000, date(2026, 9, 2), ''])
    wb.save(directory/'sample_holdings.xlsx')


if __name__ == '__main__':
    generate()
