import json
import subprocess
import sys
import tempfile
import unittest
from collections import deque
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook

SCRIPTS = Path(__file__).resolve().parents[1] / 'scripts'
sys.path.insert(0, str(SCRIPTS))
from replay_trade_lots import new_lot, consume, replay


def workbook(path, rows):
    w = Workbook()
    for row in rows:
        w.active.append(row)
    w.save(path)


HEAD = ['交易编号', '交易状态', '交易品种', '交易日期', '首期结算日期',
        '交易时间', '交易方向', '代码', '名称', '交易面额(万)', '结算金额(元)', '托管市场']


def trade(i, direction, amount, time, status='成交确认', day='2026-01-01'):
    return [i, status, '现券买卖', day, day, f'{day} {time}', direction,
            '012345', '合成债券', 10, amount, '上交所']


class HistoryTests(unittest.TestCase):
    def test_partial_cost_conservation_and_shortage(self):
        lots = deque([new_lot('synthetic', 2, '1.01')])
        first, _ = consume(lots, 1)
        second, _ = consume(lots, 1)
        self.assertEqual((first, second), (Decimal('.51'), Decimal('.50')))
        self.assertFalse(lots)
        lots = deque([new_lot('synthetic', 1, 1)])
        with self.assertRaises(ValueError):
            consume(lots, 2)
        self.assertEqual(lots[0]['used_quantity'], 0)

    def test_consumed_history_is_not_available_and_files_are_deduplicated(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / 'trades.xlsx'
            workbook(p, [HEAD, trade('1', '买入', 900, '09:00:00'),
                         trade('2', '卖出', 950, '10:00:00'),
                         trade('3', '买入', 1100, '11:00:00')])
            left, _ = replay([p, p], '012345', '2026-01-01', {'成交确认'}, 10)
            self.assertEqual(left, [{'trade_id': '3', 'quantity': '10', 'cost': '1100.00'}])

    def test_statuses_are_explicit_and_conflicting_ids_fail(self):
        with tempfile.TemporaryDirectory() as d:
            a, b = Path(d) / 'a.xlsx', Path(d) / 'b.xlsx'
            workbook(a, [HEAD, trade('1', '买入', 1000, '09:00:00', '交易执行中')])
            left, _ = replay([a], '012345', '2026-01-01', {'成交确认'}, 0)
            self.assertEqual(left, [])
            left, _ = replay([a], '012345', '2026-01-01', {'成交确认', '交易执行中'}, 10)
            self.assertEqual(left[0]['cost'], '1000.00')
            workbook(b, [HEAD, trade('1', '买入', 1001, '09:00:00', '交易执行中')])
            with self.assertRaises(ValueError):
                replay([a, b], '012345', '2026-01-01', {'交易执行中'}, 10)

    def test_cli_replay_overrides_stale_cost_and_classification_excludes_financing(self):
        with tempfile.TemporaryDirectory() as d:
            base = Path(d)
            history, today, holdings = [base / x for x in ('history.xlsx', 'today.xlsx', 'holdings.xlsx')]
            workbook(history, [HEAD, trade('1', '买入', 900, '09:00:00'),
                               trade('2', '卖出', 950, '10:00:00'),
                               trade('3', '买入', 1100, '11:00:00')])
            workbook(today, [HEAD, trade('4', '卖出', 1200, '09:00:00', day='2026-01-02')])
            workbook(holdings, [
                ['简称', '代码', '担保/非担保', '买入数量（万）', '结算金额（元）',
                 '买入结算日', '主办人', '买入收益率', '每日票息', '票息总额'],
                ['【持仓】', None, None, None, None, None, None, None, None, 1],
                ['合成债券', '012345.SH', '非担保', 10, 1000, '2026-01-01', '合成人员', .02, 1],
            ])
            config = base / 'replay.json'
            config.write_text(json.dumps({'012345': {'method': 'fifo', 'files': ['history.xlsx'],
                'cutoff': '2026-01-01', 'statuses': ['成交确认'], 'expected_quantity': 10}}), encoding='utf-8')
            cmd = [sys.executable, str(SCRIPTS / 'pnl_tables.py'), str(today),
                   '--date', '2026-01-02', '--prev', str(holdings), '--hold', str(holdings),
                   '--trades-dir', str(base), '--avg-zc', '1', '--chai', '-100',
                   '--replay-costs', str(config), '--json']
            result = subprocess.run(cmd, check=True, capture_output=True, encoding='utf-8')
            data = json.loads(result.stdout)
            self.assertEqual(data['总结算金额差（元）'], 0)
            self.assertEqual(data['撮合（万元）'], .01)
            self.assertFalse(data['有需人工确认项'])
            self.assertEqual(data['逐笔结算金额差'][0]['买入结算金额'], 1100)
            coupons = base / 'coupons.json'
            coupons.write_text(json.dumps({'date': '2026-01-02', 'source': '合成付息审核',
                'payments': {'012345': 2}}), encoding='utf-8')
            paid = json.loads(subprocess.run(cmd + ['--coupons', str(coupons)], check=True,
                              capture_output=True, encoding='utf-8').stdout)
            self.assertEqual(paid['总结算金额差（元）'], 2)
            self.assertEqual(paid['撮合（万元）'], .0102)
            executing_buy = trade('5', '买入', 1000, '10:00:00', '交易执行中', day='2026-01-02')
            executing_buy[7] = '999999'
            workbook(today, [HEAD, trade('4', '卖出', 1200, '09:00:00', day='2026-01-02'), executing_buy])
            selected = json.loads(subprocess.run(cmd, check=True, capture_output=True, encoding='utf-8').stdout)
            self.assertEqual(selected['排除的交易编号'], ['5'])
            included = json.loads(subprocess.run(cmd + ['--trade-statuses', '成交确认', '交易执行中'],
                                  check=True, capture_output=True, encoding='utf-8').stdout)
            self.assertEqual(included['排除的交易编号'], [])
            self.assertAlmostEqual(included['交易量（亿元）'] - selected['交易量（亿元）'], .00001)


if __name__ == '__main__':
    unittest.main()
