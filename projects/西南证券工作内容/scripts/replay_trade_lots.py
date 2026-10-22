"""Replay explicitly selected historical trades; do not infer FIFO authorization.

Callers must supply a complete period or opening lots, and explicitly choose
statuses. A holdings row or a matching historical price is not cost provenance.
"""
from collections import deque
from decimal import Decimal, ROUND_HALF_UP
from datetime import date, timedelta
from xlsx_reader import read_sheet

D = lambda x: Decimal(str(x))
CENT = Decimal('.01')


def money(value):
    return D(value).quantize(CENT, rounding=ROUND_HALF_UP)


def parse_day(value):
    text = str(value).strip()
    if text.replace('.', '', 1).isdigit():
        return date(1899, 12, 30) + timedelta(days=int(D(text)))
    return date.fromisoformat(text[:10].replace('/', '-'))


def read_trades(path):
    header, rows = read_sheet(str(path))
    code_col = header.index('代码')
    return [dict(zip(header, row)) for row in rows
            if len(row) > code_col and row[code_col].isdigit()]


def new_lot(trade_id, qty, cost):
    qty, cost = D(qty), money(cost)
    if qty <= 0:
        raise ValueError('买入数量必须大于0')
    return {'trade_id': str(trade_id), 'quantity': qty, 'cost': cost,
            'used_quantity': D(0), 'used_cost': D(0)}


def consume(lots, quantity):
    """FIFO within the explicitly authorized lot queue; cost conserved to cents."""
    need = D(quantity)
    if need <= 0:
        raise ValueError('卖出数量必须大于0')
    if sum(x['quantity'] - x['used_quantity'] for x in lots) < need:
        raise ValueError('历史可用买入不足：缺少期初持仓或历史交易，不能把缺口成本当0')
    cost, parts = D(0), []
    while need:
        lot = lots[0]
        take = min(need, lot['quantity'] - lot['used_quantity'])
        used = lot['used_quantity'] + take
        cumulative = money(lot['cost'] * used / lot['quantity'])
        part = cumulative - lot['used_cost']
        lot['used_quantity'], lot['used_cost'] = used, cumulative
        cost += part
        parts.append({'buy_id': lot['trade_id'], 'quantity': str(take), 'cost': str(part)})
        need -= take
        if used == lot['quantity']:
            lots.popleft()
    return cost, parts


def replay(paths, code, cutoff, statuses, expected_quantity, opening=()):
    """Settlement cutoff, chronological FIFO, transaction-ID deduplication.

No implicit status list: including 交易执行中 must be authorized by the user.
The caller must establish that the supplied period/opening inventory is complete.
"""
    if not statuses:
        raise ValueError('必须显式指定纳入的交易状态')
    cutoff = parse_day(cutoff)
    records = {}
    for path in paths:
        for t in read_trades(path):
            if t['代码'] != str(code) or t.get('交易状态') not in statuses:
                continue
            settlement = parse_day(t.get('首期结算日期') or t['交易日期'])
            if settlement > cutoff:
                continue
            key = t['交易编号']
            signature = (t['交易方向'], D(t['交易面额(万)']), D(t['结算金额(元)']), settlement,
                         t.get('交易时间', '').strip(), t.get('交易状态'))
            if key in records and records[key][0] != signature:
                raise ValueError('同一交易编号存在冲突版本：' + key)
            records[key] = (signature, t)
    lots = deque(new_lot(x['trade_id'], x['quantity'], x['cost']) for x in opening)
    audit = []
    for signature, t in sorted(records.values(), key=lambda item:
                               (item[0][3], item[0][4], item[1]['交易编号'])):
        direction = t['交易方向']
        if direction in ('买入', '分销买入'):
            lots.append(new_lot(t['交易编号'], t['交易面额(万)'], t['结算金额(元)']))
            audit.append({'id': t['交易编号'], 'direction': direction, 'status': t['交易状态']})
        elif direction == '卖出':
            cost, parts = consume(lots, t['交易面额(万)'])
            audit.append({'id': t['交易编号'], 'direction': direction, 'status': t['交易状态'],
                          'matched_cost': str(cost), 'parts': parts})
    remaining = [{'trade_id': x['trade_id'],
                  'quantity': str(x['quantity'] - x['used_quantity']),
                  'cost': str(x['cost'] - x['used_cost'])} for x in lots]
    quantity = sum(D(x['quantity']) for x in remaining)
    if quantity != D(expected_quantity):
        raise ValueError(f'历史重放余量{quantity}与已核对持仓余量{expected_quantity}不一致')
    return remaining, audit
