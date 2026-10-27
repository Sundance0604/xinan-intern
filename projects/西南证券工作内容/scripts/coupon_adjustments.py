"""Apply dated, human-confirmed receipts once to freshly calculated sell P&L."""
import json
from decimal import Decimal
from pathlib import Path


def apply_coupons(details, config_file, today):
    config = json.loads(Path(config_file).read_text(encoding='utf-8'))
    if config.get('date') != today.isoformat() or not str(config.get('source', '')).strip():
        raise ValueError('付息配置须有当日date及人工确认来源source')
    payments = config.get('payments')
    if not isinstance(payments, dict) or not payments:
        raise ValueError('付息配置payments须为非空的代码→元金额映射')
    planned = []
    used = set()
    for raw_code, raw_amount in payments.items():
        code = str(raw_code).split('.')[0]
        if code in used:
            raise ValueError(f'同一代码重复配置付息：{code}')
        used.add(code)
        amount = Decimal(str(raw_amount))
        if not amount.is_finite() or amount < 0 or amount != amount.quantize(Decimal('.01')):
            raise ValueError(f'付息金额必须为非负的元分金额：{code}')
        matches = [d for d in details if d['代码'] == code]
        if len(matches) != 1 or matches[0].get('结算金额差') is None or matches[0].get('未匹配完'):
            raise ValueError(f'付息须匹配一条已核实成本的卖出汇总；非卖出券付息需另行确认展示：{code}')
        d = matches[0]
        if '付息收入（元）' in d:
            raise ValueError(f'该券已补计付息，不能重复加入：{code}')
        planned.append((d, amount))
    audit = []
    for d, amount in planned:
        before = Decimal(str(d['结算金额差']))
        d['不含付息结算金额差'] = float(before)
        d['付息收入（元）'] = float(amount)
        d['结算金额差'] = float(before + amount)
        audit.append({'代码': d['代码'], '付息收入（元）': float(amount), '原结算金额差': float(before),
                      '补计后结算金额差': d['结算金额差'], 'source': config['source']})
    return audit
