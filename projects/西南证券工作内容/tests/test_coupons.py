import datetime
import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from coupon_adjustments import apply_coupons


class CouponTests(unittest.TestCase):
    def test_confirmed_receipt_and_duplicate_guard(self):
        with tempfile.TemporaryDirectory() as d:
            file = Path(d)/'coupon.json'
            file.write_text(json.dumps({'date': '2026-01-02', 'source': '合成审核记录',
                                        'payments': {'012345': 2.01}}), encoding='utf-8')
            details = [{'代码': '012345', '结算金额差': -1.01}]
            apply_coupons(details, file, datetime.date(2026, 1, 2))
            self.assertEqual(details[0]['结算金额差'], 1)
            self.assertEqual(details[0]['不含付息结算金额差'], -1.01)
            with self.assertRaises(ValueError):
                apply_coupons(details, file, datetime.date(2026, 1, 2))

    def test_wrong_date_or_unknown_bond_does_not_change_details(self):
        with tempfile.TemporaryDirectory() as d:
            file = Path(d)/'coupon.json'
            file.write_text(json.dumps({'date': '2026-01-02', 'source': '合成审核记录',
                                        'payments': {'012345': 2, '999999': 1}}), encoding='utf-8')
            details = [{'代码': '012345', '结算金额差': -1}]
            with self.assertRaises(ValueError):
                apply_coupons(details, file, datetime.date(2026, 1, 3))
            with self.assertRaises(ValueError):
                apply_coupons(details, file, datetime.date(2026, 1, 2))
            self.assertEqual(details, [{'代码': '012345', '结算金额差': -1}])
