import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parent))
import management
import salary_costs
import org_deliver
import openpyxl


class SalarySyncTests(unittest.TestCase):
    def person(self, **values):
        return dict(pid='P1', name='甲', daily_rate=180000, work_days=2,
                    salary_currency='IDR', status='在职',
                    first_work_date='2026-11-13', last_work_date='2026-11-14', **values)

    def test_reconcile_idempotent_correction_deletion_and_isolation(self):
        with tempfile.TemporaryDirectory() as root, tempfile.TemporaryDirectory() as other:
            manual = dict(name='交通', category='交通', amount=200, currency='IDR', date='2026-11-14')
            management.mutate(root, dict(kind='costs', action='save', row=manual))
            p = self.person()
            with patch.object(salary_costs, 'read_people', return_value=[p]):
                self.assertEqual(salary_costs.sync(root)['count'], 1)
                self.assertIn('新增0条', salary_costs.sync(root)['msg'])
                costs = management.snapshot(root)['costs']
                salary = next(r for r in costs if r.get('source'))
                self.assertEqual(salary['amount'], 360000)
                with self.assertRaisesRegex(ValueError, '自动同步'):
                    management.mutate(root, dict(kind='costs', action='delete', id=salary['id']))
                p['work_days'] = 3
                salary_costs.sync(root)
                costs = management.snapshot(root)['costs']
                self.assertEqual(next(r for r in costs if r.get('source'))['amount'], 540000)
                p['daily_rate'] = 200000
                self.assertEqual(salary_costs.sync(root)['totals']['IDR'], 600000)
                p['status'] = '已删除'
                self.assertEqual(salary_costs.sync(root)['count'], 0)
                self.assertEqual(len(management.snapshot(root)['costs']), 1)
            with patch.object(salary_costs, 'read_people', return_value=[]):
                self.assertEqual(management.snapshot(other)['costs'], [])

    def test_missing_duplicate_nonfinite_zero_and_rounding(self):
        people = [self.person()]
        people[0]['daily_rate'] = '1.235'
        rows, warnings = salary_costs.desired_rows(people)
        self.assertEqual(next(iter(rows.values()))['amount'], 2.47)
        for rate in (None, '', 'NaN', 'Infinity', -1):
            people[0]['daily_rate'] = rate
            self.assertEqual(salary_costs.desired_rows(people)[0], {})
        people[0]['daily_rate'] = 0
        self.assertEqual(salary_costs.desired_rows(people)[0], {})
        people[0]['daily_rate'] = 100
        people[0]['salary_currency'] = ''
        self.assertTrue(salary_costs.desired_rows(people)[1])
        people[0]['salary_currency'] = 'CNY'
        duplicate = dict(people[0], pid='P2', name=' 甲 ')
        self.assertEqual(salary_costs.desired_rows(people + [duplicate])[0], {})

    def test_real_workbook_deduplicated_attendance(self):
        with tempfile.TemporaryDirectory() as root:
            book = openpyxl.Workbook(); ws = book.active; ws.title = '人员信息'
            ws.append(org_deliver.ORG_HEADERS)
            p = self.person(); ws.append([p.get(k) for k in org_deliver.ORG_KEYS])
            book.save(Path(root, '26_项目人员组织架构.xlsx')); book.close()
            book = openpyxl.Workbook(); ws = book.active; ws.title = '人员明细'
            ws.append(['日期', '姓名', '工时(h)'])
            for row in [('2026-11-13', '甲', 8), ('2026-11-13', '甲', 2),
                        ('2026-11-14', '甲', 0), ('2026-11-15', '甲', None)]:
                ws.append(row)
            book.save(Path(root, '29_项目人员管理表.xlsx')); book.close()
            result = salary_costs.sync(root)
            self.assertEqual(result['totals'], {'IDR': 360000})
            self.assertEqual(management.snapshot(root)['costs'][0]['date'], '2026-11-15')


if __name__ == '__main__':
    unittest.main()
