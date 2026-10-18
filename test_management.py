import tempfile
import unittest
import management as m


class LedgerTests(unittest.TestCase):
    def test_persistence_and_project_isolation(self):
        with tempfile.TemporaryDirectory() as a, tempfile.TemporaryDirectory() as b:
            row = dict(name='测试线缆', source='甲方提供', category='线材', quantity='100', unit='米', scope='测试系统', acquired='2026-09-20')
            m.mutate(a, dict(kind='resources', action='save', row=row))
            saved = m.snapshot(a)['resources'][0]
            self.assertEqual(saved['quantity'], 100)
            self.assertEqual(m.snapshot(b)['resources'], [])
            row['quantity'] = '150'
            m.mutate(a, dict(kind='resources', action='save', id=saved['id'], row=row))
            self.assertEqual(m.snapshot(a)['resources'][0]['quantity'], 150)
            for qty in ['-1', 'nan', 'inf', '0']:
                with self.assertRaises(ValueError):
                    m.mutate(a, dict(kind='resources', action='save', row=dict(row, quantity=qty)))
            with self.assertRaises(ValueError):
                m.mutate(a, dict(kind='resources', action='save', row=dict(row, use_start='2026-09-20', use_end='2026-09-19')))
            m.mutate(a, dict(kind='costs', action='save', row=dict(name='工费', category='人工', amount='123.45', currency='CNY', date='2026-09-20')))
            self.assertEqual(m.snapshot(a)['costs'][0]['amount'], 123.45)
            m.mutate(a, dict(kind='resources', action='delete', id=saved['id']))
            self.assertEqual(m.snapshot(a)['resources'], [])
            with self.assertRaises(ValueError):
                m.mutate(a, dict(kind='resources', action='save', id=saved['id'], row=row))


if __name__ == '__main__':
    unittest.main()
