import io
import unittest
from unittest.mock import patch

import openpyxl
import server


class WorkPackageExportTests(unittest.TestCase):
    def export(self, source, task=None):
        task = task or {'id': 'test', 'source_type': 'subsystem', 'source_key': '系统',
                        'task_name': '测试工作包', 'plan_start': '2026-09-23',
                        'plan_finish': '2026-09-24', 'actual_start': '2026-09-21',
                        'actual_finish': '2026-09-26'}
        with patch.object(server, 'pm_snapshot', return_value={'tasks': [task]}), \
             patch.object(server, 'bootstrap', return_value=source), \
             patch.object(server, 'sub_schedule', return_value={'2026-09-23': 10, '2026-09-24': 10}):
            body, _ = server.export_work_package('titan', 'test')
        self.addCleanup(lambda: book.close())
        book = openpyxl.load_workbook(io.BytesIO(body))
        return book

    def test_steps_preserve_all_dates_zero_notes_and_empty_steps(self):
        book = self.export({'steps': [
            {'sub': '系统', 'no': 1, 'step': '安装', 'finish': '2026-09-24', 'note': '原备注',
             'quantities': {'2026-09-23': 12, '2026-09-24': 0},
             'daily_notes': {'2026-09-23': '已安装', '2026-09-25': '复核'}},
            {'sub': '系统', 'no': 2, 'step': '验收'},
            {'sub': '其他系统', 'step': '不得导出', 'quantities': {'2026-09-23': 99}}]})
        rows = list(book['施工步骤'].values)
        self.assertEqual(len(rows), 5)
        self.assertEqual(rows[0][8:], ('填报日期', '当日完成数量', '当日备注'))
        self.assertEqual(rows[1][8:], ('2026-09-23', 12, '已安装'))
        self.assertEqual(rows[2][8:], ('2026-09-24', 0, None))
        self.assertEqual(rows[3][8:], ('2026-09-25', None, '复核'))
        self.assertEqual(rows[1][5], '2026-09-24')
        self.assertEqual(rows[1][7], '原备注')
        self.assertEqual(rows[4][1], '验收')

    def test_actual_dates_and_records_outside_plan(self):
        book = self.export({'sub_settings': {'系统': {'start': '2026-09-23', 'finish': '2026-09-24',
                                                  'actual_start': '2026-09-21', 'actual_finish': '2026-09-26'}},
                            'metrics': [['2026-09-22', '系统', '', 3],
                                        ['2026-09-25', '系统', '', 7],
                                        ['2026-09-25', '系统', '', 2],
                                        ['2026-09-25', '其他系统', '', 99]]})
        fields = dict(book['工作包说明'].iter_rows(min_row=3, values_only=True))
        self.assertEqual(fields['实际开工时间'], '2026-09-21')
        self.assertEqual(fields['实际完工时间'], '2026-09-26')
        rows = list(book['每日计划与实际'].iter_rows(min_row=2, values_only=True))
        self.assertEqual([r[0] for r in rows], ['2026-09-21', '2026-09-22', '2026-09-23',
                                               '2026-09-24', '2026-09-25', '2026-09-26'])
        self.assertEqual(sum(r[2] for r in rows), 12)
        self.assertEqual(rows[-1][3:6], (20, 12, -8))

    def test_manual_package_keeps_actual_logs(self):
        with patch.object(server.pm_core, 'task_actual_logs', return_value=[{'work_date': '2026-09-22', 'qty': 5}]):
            book = self.export({}, {'id': 'test', 'task_name': '手工工作包', 'source_type': 'manual'})
        self.assertEqual(book.sheetnames, ['工作包说明', '实际记录'])
        self.assertEqual(book['实际记录']['B2'].value, 5)


if __name__ == '__main__':
    unittest.main()
