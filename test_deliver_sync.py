import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl
import org_deliver as delivery
import server


class DeliverySyncTests(unittest.TestCase):
    def setUp(self):
        self.data = {
            'sub_settings': {'old': {'name': '新子系统', 'start': '2026-09-23', 'finish': '2026-09-24', 'actual_start': '2026-09-21'}},
            'sub_schedules': {'old': {'2026-09-23': 6, '2026-09-24': 4}},
            'sys_status': {'old': {'total': 10, 'cum': 3, 'unit': '条', 'actual_start': '2026-09-21'}},
            'issues': [['I1','问题']+['']*18+['已复验通过']],
            'steps': [{'sub': 'old', 'step': '安装', 'quantities': {'2026-09-22': 0}, 'daily_notes': {'2026-09-23': '待复验'}}],
        }
        self.snapshot = patch.object(delivery, '_bootstrap', side_effect=lambda proj:self.data)
        self.snapshot.start(); self.addCleanup(self.snapshot.stop)

    def test_live_plan_progress_steps_and_issue_result(self):
        _, headers, rows = delivery.collect_table('titan','construction_plan')
        self.assertEqual(sum(r[2] for r in rows), 10)
        self.assertEqual(rows[0][1], '新子系统')
        self.assertEqual(rows[0][6], '2026-09-21')
        _, headers, rows = delivery.collect_table('titan','sys_status')
        self.assertEqual(rows[0][5], 3)
        self.assertEqual(rows[0][-1], 30)
        _, headers, rows = delivery.collect_table('titan','issues')
        self.assertEqual(headers[3], '优先级')
        self.assertEqual(headers[-1], '问题处理结果')
        self.assertEqual(rows[0][-1], '已复验通过')
        _, headers, rows = delivery.collect_table('titan','steps')
        self.assertEqual([r[-3:] for r in rows], [['2026-09-22',0,''],['2026-09-23','','待复验']])

    def test_daily_and_test_use_daily_records(self):
        records = [['1','2026-09-22','施工','old','布线'],['2','2026-09-22','测试','old','OTDR测试']]
        with patch.object(server, 'load_daily', return_value={'headers':server.DAILY_HEADERS,'records':records}):
            self.assertEqual(len(delivery.collect_table('titan','daily_report')[2]),2)
            tests=delivery.collect_table('titan','test_record')[2]
            self.assertEqual(len(tests),1)
            self.assertEqual(tests[0][3:5],['新子系统','OTDR测试'])

    def test_regenerate_chinese_type_reads_changes_keeps_history(self):
        with tempfile.TemporaryDirectory() as root, patch.object(delivery, 'ROOT', root), \
             patch.object(delivery, '_title_info', return_value={'proj_name':'当前项目','org':'项目部','date':'2026-09-22','no':'TEST'}):
            first=delivery.gen_deliver('titan','progress','xlsx')
            original=Path(first['path']).read_bytes()
            self.data['sys_status']['old']['cum']=7
            second=delivery.gen_deliver('titan','施工进度','xlsx')
            self.assertEqual(second['version'],'V1.1')
            self.assertEqual(Path(first['path']).read_bytes(),original)
            book=openpyxl.load_workbook(second['path'])
            rows=list(book.active.values)
            row=next(r for r in rows if r[0]=='新子系统')
            self.assertEqual(row[5],7)
            self.assertEqual(row[-1],70)
            book.close()
            self.assertEqual(len(delivery.load_manifest('titan')),2)

    def test_unknown_type_rejected(self):
        with self.assertRaises(ValueError): delivery.gen_deliver('titan','未知类型','xlsx')


if __name__=='__main__': unittest.main()
