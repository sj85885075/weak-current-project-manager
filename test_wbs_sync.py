import tempfile
import unittest
import pm_core


class SubsystemSyncTests(unittest.TestCase):
    def test_add_rename_repeat_and_preserve(self):
        with tempfile.TemporaryDirectory() as root:
            pm_core.generate_wbs(root, 'bg', '测试', ['原有系统'])
            before = pm_core.snapshot(root, 'bg', '测试', [])['tasks'][0]
            sources = {'原有系统': {'plan_qty': 100, 'actual_qty': 25},
                       '78芯链路测试-待协商': {'plan_qty': 78, 'actual_qty': 12,
                           'unit': '芯', 'owner': '测试组', 'plan_start': '2026-09-18',
                           'plan_finish': '2026-09-20'}}
            for _ in range(3):
                pm_core.sync_subsystem_tasks(root, 'bg', '测试', list(sources), sources)
            tasks = pm_core.snapshot(root, 'bg', '测试', [])['tasks']
            self.assertEqual(len(tasks), 2)
            self.assertEqual(tasks[0]['id'], before['id'])
            self.assertEqual(tasks[0]['wbs_code'], before['wbs_code'])
            added = next(t for t in tasks if t['source_key'].startswith('78'))
            self.assertEqual(added['task_name'], '78芯链路测试-待协商')
            self.assertEqual((added['plan_qty'], added['actual_qty'], added['unit']), (78, 12, '芯'))
            self.assertEqual(added['owner'], '测试组')
            sources[added['source_key']]['display_name'] = '新名称'
            pm_core.sync_subsystem_tasks(root, 'bg', '测试', list(sources), sources)
            tasks = pm_core.snapshot(root, 'bg', '测试', [])['tasks']
            self.assertEqual(next(t for t in tasks if t['id']==added['id'])['task_name'], '新名称')

    def test_empty_project(self):
        with tempfile.TemporaryDirectory() as root:
            pm_core.sync_subsystem_tasks(root, 'bg', '测试', ['新增'], {'新增': {}})
            tasks = pm_core.snapshot(root, 'bg', '测试', [])['tasks']
            self.assertEqual(len(tasks), 1)
            self.assertEqual(tasks[0]['task_name'], '新增')


if __name__ == '__main__':
    unittest.main()
