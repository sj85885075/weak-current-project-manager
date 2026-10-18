import tempfile
import unittest
import construction_flow as flow


def snapshot(durations=(3, 5, 2)):
    return {'profile': {'plan_start':'2026-09-01'}, 'tasks':[
        {'source_type':'subsystem','source_key':name,'task_name':name,'wbs_code':str(i+1),
         'plan_start':'2026-09-01','plan_finish':'2026-09-%02d'%duration}
        for i,(name,duration) in enumerate(zip('ABC',durations))]}


def edge(a,b,kind='FS',lag=0):
    return dict(source=a,target=b,type=kind,lag=lag)


class ConstructionFlowTests(unittest.TestCase):
    def test_serial_and_critical_path(self):
        m=flow.calculate(snapshot(),{'edges':[edge('A','B'),edge('B','C')]})
        self.assertEqual(m['total'],10)
        self.assertEqual(m['finish'],'2026-09-10')
        self.assertTrue(all(n['float']==0 for n in m['nodes']))
        self.assertFalse(m['parallel'])

    def test_parallel_not_sum(self):
        m=flow.calculate(snapshot(),{'edges':[]})
        self.assertEqual(m['total'],5)
        self.assertEqual(len(m['parallel']),3)
        self.assertEqual({n['id']:n['float'] for n in m['nodes']},{'A':2,'B':0,'C':3})

    def test_lag_ss_ff_and_lead(self):
        for relation,total,start in [(edge('A','B','FS',2),10,5),(edge('A','B','SS',2),7,2),
                                     (edge('B','A','FF'),5,2),(edge('A','B','FS',-1),7,2)]:
            m=flow.calculate(snapshot(),{'edges':[relation]})
            self.assertEqual(m['total'],total)
            self.assertEqual(next(n for n in m['nodes'] if n['id']==relation['target'])['start'],start)

    def test_join_and_cycle(self):
        m=flow.calculate(snapshot(),{'edges':[edge('A','C'),edge('B','C')]})
        self.assertEqual(m['total'],7)
        with self.assertRaisesRegex(ValueError,'循环'):
            flow.calculate(snapshot(),{'edges':[edge('A','B'),edge('B','A')]})
        with self.assertRaisesRegex(ValueError,'自身'):
            flow.calculate(snapshot(),{'edges':[edge('A','A')]})

    def test_missing_duration_blocks_successor(self):
        snap=snapshot();snap['tasks'][0]['plan_finish']=''
        m=flow.calculate(snap,{'edges':[edge('A','B')]})
        self.assertFalse(m['complete'])
        self.assertTrue(all(n['start'] is None for n in m['nodes'] if n['id'] in 'AB'))
        self.assertTrue(flow.calculate(snap,{'durations':{'A':4}})['complete'])

    def test_dates_sync_and_wbs_predecessor(self):
        snap=snapshot();snap['tasks'][1]['predecessor']='1'
        self.assertEqual(flow.calculate(snap)['total'],8)
        snap['tasks'][0]['plan_finish']='2026-09-04'
        self.assertEqual(flow.calculate(snap)['total'],9)

    def test_empty_and_relative(self):
        self.assertFalse(flow.calculate({'tasks':[]})['complete'])
        snap=snapshot();snap['profile']={}
        for t in snap['tasks']:t['plan_start']=t['plan_finish']=''
        m=flow.calculate(snap,{'durations':{'A':2,'B':3,'C':4}})
        self.assertEqual(m['start'],'')
        self.assertEqual(m['total'],4)

    def test_report_and_project_isolation(self):
        with tempfile.TemporaryDirectory() as a, tempfile.TemporaryDirectory() as b:
            flow.save(a,{'edges':[edge('A','B')]})
            self.assertEqual(flow.load(b),{})
            m=flow.calculate(snapshot(),flow.load(a))
            report=flow.report(m,'<script>测试</script>')
            self.assertNotIn('<script>测试</script>',report)
            self.assertIn('&lt;script&gt;',report)
            self.assertEqual(report.count('<svg '),2)
            self.assertIn('工期甘特图',report)


if __name__=='__main__':unittest.main()
