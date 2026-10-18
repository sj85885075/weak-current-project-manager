import json
import math
import pathlib
import tempfile
import unittest
import http.client
import threading
import time
import point_map as m


class PointMapTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = pathlib.Path(self.tmp.name)
        data = {'meta': {}, 'devices': [{'id':'d1','name':'AP-1','category':'AP','zone':'1F','x':1,'y':2}],
                'lines':[{'id':'l1','name':'L-1','category':'网线','zone':'1F','points':[[1,2],[5,2]],'drawing_length_m':4}], 'background':[]}
        (self.root/'点位图基础数据.json').write_text(json.dumps(data),encoding='utf8')

    def tearDown(self):
        self.tmp.cleanup()

    def edit(self, id, rev, **patch):
        return m.mutate(self.root, {'action':'update','id':id,'revision':rev,'patch':patch})

    def test_persistence_and_concurrent_edit(self):
        self.edit('d1',0,installed=True,commissioned=True,name='AP-001',note='安装完成')
        r=m.snapshot(self.root)['devices'][0]
        self.assertTrue(r['installed'] and r['commissioned'])
        self.assertEqual(r['name'],'AP-001')
        self.edit('d1',1,note='再次保存')
        self.assertEqual(m.snapshot(self.root)['devices'][0]['name'],'AP-001')
        with self.assertRaises(ValueError):self.edit('d1',1,installed=False)
        self.assertTrue(m.snapshot(self.root)['devices'][0]['installed'])
        with tempfile.TemporaryDirectory() as another:
            self.assertEqual(m.snapshot(another)['devices'],[])

    def test_invalid_states_and_lengths_leave_records_unchanged(self):
        for patch in ({'commissioned':True},{'installed':'yes'},{'measured_length_m':10}):
            with self.assertRaises(ValueError):self.edit('d1',0,**patch)
        for value in (-1,math.nan,math.inf,True,'12'):
            with self.assertRaises(ValueError):self.edit('l1',0,measured_length_m=value)
        self.edit('l1',0,laid=True,measured_length_m=15.75)
        r=m.snapshot(self.root)['lines'][0]
        self.assertTrue(r['laid']);self.assertEqual(r['drawing_length_m'],4)
        self.assertEqual(r['measured_length_m'],15.75)
        self.edit('l1',1,laid=False,measured_length_m=None)
        self.assertIsNone(m.snapshot(self.root)['lines'][0]['measured_length_m'])

    def test_manual_geometry_survives_source_refresh(self):
        r=m.mutate(self.root,dict(action='create',kind='devices',patch=dict(name='机箱-1',category='机箱',zone='1F',points=[[4,5]])))
        self.edit(r['id'],1,installed=True)
        items=m.snapshot(self.root)['devices']
        self.assertEqual(len(items),2);self.assertTrue(items[-1]['installed']);self.assertEqual(items[-1]['x'],4)
        r=m.mutate(self.root,dict(action='create',kind='lines',patch=dict(name='KGB-1',category='KGB管',points=[[0,0],[3,4],[6,4]])))
        line=m.snapshot(self.root)['lines'][-1]
        self.assertEqual(line['drawing_length_m'],8)
        for points in ([[0,0]],[[0,0],[math.nan,1]]):
            with self.assertRaises(ValueError):m.mutate(self.root,dict(action='create',kind='lines',patch=dict(name='bad',category='线',points=points)))

    def test_authenticated_http_routes(self):
        import server
        from auth import Auth
        from http.server import ThreadingHTTPServer
        old_dir,old_auth=server.DIR,server.AUTH
        server.DIR=str(self.root);server.AUTH=Auth(str(self.root))
        server.AUTH.sessions['test']={'username':'test','expires':time.time()+60}
        httpd=ThreadingHTTPServer(('127.0.0.1',0),server.Handler)
        worker=threading.Thread(target=httpd.serve_forever,daemon=True);worker.start()
        def request(path,body=None,authenticated=True,origin=None):
            connection=http.client.HTTPConnection('127.0.0.1',httpd.server_port,timeout=3)
            headers={'Content-Type':'application/json'}
            if authenticated:headers['Cookie']='bg_session=test'
            if origin:headers['Origin']=origin
            connection.request('POST' if body is not None else 'GET',path,json.dumps(body) if body is not None else None,headers)
            response=connection.getresponse();status=response.status;result=json.loads(response.read());connection.close();return status,result
        try:
            self.assertEqual(request('/api/pointmap?proj=bg',authenticated=False)[0],401)
            self.assertEqual(request('/api/pointmap?proj=unknown')[0],400)
            self.assertEqual(len(request('/api/pointmap?proj=bg')[1]['devices']),1)
            payload={'proj':'bg','action':'update','id':'l1','revision':0,'patch':{'laid':True,'measured_length_m':11.2}}
            self.assertEqual(request('/api/pointmap',payload,origin='https://other.example')[0],403)
            self.assertEqual(request('/api/pointmap',payload)[0],200)
            line=request('/api/pointmap?proj=bg')[1]['lines'][0]
            self.assertTrue(line['laid']);self.assertEqual(line['measured_length_m'],11.2)
            self.assertEqual(request('/api/pointmap',payload)[0],400)
        finally:
            httpd.shutdown();httpd.server_close();worker.join();server.DIR=old_dir;server.AUTH=old_auth


if __name__ == '__main__': unittest.main()
