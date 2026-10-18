import io
import json
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import patch
from urllib.request import Request, urlopen
from urllib.parse import urlencode
from http.server import ThreadingHTTPServer
import materials
import server


class LibraryTests(unittest.TestCase):
    def test_storage_and_validation(self):
        with tempfile.TemporaryDirectory() as root, tempfile.TemporaryDirectory() as other:
            first = materials.upload(root, '图纸.dwg', '图纸资料', io.BytesIO(b'CAD'), 3)['id']
            second = materials.upload(root, '图纸.dwg', '图纸资料', io.BytesIO(b'NEW'), 3)['id']
            self.assertNotEqual(first, second)
            self.assertEqual(materials.get_file(root, first)[1].read_bytes(), b'CAD')
            self.assertEqual(materials.listing(other)['files'], [])
            self.assertTrue(materials.listing(root)['files'][0]['cad'])
            materials.action(root, dict(id=first, action='edit', category='技术资料', note='修订说明'))
            self.assertEqual(materials.get_file(root, first)[0]['category'], '技术资料')
            with patch.object(materials.os, 'startfile') as opener:
                materials.action(root, dict(id=first, action='open'))
                opener.assert_called_once_with(str(materials.get_file(root, first)[1]))
            for name in ['../evil.txt', 'C:\\evil.txt', 'CON.txt', 'bad:stream', 'test.']:
                with self.assertRaises(ValueError):
                    materials.upload(root, name, '项目资料', io.BytesIO(b'x'), 1)
            with self.assertRaises(ValueError):
                materials.upload(root, 'broken.txt', '项目资料', io.BytesIO(b'x'), 2)
            with self.assertRaises(ValueError):
                materials.upload(root, 'big.txt', '项目资料', io.BytesIO(), materials.MAX_SIZE+1)
            self.assertEqual(len(materials.listing(root)['files']), 2)
            executable = materials.upload(root, 'test.exe', '项目资料', io.BytesIO(b'x'), 1)['id']
            with self.assertRaises(ValueError):
                materials.action(root, dict(id=executable, action='open'))

    def test_http_upload_download_and_preview(self):
        with tempfile.TemporaryDirectory() as root:
            with patch.object(server, 'project_root', return_value=root), patch.object(server, 'refresh_subsystems'):
                http = ThreadingHTTPServer(('127.0.0.1', 0), server.Handler)
                worker = threading.Thread(target=http.serve_forever, daemon=True)
                worker.start()
                base = 'http://127.0.0.1:' + str(http.server_port)
                try:
                    query = urlencode(dict(proj='bg', name='说明.txt', category='项目资料', note='测试'))
                    payload = '资料上传测试'.encode('utf-8')
                    with urlopen(Request(base+'/api/materials/upload?'+query, data=payload, headers={'Content-Type':'application/octet-stream', 'Origin':base})) as response:
                        file_id = json.load(response)['id']
                    for suffix in ['', '&preview=1']:
                        with urlopen(base+'/api/materials/file?proj=bg&id='+file_id+suffix) as response:
                            self.assertEqual(response.read(), payload)
                            self.assertIn('inline' if suffix else 'attachment', response.headers['Content-Disposition'])
                    with urlopen(base+'/api/materials?proj=bg') as response:
                        self.assertEqual(len(json.load(response)['files']), 1)
                finally:
                    http.shutdown()
                    http.server_close()
                    worker.join()


if __name__ == '__main__':
    unittest.main()
