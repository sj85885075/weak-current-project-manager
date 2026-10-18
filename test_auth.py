import http.client
import json
import os
import tempfile
import threading
import unittest
from http.server import ThreadingHTTPServer
import server
from auth import Auth


class AuthenticationTest(unittest.TestCase):
    def test_authentication(self):
        original = server.AUTH
        with tempfile.TemporaryDirectory() as folder:
            server.AUTH = Auth(folder)
            httpd = ThreadingHTTPServer(('127.0.0.1', 0), server.Handler)
            worker = threading.Thread(target=httpd.serve_forever, daemon=True)
            worker.start()
            try:
                def request(path, data=None, cookie=None, origin=None):
                    conn = http.client.HTTPConnection('127.0.0.1', httpd.server_port)
                    headers = {'Content-Type': 'application/json'}
                    if cookie:
                        headers['Cookie'] = cookie
                    if origin:
                        headers['Origin'] = origin
                    conn.request('POST' if data is not None else 'GET', path, json.dumps(data) if data is not None else None, headers)
                    response = conn.getresponse()
                    result = response.status, response.read(), response.getheader('Set-Cookie')
                    conn.close()
                    return result

                self.assertEqual(request('/api/data')[0], 401)
                self.assertEqual(request('/')[0], 302)
                self.assertEqual(request('/login')[0], 200)
                credentials = {'username': 'admin', 'password': 'TestPassword123!'}
                self.assertEqual(request('/api/auth/setup', credentials, origin='https://evil.example')[0], 403)
                self.assertEqual(request('/api/auth/setup', {'username': 'admin', 'password': 'short'})[0], 400)
                status, body, cookie = request('/api/auth/setup', credentials)
                self.assertEqual(status, 200)
                self.assertIn('HttpOnly', cookie)
                with open(server.AUTH.path) as f:
                    stored = json.load(f)
                self.assertNotIn(credentials['password'], json.dumps(stored))
                self.assertEqual(request('/api/version', cookie=cookie)[0], 200)
                self.assertEqual(request('/api/auth/setup', credentials)[0], 400)
                self.assertEqual(request('/api/auth/logout', {}, cookie)[0], 200)
                self.assertEqual(request('/api/version', cookie=cookie)[0], 401)
                self.assertEqual(request('/api/auth/login', {'username': 'admin', 'password': 'wrong'})[0], 401)
                status, _, cookie = request('/api/auth/login', credentials)
                self.assertEqual(status, 200)
                for session in server.AUTH.sessions.values():
                    session['expires'] = 0
                self.assertEqual(request('/api/version', cookie=cookie)[0], 401)
                for _ in range(5):
                    self.assertEqual(request('/api/auth/login', {'username': 'admin', 'password': 'wrong'})[0], 401)
                self.assertEqual(request('/api/auth/login', credentials)[0], 429)
            finally:
                httpd.shutdown()
                httpd.server_close()
                worker.join()
                server.AUTH = original


if __name__ == '__main__':
    unittest.main()
