"""Local account authentication with salted password hashes and memory sessions."""
import hashlib
import hmac
import json
import os
import secrets
import threading
import time
from http.cookies import SimpleCookie
from urllib.parse import urlparse


class Auth:
    def __init__(self, base):
        self.path = os.path.join(base, 'account_credentials.json')
        self.lock = threading.RLock()
        self.sessions = {}
        self.failures = {}

    def account(self):
        try:
            with open(self.path, encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            return None

    @staticmethod
    def digest(password, salt):
        return hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), bytes.fromhex(salt), 600000).hex()

    def session(self, handler):
        cookie = SimpleCookie()
        try:
            cookie.load(handler.headers.get('Cookie', ''))
            token = cookie['bg_session'].value if 'bg_session' in cookie else ''
        except Exception:
            return None
        with self.lock:
            now = time.time()
            self.sessions = {k: v for k, v in self.sessions.items() if v['expires'] > now}
            return self.sessions.get(token)

    def reply(self, handler, code, data, cookie=None):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        handler.send_response(code)
        handler.send_header('Content-Type', 'application/json; charset=utf-8')
        handler.send_header('Content-Length', str(len(body)))
        handler.send_header('Cache-Control', 'no-store')
        if cookie:
            handler.send_header('Set-Cookie', cookie)
        handler.end_headers()
        handler.wfile.write(body)

    def gate(self, handler):
        path = urlparse(handler.path).path
        session = self.session(handler)
        if handler.command == 'GET' and path == '/api/auth/status':
            self.reply(handler, 200, {'setup': self.account() is None, 'authenticated': bool(session), 'username': session['username'] if session else '', 'version': '2026.10.18-performance', 'app_dir': os.path.dirname(self.path)})
            return False
        if handler.command == 'POST':
            origin = handler.headers.get('Origin')
            host = handler.headers.get('Host', '')
            if host not in ('localhost:%d' % handler.server.server_port, '127.0.0.1:%d' % handler.server.server_port) or (origin and origin != 'http://' + host) or handler.headers.get('Sec-Fetch-Site') == 'cross-site':
                self.reply(handler, 403, {'ok': False, 'err': '请求来源无效'})
                return False
            if path.startswith('/api/auth/'):
                self.post(handler, path, session)
                return False
        if handler.command == 'GET' and path == '/login':
            handler._send_file('login.html', 'text/html; charset=utf-8')
            return False
        if session:
            return True
        if path.startswith('/api/'):
            self.reply(handler, 401, {'ok': False, 'err': '请先登录'})
        else:
            handler.send_response(302)
            handler.send_header('Location', '/login')
            handler.send_header('Content-Length', '0')
            handler.end_headers()
        return False

    def post(self, handler, path, session):
        try:
            length = int(handler.headers.get('Content-Length', '0'))
            if not 0 < length <= 4096:
                raise ValueError('请求大小无效')
            data = json.loads(handler.rfile.read(length))
            if not isinstance(data, dict):
                raise ValueError('请求格式无效')
            with self.lock:
                if path == '/api/auth/logout':
                    if session:
                        self.sessions = {k: v for k, v in self.sessions.items() if v is not session}
                    self.reply(handler, 200, {'ok': True}, 'bg_session=; HttpOnly; SameSite=Strict; Path=/; Max-Age=0')
                    return
                if path not in ('/api/auth/setup', '/api/auth/login'):
                    self.reply(handler, 404, {'ok': False})
                    return
                username, password = data.get('username'), data.get('password')
                if not isinstance(username, str) or not isinstance(password, str) or len(password) > 256:
                    raise ValueError('账号或密码格式无效')
                username = username.strip()
                account = self.account()
                now = time.time()
                attempts = [t for t in self.failures.get(handler.client_address[0], []) if now - t < 300]
                if len(attempts) >= 5:
                    self.reply(handler, 429, {'ok': False, 'err': '尝试次数过多，请5分钟后重试'})
                    return
                if path == '/api/auth/setup':
                    if account is not None:
                        raise ValueError('账号已设置，请登录')
                    if handler.client_address[0] not in ('127.0.0.1', '::1'):
                        raise ValueError('请在本机设置账号')
                    if not 3 <= len(username) <= 64 or len(password) < 12:
                        raise ValueError('账号需3至64字符，密码至少12字符')
                    salt = secrets.token_hex(32)
                    account = {'username': username, 'salt': salt, 'hash': self.digest(password, salt), 'algorithm': 'pbkdf2_sha256', 'iterations': 600000}
                    with open(self.path, 'x', encoding='utf-8') as f:
                        json.dump(account, f, indent=2)
                elif account is None or not (hmac.compare_digest(self.digest(password, account['salt']), account['hash']) and hmac.compare_digest(username.encode(), account['username'].encode())):
                    attempts.append(now)
                    self.failures[handler.client_address[0]] = attempts
                    self.reply(handler, 401, {'ok': False, 'err': '账号或密码错误'})
                    return
                self.failures.pop(handler.client_address[0], None)
                token = secrets.token_urlsafe(32)
                self.sessions[token] = {'username': username, 'expires': now + 8 * 3600}
                self.reply(handler, 200, {'ok': True}, 'bg_session=%s; HttpOnly; SameSite=Strict; Path=/; Max-Age=28800' % token)
        except (ValueError, TypeError, OSError):
            self.reply(handler, 400, {'ok': False, 'err': '请求无效，请检查账号长度和密码长度（至少12字符）；已设置账号请登录'})
