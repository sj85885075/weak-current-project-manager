"""Project document library. Files are immutable; metadata is stored in SQLite."""
import datetime
import os
import re
import sqlite3
import uuid
import subprocess
import time
from contextlib import contextmanager
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

CATEGORIES = ['甲方资料', '商务资料', '技术资料', '图纸资料', '项目资料']
MAX_SIZE = 100 * 1024 * 1024
PREVIEWS = {'.pdf': 'application/pdf', '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif', '.webp': 'image/webp', '.txt': 'text/plain; charset=utf-8', '.csv': 'text/plain; charset=utf-8', '.log': 'text/plain; charset=utf-8'}
CAD = {'.dwg', '.dxf', '.dwt', '.dws', '.dwf', '.dwfx'}
WPS_TYPES = {
    **dict.fromkeys(('.xls', '.xlsx', '.xlsm', '.xlsb', '.xlt', '.xltx', '.csv', '.et', '.ett', '.ods'), 'et.exe'),
    **dict.fromkeys(('.doc', '.docx', '.docm', '.dot', '.dotx', '.rtf', '.wps', '.wpt', '.odt'), 'wps.exe'),
    **dict.fromkeys(('.ppt', '.pptx', '.pptm', '.pps', '.ppsx', '.pot', '.potx', '.dps', '.dpt', '.odp'), 'wpp.exe'),
}
LOCAL = CAD | set(PREVIEWS) | set(WPS_TYPES)

def find_wps(executable):
    """Resolve installed WPS without relying on Windows file associations."""
    candidates = []
    try:
        import winreg
        for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
                for app in (executable, 'wps.exe'):
                    try:
                        with winreg.OpenKey(hive, 'Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\'+app, 0, winreg.KEY_READ | view) as key:
                            value = winreg.QueryValueEx(key, '')[0]
                            candidates.append(Path(os.path.expandvars(value.strip('"'))).parent / executable)
                    except OSError:
                        pass
    except ImportError:
        pass
    for variable in ('ProgramFiles(x86)', 'ProgramFiles', 'LOCALAPPDATA'):
        base = os.environ.get(variable)
        if not base:
            continue
        for folder in ('WPS Office', 'Kingsoft/WPS Office', 'Kingsoft/WPS', 'Kingsoft'):
            root = Path(base) / folder
            for pattern in ('office6/'+executable, '*/office6/'+executable, executable):
                candidates.extend(sorted(root.glob(pattern), reverse=True))
    return next((p for p in candidates if p.is_file()), None)

def launch_on_desktop(executable, arguments):
    import ctypes
    from ctypes import wintypes as w
    class Startup(ctypes.Structure):
        _fields_ = [('cb',w.DWORD),('reserved',w.LPWSTR),('desktop',w.LPWSTR),('title',w.LPWSTR),('x',w.DWORD),('y',w.DWORD),('cx',w.DWORD),('cy',w.DWORD),('charsx',w.DWORD),('charsy',w.DWORD),('fill',w.DWORD),('flags',w.DWORD),('show',w.WORD),('reserved2size',w.WORD),('reserved2',ctypes.c_void_p),('stdin',w.HANDLE),('stdout',w.HANDLE),('stderr',w.HANDLE)]
    class Process(ctypes.Structure):
        _fields_ = [('process',w.HANDLE),('thread',w.HANDLE),('pid',w.DWORD),('tid',w.DWORD)]
    kernel=ctypes.WinDLL('kernel32',use_last_error=True)
    kernel.CreateProcessW.argtypes=[w.LPCWSTR,w.LPWSTR,ctypes.c_void_p,ctypes.c_void_p,w.BOOL,w.DWORD,ctypes.c_void_p,w.LPCWSTR,ctypes.POINTER(Startup),ctypes.POINTER(Process)]
    kernel.CloseHandle.argtypes=[w.HANDLE]
    startup=Startup();startup.cb=ctypes.sizeof(startup);startup.desktop='winsta0\\default';startup.flags=1;startup.show=1
    process=Process();command=ctypes.create_unicode_buffer(subprocess.list2cmdline([str(executable),*arguments]))
    if not kernel.CreateProcessW(str(executable),command,None,None,False,0,None,str(Path(executable).parent),ctypes.byref(startup),ctypes.byref(process)):
        raise ctypes.WinError(ctypes.get_last_error())
    kernel.CloseHandle(process.thread);kernel.CloseHandle(process.process)

def show_wps_document(name, timeout=20):
    """Restore an existing WPS document window, including hidden instances."""
    import ctypes
    from ctypes import wintypes
    user = ctypes.WinDLL('user32', use_last_error=True)
    callback_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    user.EnumWindows.argtypes = [callback_type, wintypes.LPARAM]
    user.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    user.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user.ShowWindowAsync.argtypes = [wintypes.HWND, ctypes.c_int]
    user.IsWindowVisible.argtypes = [wintypes.HWND]
    user.IsIconic.argtypes = [wintypes.HWND]
    user.SetForegroundWindow.argtypes = [wintypes.HWND]
    user.OpenDesktopW.argtypes = [wintypes.LPCWSTR,wintypes.DWORD,wintypes.BOOL,wintypes.DWORD]
    user.OpenDesktopW.restype = wintypes.HANDLE
    user.EnumDesktopWindows.argtypes = [wintypes.HANDLE,callback_type,wintypes.LPARAM]
    user.CloseDesktop.argtypes = [wintypes.HANDLE]
    user.GetThreadDesktop.argtypes=[wintypes.DWORD]
    user.GetThreadDesktop.restype=wintypes.HANDLE
    user.SetThreadDesktop.argtypes=[wintypes.HANDLE]
    original=user.GetThreadDesktop(ctypes.windll.kernel32.GetCurrentThreadId())
    desktop=user.OpenDesktopW('default',0,False,0xC1)
    if not desktop:
        return False
    try:
        if not user.SetThreadDesktop(desktop):
            return False
        deadline=time.monotonic()+timeout
        while True:
            matches=[]
            @callback_type
            def collect(hwnd,_):
                title=ctypes.create_unicode_buffer(user.GetWindowTextLengthW(hwnd)+1)
                user.GetWindowTextW(hwnd,title,len(title))
                if name.casefold() in title.value.casefold() and 'wps' in title.value.lower():
                    matches.append(hwnd)
                return True
            user.EnumWindows(collect,0)
            for hwnd in matches:
                user.ShowWindowAsync(hwnd,9 if user.IsIconic(hwnd) else 5)
                user.SetForegroundWindow(hwnd)
                if user.IsWindowVisible(hwnd):
                    return True
            if time.monotonic()>=deadline:
                return False
            time.sleep(0.15)
    finally:
        user.SetThreadDesktop(original)
        user.CloseDesktop(desktop)


def open_document(path):
    app = WPS_TYPES.get(path.suffix.lower())
    executable = find_wps(app) if app else None
    if executable:
        try:
            # Launch a separate instance on the user's desktop so a hidden
            # instance on the server's desktop cannot consume the request.
            launch_on_desktop(executable, ['/n',str(path.resolve())])
        except OSError as error:
            raise ValueError('WPS 启动失败，请检查 WPS 是否能正常运行：'+str(error)) from error
        if show_wps_document(path.name):
            return '已显示 WPS 文件窗口：'+path.name
        raise ValueError('已发送打开请求，但未检测到可见的 WPS 文件窗口。请使用“下载”后打开，或检查 WPS 是否有未关闭的提示框。')
    try:
        os.startfile(str(path))
    except OSError as error:
        message = '未检测到 WPS，且默认程序无法打开此文件，请安装 WPS 后重试。' if app else '无法打开：请先安装对应软件，并在 Windows 中设置此格式的默认打开程序。'
        raise ValueError(message) from error
    return '已调用电脑默认关联程序打开资料'

@contextmanager
def database(root):
    folder = Path(root) / '资料管理'
    folder.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(folder / '资料索引.db')
    db.row_factory = sqlite3.Row
    try:
        with db:
            db.execute('CREATE TABLE IF NOT EXISTS documents (id TEXT PRIMARY KEY, name TEXT, category TEXT, size INTEGER, uploaded TEXT, note TEXT, archived INTEGER DEFAULT 0)')
            yield db
    finally:
        db.close()

def listing(root):
    with database(root) as db:
        rows = [dict(row) for row in db.execute('SELECT * FROM documents WHERE archived=0 ORDER BY uploaded DESC, rowid DESC')]
    for row in rows:
        ext = Path(row['name']).suffix.lower()
        row.update(preview=ext in PREVIEWS, cad=ext in CAD, local=ext in LOCAL, wps=ext in WPS_TYPES)
    return {'ok': True, 'files': rows, 'categories': CATEGORIES}

def get_file(root, file_id):
    if not re.fullmatch('[0-9a-f]{32}', file_id):
        raise ValueError('无效的资料编号')
    with database(root) as db:
        row = db.execute('SELECT * FROM documents WHERE id=? AND archived=0', (file_id,)).fetchone()
    if not row:
        raise ValueError('资料不存在或已移出列表')
    path = Path(root) / '资料管理' / file_id / row['name']
    if not path.is_file():
        raise ValueError('资料文件缺失')
    return dict(row), path

def upload(root, name, category, stream, size, note=''):
    if category not in CATEGORIES:
        raise ValueError('请选择资料分类')
    if not 0 < size <= MAX_SIZE:
        raise ValueError('单个文件须大于 0 字节且不超过 100 MB')
    if not name or len(name) > 150 or re.search(r'[<>:"/\\|?*\x00-\x1f]', name) or name.endswith((' ', '.')) or name.split('.')[0].upper() in {'CON', 'PRN', 'AUX', 'NUL', *('COM'+str(i) for i in range(1,10)), *('LPT'+str(i) for i in range(1,10))}:
        raise ValueError('文件名无效，请使用不含特殊符号的文件名（最多150字符）')
    file_id = uuid.uuid4().hex
    folder = Path(root) / '资料管理' / file_id
    folder.mkdir(parents=True)
    path = folder / name
    try:
        with path.open('xb') as output:
            remaining = size
            while remaining:
                chunk = stream.read(min(1024 * 1024, remaining))
                if not chunk:
                    raise ValueError('上传中断，请重新上传')
                output.write(chunk)
                remaining -= len(chunk)
        with database(root) as db:
            db.execute('INSERT INTO documents VALUES (?,?,?,?,?,?,0)', (file_id, name, category, size, datetime.datetime.now().isoformat(timespec='seconds'), str(note)[:2000]))
    except Exception:
        path.unlink(missing_ok=True)
        folder.rmdir()
        raise
    return {'ok': True, 'msg': '资料已上传', 'id': file_id}

def action(root, payload):
    row, path = get_file(root, str(payload.get('id', '')))
    operation = payload.get('action')
    if operation == 'open':
        if path.suffix.lower() not in LOCAL:
            raise ValueError('此格式请下载后浏览')
        return {'ok': True, 'msg': open_document(path)}
    with database(root) as db:
        if operation == 'archive':
            db.execute('UPDATE documents SET archived=1 WHERE id=?', (row['id'],))
        elif operation == 'edit':
            category = payload.get('category')
            if category not in CATEGORIES:
                raise ValueError('请选择资料分类')
            db.execute('UPDATE documents SET category=?,note=? WHERE id=?', (category, str(payload.get('note', ''))[:2000], row['id']))
        else:
            raise ValueError('未知操作')
    return {'ok': True, 'msg': '已移出列表，原文件保留在资料管理文件夹' if operation == 'archive' else '分类和备注已保存'}

def handle_get(handler, root):
    url = urlparse(handler.path)
    q = parse_qs(url.query)
    if url.path == '/api/materials':
        handler._send(200, listing(root))
        return
    row, path = get_file(root, (q.get('id') or [''])[0])
    mime = PREVIEWS.get(path.suffix.lower()) if q.get('preview') == ['1'] else None
    handler.send_response(200)
    handler.send_header('Content-Type', mime or 'application/octet-stream')
    handler.send_header('Content-Disposition', ('inline' if mime else 'attachment') + "; filename*=UTF-8''" + quote(row['name']))
    handler.send_header('Content-Length', str(path.stat().st_size))
    handler.send_header('X-Content-Type-Options', 'nosniff')
    handler.send_header('Cache-Control', 'no-store')
    handler.end_headers()
    with path.open('rb') as file:
        while chunk := file.read(1024 * 1024):
            handler.wfile.write(chunk)

def check_local_request(handler):
    host = handler.headers.get('Host', '')
    if handler.client_address[0] not in ('127.0.0.1', '::1') or urlparse('http://' + host).hostname not in ('localhost', '127.0.0.1', '::1'):
        raise ValueError('请从本机访问资料管理')
    origin = handler.headers.get('Origin')
    if origin and origin != 'http://' + host:
        raise ValueError('请从资料管理页面执行此操作')
