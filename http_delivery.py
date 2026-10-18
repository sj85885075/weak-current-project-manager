"""Authenticated, conditional static delivery with bounded precompressed files."""
import collections,gzip,os,threading

_files=collections.OrderedDict();_lock=threading.Lock();_bytes=0;LIMIT=64*1024*1024

def gzip_ok(handler):
    for part in handler.headers.get('Accept-Encoding','').lower().split(','):
        pieces=part.strip().split(';')
        if pieces[0]!='gzip':continue
        try:return not any(p.strip().startswith('q=') and float(p.strip()[2:])==0 for p in pieces[1:])
        except ValueError:return False
    return False

def send_bytes(handler,code,body,content_type,cache='no-store'):
    encoded=gzip.compress(body,compresslevel=3,mtime=0) if len(body)>=2048 and gzip_ok(handler) else None
    handler.send_response(code);handler.send_header('Content-Type',content_type)
    handler.send_header('Cache-Control',cache);handler.send_header('Vary','Accept-Encoding')
    if encoded is not None:handler.send_header('Content-Encoding','gzip');body=encoded
    handler.send_header('Content-Length',str(len(body)));handler.end_headers();handler.wfile.write(body)

def send_file(handler,path,content_type):
    global _bytes
    stat=os.stat(path);etag='W/"%x-%x"'%(stat.st_mtime_ns,stat.st_size)
    if handler.headers.get('If-None-Match')==etag:
        handler.send_response(304);handler.send_header('ETag',etag);handler.send_header('Cache-Control','private, max-age=0, must-revalidate');handler.send_header('Vary','Accept-Encoding');handler.end_headers();return
    key=(os.path.abspath(path),stat.st_mtime_ns,stat.st_size)
    compress=any(token in content_type for token in ('javascript','text/','svg','json'))
    with _lock:
        cached=_files.get(key)
        if cached:_files.move_to_end(key)
    if cached is None:
        with open(path,'rb') as stream:raw=stream.read()
        zipped=gzip.compress(raw,compresslevel=3,mtime=0) if compress and len(raw)>=2048 else None
        cached=(raw,zipped)
        with _lock:
            # Drop outdated versions of this path.
            for previous in list(_files):
                if previous[0]==key[0]:
                    a,b=_files.pop(previous);_bytes-=len(a)+len(b or b'')
            size=len(raw)+len(zipped or b'')
            if size<=LIMIT:
                _files[key]=cached;_bytes+=size
                while _bytes>LIMIT:
                    _,(a,b)=_files.popitem(last=False);_bytes-=len(a)+len(b or b'')
    raw,zipped=cached;use_gzip=zipped is not None and gzip_ok(handler);body=zipped if use_gzip else raw
    handler.send_response(200);handler.send_header('Content-Type',content_type);handler.send_header('Content-Length',str(len(body)));handler.send_header('Cache-Control','private, max-age=0, must-revalidate');handler.send_header('ETag',etag);handler.send_header('Vary','Accept-Encoding');handler.send_header('X-Content-Type-Options','nosniff')
    if use_gzip:handler.send_header('Content-Encoding','gzip')
    handler.end_headers();handler.wfile.write(body)
