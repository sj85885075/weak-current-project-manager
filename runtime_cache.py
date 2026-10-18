"""Bounded, copy-safe caches invalidated by source file signatures."""
import collections,copy,os,threading,time,datetime

class SnapshotCache:
    def __init__(self,limit=8,ttl=30):
        self.limit=limit;self.ttl=ttl;self.items=collections.OrderedDict();self.lock=threading.RLock();self.hits=0;self.misses=0
    def get(self,key,signature):
        with self.lock:
            entry=self.items.get(key)
            if entry is None or entry[0]!=signature or time.monotonic()-entry[1]>self.ttl:
                self.misses+=1;self.items.pop(key,None);return None
            self.hits+=1;self.items.move_to_end(key);return copy.deepcopy(entry[2])
    def put(self,key,signature,value):
        with self.lock:
            self.items[key]=(signature,time.monotonic(),copy.deepcopy(value));self.items.move_to_end(key)
            while len(self.items)>self.limit:self.items.popitem(last=False)
        return value
    def clear(self):
        with self.lock:self.items.clear()

def file_signature(paths):
    result=[]
    for path in sorted(set(map(os.fspath,paths))):
        try:s=os.stat(path);result.append((path,s.st_mtime_ns,s.st_size))
        except FileNotFoundError:result.append((path,None,None))
    return tuple(result)

def project_signature(root,extra=()):
    # Only live ledgers/configs; backups and original CAD binaries do not affect page summaries.
    paths=list(extra)
    for folder in (root,os.path.join(root,'甲方交付资料'),os.path.join(root,'工程审计')):
        try:
            with os.scandir(folder) as entries:
                paths.extend(e.path for e in entries if e.is_file() and e.name.endswith(('.xlsx','.json','.db','-wal')))
        except FileNotFoundError:pass
    return datetime.date.today().isoformat(),file_signature(paths)
