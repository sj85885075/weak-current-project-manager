import server,tempfile,shutil,org_deliver
from pathlib import Path
with tempfile.TemporaryDirectory() as root:
 for p in Path(server.DIR).iterdir():
  if p.suffix in ('.xlsx','.json'):shutil.copy2(p,Path(root,p.name))
 Path(root,'泰坦IDC').mkdir()
 for p in Path(server.DIR,'泰坦IDC').iterdir():
  if p.suffix in ('.xlsx','.json'):shutil.copy2(p,Path(root,'泰坦IDC',p.name))
 server.DIR=root;org_deliver.ROOT=root
 server.ThreadingHTTPServer(('127.0.0.1',8766),server.Handler).serve_forever()
