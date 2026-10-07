"""Launch the browser capture tool with checked dependencies."""
import runpy
from _bootstrap import prepare

if __name__ == '__main__':
    root = prepare('browser')
    runpy.run_path(str(root / 'tools' / 'site_capture' / 'site_dump.py'), run_name='__main__')
