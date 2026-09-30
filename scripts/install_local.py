"""Install this local skill without overwriting an existing skill."""
from pathlib import Path
import argparse
import json
import os
import shutil
import tempfile

SKILL_NAME = 'seqara-ppt-personal-skill'

def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--target',choices=['codex','seqara'],required=True);p.add_argument('--home');a=p.parse_args()
    if a.target=='seqara' and not a.home:p.error('Seqara requires its confirmed harness --home path')
    root=Path(__file__).resolve().parents[1]
    home=Path(a.home).expanduser().resolve() if a.home else Path(os.environ.get('CODEX_HOME',str(Path.home()/'.codex'))).resolve()
    dest=home/'skills'/SKILL_NAME
    if dest.exists():raise SystemExit('Refusing to overwrite existing skill: '+str(dest))
    if home==root or root in home.parents:raise SystemExit('Install destination cannot be inside source skill')
    ignored={'__pycache__','.git','.pytest_cache'}
    files=[f for f in root.rglob('*') if f.is_file() and not ignored.intersection(f.parts) and f.suffix!='.pyc']
    if len(files)>2000 or sum(f.stat().st_size for f in files)>50_000_000:raise SystemExit('Seqara package limit exceeded')
    if any(p.is_symlink() or getattr(p,'is_junction',lambda:False)() for p in root.rglob('*')):raise SystemExit('Source contains a link/junction')
    dest.parent.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='.personal-skill-',dir=home) as staging:
        staged=Path(staging)/SKILL_NAME
        shutil.copytree(root,staged,ignore=shutil.ignore_patterns('__pycache__','*.pyc','.git','.pytest_cache'))
        staged.rename(dest)
    print(json.dumps({'installed':str(dest),'target':a.target,'availability':'next turn'},ensure_ascii=False))

if __name__=='__main__':main()
