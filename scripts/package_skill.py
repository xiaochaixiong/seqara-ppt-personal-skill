"""Produce a relocatable ZIP with checksums and Seqara's import limits."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import argparse
import hashlib
import json

SKILL_NAME = 'seqara-ppt-personal-skill'

def main():
    p=argparse.ArgumentParser();p.add_argument('--out',required=True);args=p.parse_args()
    root=Path(__file__).resolve().parents[1];out=Path(args.out).resolve()
    if out==root or root in out.parents:raise ValueError('ZIP must be outside skill directory')
    paths=[p for p in root.rglob('*') if p.is_file() and '__pycache__' not in p.parts and '.git' not in p.parts and p.suffix!='.pyc' and p.name!='bundle-manifest.json']
    if len(paths)>1999 or sum(p.stat().st_size for p in paths)>49_000_000:raise ValueError('Seqara import limit exceeded')
    if any(p.is_symlink() or getattr(p,'is_junction',lambda:False)() for p in root.rglob('*')):raise ValueError('Bundle may not contain links')
    manifest={'name':SKILL_NAME,'version':'1.2.0-public.2','files':{p.relative_to(root).as_posix():hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(paths)}}
    mp=root/'bundle-manifest.json';mp.write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n',encoding='utf-8');paths.append(mp)
    out.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(out,'w',ZIP_DEFLATED) as z:
        for p in paths:z.write(p,SKILL_NAME+'/'+p.relative_to(root).as_posix())
    with ZipFile(out) as z:
        if z.testzip() is not None:raise ValueError('ZIP integrity check failed')
        for rel,digest in manifest['files'].items():
            if hashlib.sha256(z.read(SKILL_NAME+'/'+rel)).hexdigest()!=digest:raise ValueError('ZIP content mismatch')
    print(json.dumps({'zip':str(out),'files':len(paths),'uncompressed_bytes':sum(p.stat().st_size for p in paths),'zip_bytes':out.stat().st_size,'sha256':hashlib.sha256(out.read_bytes()).hexdigest()},ensure_ascii=False))

if __name__=='__main__':main()
