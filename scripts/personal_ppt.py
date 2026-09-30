#!/usr/bin/env python3
"""Host-neutral CLI for Seqara and Codex. Run with Python 3.10+."""
from pathlib import Path
import argparse
import json
import sys
import importlib.util
if importlib.util.find_spec('lxml') is None:
    print(json.dumps({'status':'error','error':'Missing lxml. Install this skill requirements.txt with the same Python interpreter.'}))
    raise SystemExit(2)
from native_pptx import *
from pipeline import extract_source, validate_plan, validate_storyboard, build, data_summary
from data_sources import extract_data
from transitions import suggest_transitions, apply_transitions
from media_assets import (write_inventory, verify_inventory, draft_project_catalog,
                          write_project_matches, verify_project_matches,
                          suggest_photo_layout, suggest_project_photo_pages,
                          validate_asset_bindings)
from quality import doctor, verify, review_init, finalize, render_powerpoint, compare_images

def draft(source_paths, numbers):
    inv=inventory();pages=[]
    for i,n in enumerate(numbers,1):
        page={'id':f'page-{i:02d}','template_slide':n,'purpose':'填写本页结论','slots':{},'pictures':{},'notes_refs':[],'coverage':[]}
        for s in inv['slides'][n-1]['slots']:
            if s['text'] and s['name']!='page-num' and s['type']!='grpSp':
                page['slots'][s['key']]={'action':'replace','kind':'gap','paragraphs':[{'runs':[{'text':'','style_run':0}]}],'sources':[]}
            if s.get('role')=='evidence':page['pictures'][s['key']]={'action':'remove','reason':'等待本期真实照片'}
        if n==7:page['chart']={'action':'replace','categories':[],'values':[],'fact_ids':[]}
        pages.append(page)
    return {'schema_version':1,'title':'演示汇报','template_sha256':inv['template_sha256'],
            'sources':source_paths,'facts':[],'pages':pages,'gaps':[]}

def main():
    p=argparse.ArgumentParser(description=__doc__);sub=p.add_subparsers(dest='command',required=True)
    sub.add_parser('doctor')
    a=sub.add_parser('inspect');a.add_argument('--out',required=True)
    a=sub.add_parser('extract');a.add_argument('input');a.add_argument('--id',required=True);a.add_argument('--role',choices=['primary','evidence','reference'],default='primary');a.add_argument('--out',required=True)
    a=sub.add_parser('extract-data');a.add_argument('input');a.add_argument('--id',required=True);a.add_argument('--out',required=True)
    a=sub.add_parser('draft');a.add_argument('--source',action='append',required=True);a.add_argument('--pages',required=True,help='Comma-separated reference page numbers');a.add_argument('--out',required=True)
    a=sub.add_parser('check-plan');a.add_argument('plan')
    a=sub.add_parser('check-data');a.add_argument('plan',help='Plan JSON with optional data_sources and metrics')
    a=sub.add_parser('check-story');a.add_argument('plan',help='Plan JSON with storyboard metadata')
    a=sub.add_parser('suggest-transitions');a.add_argument('plan')
    a=sub.add_parser('apply-transitions');a.add_argument('plan');a.add_argument('--out',required=True)
    a=sub.add_parser('media-inventory');a.add_argument('input');a.add_argument('--id',required=True);a.add_argument('--out',required=True)
    a=sub.add_parser('media-project-candidates');a.add_argument('inventory');a.add_argument('--out',required=True)
    a=sub.add_parser('media-match');a.add_argument('inventory');a.add_argument('--projects',required=True);a.add_argument('--out',required=True)
    a=sub.add_parser('media-layout');a.add_argument('inventory');a.add_argument('--assets',required=True);a.add_argument('--out',required=True)
    a=sub.add_parser('media-layout-projects');a.add_argument('matches');a.add_argument('--out',required=True)
    a=sub.add_parser('media-check');a.add_argument('plan')
    a=sub.add_parser('build');a.add_argument('plan');a.add_argument('--out',required=True)
    a=sub.add_parser('verify');a.add_argument('directory');a.add_argument('--structure-only',action='store_true')
    a=sub.add_parser('review-init');a.add_argument('directory')
    a=sub.add_parser('finalize');a.add_argument('directory')
    a=sub.add_parser('render-reference');a.add_argument('--out',required=True)
    a=sub.add_parser('compare');a.add_argument('--reference',required=True);a.add_argument('--output',required=True);a.add_argument('--pairs',required=True,help='1:1,3:3,5:5,11:11,19:19,36:35');a.add_argument('--out',required=True)
    args=p.parse_args();cmd=args.command
    if cmd=='doctor':result=doctor()
    elif cmd=='inspect':result=inventory();write_json(args.out,result);result={'profile':str(Path(args.out).resolve()),'slides':len(result['slides'])}
    elif cmd=='extract':
        result=extract_source(args.input,args.id,args.role)
        import os
        result['path']=os.path.relpath(Path(args.input).resolve(),Path(args.out).resolve().parent)
        write_json(args.out,result);result={'source':str(Path(args.out).resolve()),'fragments':len(result['fragments'])}
    elif cmd=='extract-data':
        result=extract_data(args.input,args.id)
        import os
        result['path']=os.path.relpath(Path(args.input).resolve(),Path(args.out).resolve().parent)
        write_json(args.out,result)
        result={'data_source':str(Path(args.out).resolve()),'cells':len(result['cells']),
                'format':result['format']}
    elif cmd=='draft':result=draft(args.source,[int(n) for n in args.pages.split(',')]);write_json(args.out,result);result={'plan':str(Path(args.out).resolve()),'status':'needs_content_and_coverage'}
    elif cmd=='check-plan':
        plan=read_json(args.plan)
        sources,refs,inputs,gaps=validate_plan(plan,Path(args.plan).resolve().parent)
        result={'valid':True,'sources':len(sources),'fragments':len(refs),'gaps':gaps}
        story=validate_storyboard(plan)
        if story is not None:result['storyboard']=story
        if plan.get('data_sources') or plan.get('metrics'):
            summary=data_summary(plan,Path(args.plan).resolve().parent)
            result['data']={'data_sources':summary['data_sources'],'cells':summary['cells'],
                            'metrics':len(summary['metrics'])}
    elif cmd=='check-data':
        plan=read_json(args.plan)
        result=data_summary(plan,Path(args.plan).resolve().parent)
    elif cmd=='check-story':result=validate_storyboard(read_json(args.plan),require=True)
    elif cmd=='suggest-transitions':result=suggest_transitions(read_json(args.plan))
    elif cmd=='apply-transitions':
        source=Path(args.plan).resolve();output=Path(args.out).resolve()
        if source.parent!=output.parent:raise ValueError('Transition output must be beside the plan so relative sources stay valid')
        result=apply_transitions(read_json(source));write_json(output,result)
        result={'plan':str(output),'pages':len(result['pages']),'transitions':len(result.get('transitions',[]))}
    elif cmd=='media-inventory':
        ledger=write_inventory(args.input,args.id,args.out)
        result={'inventory':str(Path(args.out).resolve()),'occurrences':len(ledger['assets']),
                'ready':sum(a['status']=='ready' for a in ledger['assets'])}
    elif cmd=='media-project-candidates':
        result=draft_project_catalog(verify_inventory(args.inventory));write_json(args.out,result)
        result={'catalog':str(Path(args.out).resolve()),'projects':len(result['projects']),'review_required':True}
    elif cmd=='media-match':
        report=write_project_matches(args.inventory,args.projects,args.out)
        result={'matches':str(Path(args.out).resolve()),'review_html':str(Path(args.out).resolve().with_suffix('.review.html')),
                'automatic':sum(x['status']=='auto' for x in report['decisions']),
                'review':sum(x['status']=='review' for x in report['decisions'])}
    elif cmd=='media-layout':
        result=suggest_photo_layout([x.strip() for x in args.assets.split(',')],verify_inventory(args.inventory));write_json(args.out,result)
    elif cmd=='media-layout-projects':
        report,inventory=verify_project_matches(args.matches)
        result=suggest_project_photo_pages(inventory,report);write_json(args.out,result)
    elif cmd=='media-check':
        path=Path(args.plan).resolve();result=validate_asset_bindings(read_json(path),path.parent)
    elif cmd=='build':result=build(args.plan,args.out)
    elif cmd=='verify':result=verify(args.directory,not args.structure_only)
    elif cmd=='review-init':result=review_init(args.directory)
    elif cmd=='finalize':result=finalize(args.directory)
    elif cmd=='render-reference':
        if not doctor()['ready_for_verified_delivery']:raise ValueError('Run doctor and resolve dependencies/fonts')
        result=render_powerpoint(TEMPLATE,Path(args.out).resolve());write_json(Path(args.out)/'reference-render.json',result)
        result={'directory':str(Path(args.out).resolve()),'slides':len(result['slides'])}
    elif cmd=='compare':
        result=compare_images(args.reference,args.output,[tuple(map(int,x.split(':'))) for x in args.pairs.split(',')]);write_json(args.out,result)
    print(json.dumps(result,ensure_ascii=False,indent=2))
    if cmd=='verify' and result['errors']:return 2
    if cmd=='doctor' and not result['ready_for_verified_delivery']:return 2
    return 0

if __name__=='__main__':
    try:
        if hasattr(sys.stdout,'reconfigure'):sys.stdout.reconfigure(encoding='utf-8')
        sys.exit(main())
    except Exception as exc:
        print(json.dumps({'status':'error','error':str(exc)},ensure_ascii=False),file=sys.stderr)
        sys.exit(2)
