"""Linux full historical suite. Python 3.9+, standard library only.

Uses the shared HTTP engine; original source documents are never modified.
"""
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import hashlib
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
import zipfile

import run_tests as engine

ROOT = engine.ROOT
S = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def workbook_rows(path):
    """Read all worksheet text without interpreting embedded content as code."""
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            shared = [''.join(t.text or '' for t in si.findall('.//s:t', S))
                      for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('s:si', S)]
        result = {}
        for number in (1, 2):
            cells = {}
            for c in ET.fromstring(z.read(f'xl/worksheets/sheet{number}.xml')).findall('.//s:c', S):
                value = c.find('s:v', S)
                if c.get('t') == 'inlineStr':
                    text = ''.join(t.text or '' for t in c.findall('.//s:t', S))
                elif value is not None:
                    text = shared[int(value.text)] if c.get('t') == 's' else value.text
                else:
                    continue
                cells[c.get('r')] = text
            result[number] = cells
        return result


def legacy_text(path):
    """Try reviewed sidecar, antiword, then isolated LibreOffice conversion."""
    for sidecar in (path.with_suffix('.txt'), path.with_suffix('.docx')):
        if sidecar.exists():
            return (sidecar.read_text(encoding='utf-8-sig') if sidecar.suffix == '.txt'
                    else engine.read_docx(sidecar)), 'sidecar:' + sidecar.name
    errors = []
    if shutil.which('antiword'):
        try:
            run = subprocess.run(['antiword', '-m', 'UTF-8.txt', str(path)],
                                 capture_output=True, timeout=60, check=True)
            text = run.stdout.decode('utf-8').strip()
            if text:
                return text, 'antiword UTF-8'
        except Exception as exc:
            errors.append(str(exc))
    converter = shutil.which('libreoffice') or shutil.which('soffice')
    if converter:
        try:
            with tempfile.TemporaryDirectory() as temp:
                profile = (Path(temp) / 'profile').as_uri()
                subprocess.run([converter, '-env:UserInstallation=' + profile, '--headless',
                                '--convert-to', 'docx', '--outdir', temp, str(path)],
                               capture_output=True, timeout=120, check=True)
                return engine.read_docx(Path(temp) / (path.stem + '.docx')), 'LibreOffice docx conversion'
        except Exception as exc:
            errors.append(str(exc))
    return None, '旧DOC无法提取；提供同名UTF-8 .txt/.docx，或安装antiword/LibreOffice。' + '; '.join(errors)


def documents():
    docs, inventory = {}, []
    for path in sorted(ROOT.iterdir()):
        if not path.is_file() or path.name.startswith('~$') or path.suffix.lower() not in ('.doc', '.docx', '.xlsx'):
            continue
        item = dict(file=path.name, bytes=path.stat().st_size,
                    sha256=hashlib.sha256(path.read_bytes()).hexdigest())
        if path.suffix == '.docx':
            text, method = engine.read_docx(path), 'OOXML paragraphs'
        elif path.suffix == '.doc':
            text, method = legacy_text(path)
        else:
            text, method = None, '历史用例/启动时间读取；截图仅作历史证据，不是可执行用例'
        item.update(extraction=method, chars=len(text) if text else None)
        inventory.append(item)
        if text and not path.name.startswith('测试记录'):
            docs[path.name] = text
    return docs, inventory


def build_suite(cfg):
    docs, inventory = documents()
    excel = next(p for p in ROOT.glob('*.xlsx') if not p.name.startswith('~$'))
    sheets = workbook_rows(excel)
    cases, batches, coverage = [], [], []
    mapping = {}
    for size, key in ((1000, '费用报销'), (3000, '公司考勤'), (5000, '公司规章制度'),
                      (12000, '约12000'), (14000, '约14000'), (15000, '大概15000'), (15750, '大概15986')):
        name = next((name for name in docs if key in name), None)
        mapping[size] = (name, docs[name]) if name else (None, None)
    # Only a real compliance source may satisfy the absent historical 10k source.
    name = next((name for name in docs if '合规管理' in name), None)
    mapping[10000] = (name, docs[name]) if name else (None, None)

    def task(cid, kind, size=None, size2=None, source='', original='', prompt=None, **kw):
        c = dict(id=cid, group=kind, source=source, original_description=original,
                 rule='manual', prompt=prompt or '', **kw)
        if size:
            name, text = mapping[size]
            if not text:
                c['skip'] = f'缺少可读的历史{size}字文档。' + ('旧DOC需转换或安装转换器。' if size == 5000 else '')
            else:
                c.update(document=name, input_document_chars=len(text))
                query = ('费用报销管理制度的目的？' if size == 1000 else '打卡要求有哪些？' if size == 3000 else
                         '员工的行为规范有哪些？' if size == 5000 else '公司合规管理工作的目标是什么？' if size == 10000 else '公司管理制度的适用范围？')
                if kind == 'summary':
                    query = '请给出文档的总结、摘要和大纲。'
                elif kind == 'review':
                    query = '审核本文的内部矛盾、错别字、表达不清、执行风险，引用原文并给出修改建议。'
                elif kind == 'translate':
                    query = '将以下全文翻译为英文，保留结构和条款编号。'
                elif kind == 'page_translate':
                    # Physical pages are not encoded reliably in DOCX paragraph XML.
                    text = text[:1000]
                    query = '将以下文档开头片段翻译为英文，保留结构。'
                    c['adaptation'] = '历史第一页翻译改为前1000字符翻译；没有复现Word分页'
                    c['input_document_chars'] = len(text)
                elif kind == 'diff':
                    other_name, other = mapping[size2]
                    if not other:
                        c['skip'] = f'缺少可读的历史{size2}字对比文档'
                    else:
                        if size == size2:
                            other = text + '\n新增条款：所有申请必须在三个工作日内完成审批。'
                            c['adaptation'] = '原资料没有同长度版本对；用真实原文加已知新增条款构造受控版本'
                            c['rule'] = 'diff'
                        else:
                            c['adaptation'] = '比较现有两份真实制度；历史未提供原始文档对版本，不能视为原版本重放'
                        c.update(second_document=other_name, input_document_chars=len(text)+len(other))
                        c['prompt'] = f'比较A和B，列出新增、删除和修改内容，并引用原文。\n<A>{text}</A>\n<B>{other}</B>'
                if kind != 'diff':
                    c['prompt'] = query + '\n文档原文：\n<document>' + text + '</document>'
                if size >= 14000:
                    c['source_mapping_note'] = '历史报告部分称合规管理制度，按目录同长度管理规章制度映射；并非确认相同文件'
        cases.append(c)
        return c

    base_cases, _ = engine.build_cases(cfg)
    general = {c['id']: c for c in base_cases}
    common = {3:'chat_travel',11:'count_1000',12:'repeat',13:'speech',14:'it_policy',15:'count_5000',16:'poem',17:'js_copy',18:'vue_methods'}
    def src(row):
        return f'{excel.name} / 测试记录0827 / C{row}'
    functional = []
    for row in range(3, 54):
        original = sheets[1].get(f'C{row}', '')
        if not original:
            continue
        if row in common:
            template = general.get(common[row])
            c = dict(template, id=f'h0827_r{row}', source=src(row), original_description=original) if template else dict(id=f'h0827_r{row}',group='翻译',source=src(row),prompt='',rule='manual',skip='缺少英文诗原文')
            cases.append(c)
        else:
            if 4 <= row <= 10:
                kind, size = 'qa', [1000,3000,5000,10000,14000,15000,15750][row-4]
            elif 19 <= row <= 24:
                kind, size = 'summary', [1000,3000,5000,10000,14000,15000][row-19]
            elif 25 <= row <= 31:
                kind, size = 'qa', [1000,3000,5000,10000,14000,15000,15750][row-25]
            elif 32 <= row <= 38:
                kind, size = ('page_translate' if row <= 35 else 'qa'), [1000,3000,5000,10000,14000,15000,15750][row-32]
            elif 39 <= row <= 45:
                kind, size = 'review', [1000,3000,5000,10000,14000,15000,15750][row-39]
            else:
                kind, size = 'diff', [1000,3000,5000,10000,1000,3000,1000,1000][row-46]
            size2 = [1000,3000,5000,10000,10000,10000,14000,15000][row-46] if kind == 'diff' else None
            c = task(f'h0827_r{row}', kind, size, size2, src(row), original)
            if 39 <= row <= 45:
                c['adaptation'] = '历史场景为审核、单元格写获取摘要；本项执行审核，摘要另有对应场景'
        functional.append(c['id'])
        coverage.append(dict(sheet='测试记录0827', row=row, description=original, case_ids=[c['id']]))
    for row, size in zip(range(90,94),(1000,3000,5000,10000)):
        c = task(f'h0827_r{row}','translate',size,source=src(row),original=sheets[1].get(f'C{row}',''))
        functional.append(c['id'])
        coverage.append(dict(sheet='测试记录0827',row=row,description=c['original_description'],case_ids=[c['id']]))
    for first,last,kind in ((54,59,'summary'),(60,65,'qa'),(66,71,'review'),(72,77,'diff'),(78,82,'diff'),(83,86,'diff'),(87,89,'diff')):
        members=[]
        for j,row in enumerate(range(first,last+1)):
            size = ([1000,3000,5000,10000,10000,14000] if kind=='diff' else [1000,1000,3000,3000,5000,5000])[j]
            size2 = [1000,3000,5000,1000,3000,1000][j] if kind=='diff' else None
            c=task(f'h0827_r{row}',kind,size,size2,src(row),sheets[1].get(f'C{row}',''))
            members.append(dict(case_id=c['id'],offset_s=0))
            coverage.append(dict(sheet='测试记录0827',row=row,description=c['original_description'],case_ids=[c['id']]))
        batches.append(dict(id=f'h0827_{kind}_{first}_{last}',members=members,source=f'测试记录0827 C{first}:C{last}',schedule='simultaneous'))
    # September 4 mixed-load experiments; row 4 and row 5 deliberately retained.
    mixes={3:[1000]*11,4:[1000]*9+[3000]*2,5:[1000]*9+[3000]*2,6:[3000]*11,
           7:[1000]*8+[3000]*3,8:[3000]*8,9:[3000]*6,10:[3000]*7,11:[5000]*10}
    for row,sizes in mixes.items():
        members=[]
        for j,size in enumerate(sizes):
            c=task(f'h0904_r{row}_{j+1}','qa',size,source=f'{excel.name} / 0904测试 B{row}',original=sheets[2].get(f'B{row}',''))
            if row==11:
                c['adaptation']='原描述按输出阶段追加；非流式不能观察输出阶段，改为前2个同时、后续每10秒追加。另有历史实际时间表重放。'
            members.append(dict(case_id=c['id'],offset_s=0 if j<2 or row!=11 else (j-1)*10))
        batches.append(dict(id=f'h0904_r{row}',members=members,source=f'0904测试 B{row}',schedule='staggered' if row==11 else 'simultaneous'))
        coverage.append(dict(sheet='0904测试',row=row,description=sheets[2].get(f'B{row}',''),case_ids=[m['case_id'] for m in members]))
    # Historical start timestamps, including repeated labels, are mapped by position.
    schedules=[('5000对话',5000,'qa',15,'BCDEFGH',17,'BCD'),('15000对话',15000,'qa',24,'BCDEFGH',26,'BCD'),
               ('5000审核',5000,'review',33,'BCDEFGH',35,'BCD'),('15000混合',15000,'review',41,'BCDEFGH',43,'BCD')]
    for label,size,kind,r1,cols1,r2,cols2 in schedules:
        values=[(col+str(r),sheets[2].get(col+str(r),'')) for r,cols in ((r1,cols1),(r2,cols2)) for col in cols]
        def seconds(value):
            match=re.match(r'^(\d{1,2})\.(\d{2})(?:-|$)',value)
            if match:
                return int(match[1])*3600+int(match[2])*60
            return round(float(value)*86400)
        times=[seconds(v) for _,v in values]
        members=[]
        # The first cell of the 5000 dialogue table represents two simultaneous requests.
        expanded=[(values[0],times[0])]+list(zip(values,times)) if r1==15 else list(zip(values,times))
        for j,((cell,value),when) in enumerate(expanded):
            this_kind=kind
            size2=None
            this_size=size
            if '差异' in value:
                this_kind,this_size,size2='diff',14000,1000
            elif '翻译' in value:
                this_kind='translate'
            elif '获取总结' in value:
                this_kind='summary'
            elif '对话' in value:
                this_kind='qa'
            c=task(f'schedule_{r1}_{j+1}',this_kind,this_size,size2,source=f'{excel.name} / 0904测试 {cell}',original=value)
            c['adaptation']='按表格单元格位置重放开始时间；原表重复序号按独立请求处理，非复现账号/UI操作'
            members.append(dict(case_id=c['id'],offset_s=when-min(times)))
        batches.append(dict(id=f'schedule_{r1}',members=members,source=f'0904测试 {label}',schedule='staggered'))
        coverage.append(dict(sheet='0904测试',row=r1,description=label+'启动时间表',case_ids=[m['case_id'] for m in members]))
    # Include all readable source documents, including the 12k and 15986 variants.
    for i,(name,text) in enumerate(docs.items(),1):
        for kind,query in (('summary','总结全文并列出大纲。'),('qa','本文的适用范围和员工行为规范是什么？'),('review','审核错别字、内部矛盾和执行风险，引用原文。'),('translate','将全文翻译为英文，保留条款结构。')):
            c=task(f'file_{i}_{kind}',kind,source=name,prompt=query+'\n<document>'+text+'</document>',input_document_chars=len(text))
            functional.append(c['id'])
    c=task('python_sort','代码',source='用户提供示例',prompt='请用 Python 写一段快速排序代码')
    functional.append(c['id'])
    # Every historical text test row must have an explicit mapping, even when blocked.
    expected={r for r in range(3,94) if sheets[1].get(f'C{r}')}
    actual={r['row'] for r in coverage if r['sheet']=='测试记录0827'}
    if expected != actual:
        raise ValueError(f'历史用例覆盖异常：missing={expected-actual}, extra={actual-expected}')
    return dict(cases=cases,batches=batches,functional=functional,coverage=coverage,inventory=inventory,
                limitations=['API层不包含上传、检索、页面、账号登录、硬件利用率及历史截图重建。',
                             '缺失文档明确跳过；同长度差异对、第一页翻译等改编项逐项标注。'])


def run(args):
    cfg=json.loads(args.config.read_text(encoding='utf-8-sig'))
    for name in ('endpoint','max_tokens','temperature','request_timeout_seconds'):
        value=getattr(args,name,None)
        if value is not None:
            cfg[name]=value
    if args.max_tokens is not None:
        cfg['count_5000_max_tokens']=args.max_tokens
    cfg['stream']=args.stream
    cfg['assistant_prefix']=''
    if args.endpoint:
        cfg['model_endpoints']={}
    models=args.models or cfg['models']
    if len(set(models))!=len(models) or any(not re.fullmatch(r'[\w-][\w.-]*',m) for m in models):
        raise ValueError('模型名称为空、重复或包含不支持的目录字符')
    plan=build_suite(cfg)
    by_id={c['id']:c for c in plan['cases']}
    if args.mode=='smoke':
        jobs=[dict(id='smoke',members=[dict(case_id='python_sort',offset_s=0)],schedule='single')]
    else:
        jobs=[]
        if args.mode in ('all','functional'):
            jobs += [dict(id=cid,members=[dict(case_id=cid,offset_s=0)],schedule='single') for cid in plan['functional']]
        if args.mode in ('all','stress'):
            jobs += plan['batches']
    stamp=datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    out=ROOT/'Linux测试结果'/stamp
    out.mkdir(parents=True)
    engine.save_json(out/'plan.json',plan)
    engine.save_json(out/'config.json',dict(config=cfg,models=models,mode=args.mode,stagger_scale=args.stagger_scale,dry_run=args.dry_run))
    engine.save_json(out/'selected_jobs.json',jobs)
    coverage=['# 历史测试覆盖清单','',*plan['limitations'],'',f"计划：{len(plan['cases'])}个请求用例，{len(plan['batches'])}个历史压力批次。",'']
    for row in plan['coverage']:
        blocked=[by_id[i].get('skip') for i in row['case_ids'] if by_id[i].get('skip')]
        adapted=any(by_id[i].get('adaptation') or by_id[i].get('source_mapping_note') for i in row['case_ids'])
        coverage.append(f"- {row['sheet']} 行{row['row']}：{'缺资料/部分缺资料' if blocked else '已映射（有改编）' if adapted else '已映射'}；{row['description']}；用例：{', '.join(row['case_ids'])}")
    (out/'coverage.md').write_text('\n'.join(coverage),encoding='utf-8')
    print(f"输出：{out}\n历史行已逐项映射，缺资料项详见coverage.md。选中{len(jobs)}个任务。",flush=True)
    all_rows,all_stats=[],[]
    aborted=False
    for model in models:
        folder=out/model
        folder.mkdir()
        rows,stats=[],[]
        engine.report(folder,rows,model,stats)
        try:
            for job in jobs:
                members=job['members']
                if args.dry_run:
                    rows.extend(engine.skipped(model,by_id[m['case_id']],by_id[m['case_id']].get('skip','DRY_RUN：未发送')) for m in members)
                    continue
                print(f"[{model}] {job['id']} 请求数={len(members)} 启动跨度={max(m['offset_s'] for m in members)*args.stagger_scale:.1f}s",flush=True)
                start=time.perf_counter()
                batch=[]
                def send(member):
                    case=by_id[member['case_id']]
                    if case.get('skip'):
                        return engine.skipped(model,case,case['skip'])
                    delay=member['offset_s']*args.stagger_scale-(time.perf_counter()-start)
                    if delay>0:
                        time.sleep(delay)
                    r=engine.request_model(cfg,model,case,folder/'requests',case['id'])
                    r['scheduled_offset_s']=member['offset_s']*args.stagger_scale
                    return r
                with ThreadPoolExecutor(max_workers=len(members)) as pool:
                    futures=[pool.submit(send,m) for m in members]
                    for future in as_completed(futures):
                        r=future.result()
                        batch.append(r)
                        rows.append(r)
                        engine.report(folder,rows,model,stats)
                        print(f"  {r['case_id']}: {r['status']}",flush=True)
                wall=time.perf_counter()-start
                attempted=[r for r in batch if r['status']!='SKIPPED']
                returned=[r for r in attempted if r['status'] in ('OK','TRUNCATED')]
                complete=[r for r in attempted if r['status']=='OK']
                stats.append(dict(model=model,batch=job['id'],planned=len(members),attempted=len(attempted),skipped=len(members)-len(attempted),
                                  returned=len(returned),complete=len(complete),truncated=sum(r['status']=='TRUNCATED' for r in batch),
                                  response_rate=len(returned)/len(attempted) if attempted else None,
                                  completion_rate=len(complete)/len(attempted) if attempted else None,wall_seconds=wall,
                                  completed_requests_per_second=len(complete)/wall,
                                  p50_seconds=engine.percentile([r['total_s'] for r in complete],.5),
                                  p95_seconds=engine.percentile([r['total_s'] for r in complete],.95)))
                engine.report(folder,rows,model,stats)
                # Truncation never suppresses subsequent historical scenarios.
                if attempted and all(r.get('request_phase')=='connect' and r['status'] in ('ERROR','TIMEOUT') for r in attempted):
                    remaining=jobs[jobs.index(job)+1:]
                    rows.extend(engine.skipped(model,by_id[m['case_id']],'本模型连接失败，后续请求未发送') for j in remaining for m in j['members'])
                    break
        except KeyboardInterrupt:
            aborted=True
        finally:
            engine.report(folder,rows,model,stats)
            all_rows.extend(rows)
            all_stats.extend(stats)
            engine.report(out,all_rows,'Linux历史全场景测试',all_stats)
        if aborted:
            break
    print(f"完成：{out/'report.html'}",flush=True)
    return 130 if aborted else 2 if not args.dry_run and any(r['status'] in ('ERROR','TIMEOUT','TRUNCATED') or r.get('quality')=='FAIL' for r in all_rows) else 0


def main():
    p=argparse.ArgumentParser(description='Linux历史全场景模型测试，缺失/改编项在coverage.md登记')
    p.add_argument('--config',type=Path,default=engine.HERE/'config.json')
    p.add_argument('--models',nargs='+')
    p.add_argument('--endpoint')
    p.add_argument('--mode',choices=['all','functional','stress','smoke'],default='all')
    p.add_argument('--max-tokens',type=int)
    p.add_argument('--temperature',type=float)
    p.add_argument('--timeout',dest='request_timeout_seconds',type=int)
    p.add_argument('--stagger-scale',type=float,default=1.0,help='历史启动间隔倍数；默认1原时间，0同时启动，改变后不等同历史负载')
    p.add_argument('--stream',action='store_true')
    p.add_argument('--dry-run',action='store_true')
    args=p.parse_args()
    if args.stagger_scale<0 or (args.max_tokens is not None and args.max_tokens<1) or (args.request_timeout_seconds is not None and args.request_timeout_seconds<1):
        p.error('时间倍数须>=0，token/超时须>0')
    if args.temperature is not None and not 0<=args.temperature<=2:
        p.error('temperature须在0到2之间')
    return run(args)


if __name__=='__main__':
    raise SystemExit(main())
