"""串联读取、提取、校验、保存和导出；窗口与此处共用真实入口。"""
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
import hashlib
import threading

from .文档读取 import read_document, chunk_blocks
from .字段配置 import FIELDS
from .模型调用 import Client, ConfigurationError, RangeTooLarge, Cancelled
from .结果校验 import merge_records, validate_records
from .任务存储 import Store
from .模板输出 import export_results

SUPPORTED={'.doc','.docx','.pdf','.png','.jpg','.jpeg','.tif','.tiff'}


def file_hash(path):
    digest=hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda:stream.read(1024*1024),b''):
            digest.update(block)
    return digest.hexdigest()


def collect_files(paths,excluded=()):
    excluded=[Path(p).resolve() for p in excluded]
    found=[];seen=set()
    for entry in paths:
        entry=Path(entry).resolve()
        candidates=sorted(entry.rglob('*')) if entry.is_dir() else [entry]
        for p in candidates:
            if not p.is_file() or p.suffix.lower() not in SUPPORTED or p.name.startswith('~$'):
                continue
            if any(p.is_relative_to(q) for q in excluded):
                continue
            if any(n in {'.local','.git','__pycache__'} for n in p.parts):
                continue
            key=str(p).casefold()
            if key not in seen:
                seen.add(key);found.append(p)
    return found


class Pipeline:
    def __init__(self,settings,kind,output_dir,batch_dir,template=None,instructions='',on_event=None,client=None):
        if kind not in FIELDS:
            raise ValueError('请选择合同、会议纪要、发票或会计凭证。')
        self.settings=dict(settings)
        self.kind=kind
        self.output_dir=Path(output_dir)
        self.batch_dir=Path(batch_dir)
        self.template=template
        self.instructions=instructions
        self.on_event=on_event or (lambda event:None)
        self.client=client

    def emit(self,message,**values):
        self.on_event({'message':message,**values})

    def _extract_chunk(self,blocks,client,fields,cancel,on_attempt,depth=0):
        if cancel.is_set():
            raise Cancelled('已停止后续处理。')
        try:
            return client.extract(blocks,self.kind,fields=fields,instructions=self.instructions,cancel_event=cancel,on_attempt=on_attempt)
        except RangeTooLarge:
            if depth>=6:
                raise ValueError('当前片段超过服务限制，多次拆分仍无法处理，需要人工核对。')
            if len(blocks)>1:
                mid=len(blocks)//2;parts=[blocks[:mid],blocks[mid:]]
            elif blocks and blocks[0].get('kind')=='text' and len(blocks[0].get('text',''))>400:
                text=blocks[0]['text'];mid=len(text)//2
                parts=[[dict(blocks[0],text=text[:mid])],[dict(blocks[0],text=text[mid:])]]
            else:
                raise ValueError('单个页面或最小片段超过服务限制，请在核对清单中检查该页。')
            results=[self._extract_chunk(p,client,fields,cancel,on_attempt,depth+1) for p in parts]
            return {'records':[r for x in results for r in x.get('records',[])], 'complete':all(x.get('complete') for x in results), 'unprocessed':[r for x in results for r in x.get('unprocessed',[])], 'raw':'', 'usage':{}}

    def _checked_rows(self,checkpoint,complete,warnings):
        raw=[dict(r) for key,value in sorted(checkpoint.get('chunks',{}).items(),key=lambda x:int(x[0])) for r in value.get('records',[])]
        rows=validate_records(merge_records(raw,self.kind),self.kind,complete=complete)
        if warnings:
            for row in rows:
                row['status']='review'
                row['reasons']=list(dict.fromkeys(row.get('reasons',[])+warnings))
        return rows

    def run(self,paths,resume=False,retry_only=False,cancel_event=None):
        cancel=cancel_event or threading.Event()
        self.batch_dir.mkdir(parents=True,exist_ok=True)
        store=Store(self.batch_dir)
        output=None;cancelled=False;paused=False
        if self.template:
            saved=store.get_config()
            if saved and saved.get('template'):
                if self.template['hash']!=saved['template']['hash']:
                    store.close()
                    raise ValueError('模板与本批次不同，请建立新批次。')
                self.template=saved['template']
            else:
                original=Path(self.template['path'])
                if file_hash(original)!=self.template['hash']:
                    store.close()
                    raise ValueError('模板选择后已改变，请重新选择。')
                snapshot=self.batch_dir/'模板.xlsx'
                if not snapshot.exists():
                    with snapshot.open('xb') as stream:
                        stream.write(original.read_bytes())
                if file_hash(snapshot)!=self.template['hash']:
                    store.close()
                    raise ValueError('模板副本校验失败，请建立新批次。')
                self.template={**self.template,'path':str(snapshot)}
        fields=list(self.template['fields']) if self.template else list(FIELDS[self.kind])
        safe={k:v for k,v in self.settings.items() if k not in {'api_key','remember_key'}}
        config={'settings':safe,'kind':self.kind,'fields':fields,'instructions':self.instructions,'template':self.template,'output_dir':str(self.output_dir)}
        try:
            store.configure(config)
            sources=collect_files(paths,excluded=[self.batch_dir])
            saved_sources=store.sources()
            if resume or retry_only:
                sources=list(dict.fromkeys([Path(s['path']) for s in saved_sources]+sources))
            if not sources and not cancel.is_set():
                raise ValueError('没有找到可处理的Word、图片或PDF文件。')
            self.client=self.client or Client(self.settings)
            # 先登记完整输入清单，停止后仍能找到尚未开始的文件。
            saved_by_path={str(Path(s['path']).resolve()).casefold():s for s in saved_sources}
            indexed=[];seen_content=set()
            for path in sources:
                old=saved_by_path.get(str(path.resolve()).casefold())
                try:
                    source_id=file_hash(path)
                except OSError:
                    source_id=old['source_id'] if old else hashlib.sha256(str(path).encode()).hexdigest()
                if old and old['source_id']!=source_id:
                    raise ValueError('原文件内容已改变，请建立新批次，原批次结果继续保留。')
                if source_id not in seen_content:
                    indexed.append((source_id,path));seen_content.add(source_id)
                else:
                    self.emit('内容相同的重复文件已合并：'+path.name)
            registered={s['source_id'] for s in saved_sources}
            for source_id,path in indexed:
                if source_id not in registered:
                    store.replace_source(source_id,str(path),[],complete=False)
            review_paths={r.get('source_file') for r in store.records() if r.get('status')=='review'}
            for number,(source_id,path) in enumerate(indexed,1):
                if cancel.is_set():
                    cancelled=True;break
                if retry_only and str(path) not in review_paths and store.source_done(source_id):
                    continue
                if store.source_done(source_id) and not retry_only:
                    self.emit('已保存，跳过重复提取：'+path.name,done=number,total=len(sources));continue
                self.emit('读取：'+path.name,phase='读取文档',done=number-1,total=len(sources))
                try:
                    blocks=read_document(path,self.batch_dir/'转换文件'/source_id)
                    warnings=[b.get('text','存在未读取区域') for b in blocks if b.get('kind')=='warning']
                    processable=[b for b in blocks if b.get('kind') in {'text','image'}]
                    if not processable:
                        raise ValueError('文件没有可识别的文字或图像。')
                    chunks=chunk_blocks(processable)
                    checkpoint=store.get_checkpoint(source_id) or {'chunks':{},'total':len(chunks)}
                    if checkpoint.get('total')!=len(chunks):
                        checkpoint={'chunks':{},'total':len(chunks)}
                    if retry_only:
                        # 核对项可能由跨片段关联引起，重试该文件，其他正常文件不会重问。
                        checkpoint={'chunks':{},'total':len(chunks)}
                    todo=[i for i in range(len(chunks)) if not checkpoint['chunks'].get(str(i),{}).get('complete')]
                    failures=[]
                    def save_attempt(attempt,identifier=source_id):
                        store.save_attempt(identifier,attempt.get('raw',''),attempt.get('usage'),error=attempt.get('error'))
                    with ThreadPoolExecutor(max_workers=2) as pool:
                        future_map={pool.submit(self._extract_chunk,chunks[i],self.client,fields,cancel,save_attempt):i for i in todo}
                        for future in as_completed(future_map):
                            index=future_map[future]
                            try:
                                answer=future.result()
                                positions={b['position'] for b in chunks[index]}
                                for record in answer.get('records',[]):
                                    if not record.get('positions') or any(p not in positions for p in record.get('positions',[])):
                                        record.setdefault('issues',[]).append('回答的来源位置缺失或与该片段不一致。')
                                if not answer.get('records'):
                                    answer['complete']=False
                                    answer.setdefault('unprocessed',[]).append('该片段未提取到业务记录，需确认是否为空白或遗漏。')
                                checkpoint['chunks'][str(index)]={k:answer.get(k,[] if k!='complete' else False) for k in ['records','complete','unprocessed']}
                                complete=len(checkpoint['chunks'])==len(chunks) and all(v.get('complete') and not v.get('unprocessed') for v in checkpoint['chunks'].values())
                                current_warnings=warnings+[str(x) for v in checkpoint['chunks'].values() for x in v.get('unprocessed',[])]
                                rows=self._checked_rows(checkpoint,complete,current_warnings)
                                store.replace_source(source_id,str(path),rows,complete=complete and not current_warnings,checkpoint=checkpoint)
                                self.emit(f'{path.name}：已保存 {len(checkpoint["chunks"])} / {len(chunks)} 个片段',phase='提取并保存')
                            except Cancelled:
                                cancelled=True;cancel.set()
                            except ConfigurationError:
                                paused=True;cancel.set();failures.append('连接配置未通过，请检查地址、密钥和模型名称后恢复。')
                            except Exception as error:
                                failures.append(str(error))
                    complete=len(checkpoint['chunks'])==len(chunks) and all(v.get('complete') and not v.get('unprocessed') for v in checkpoint['chunks'].values())
                    all_warnings=list(dict.fromkeys(warnings+failures+[str(x) for v in checkpoint['chunks'].values() for x in v.get('unprocessed',[])]))
                    if cancelled or paused:
                        all_warnings.append('该文件尚未完整处理，可恢复继续。')
                    rows=self._checked_rows(checkpoint,complete,all_warnings)
                    if not rows:
                        rows=[{'document_id':'未完成文件','fields':{},'positions':[b['position'] for b in blocks[:1]],'status':'review','reasons':all_warnings or ['未提取到可确认数据。']}]
                    store.replace_source(source_id,str(path),rows,complete=complete and not all_warnings,checkpoint=checkpoint)
                except (ConfigurationError,Cancelled):
                    paused=True;cancel.set()
                except Exception as error:
                    prior=[r for r in store.records() if r.get('source_file')==str(path)]
                    if not prior:
                        prior=[{'document_id':'无法读取','fields':{},'positions':[],'status':'review','reasons':[]}]
                    for row in prior:
                        row['status']='review'
                        row['reasons']=list(dict.fromkeys(row.get('reasons',[])+[str(error)]))
                    store.replace_source(source_id,str(path),prior,complete=False,checkpoint=store.get_checkpoint(source_id))
                    self.emit(path.name+'：已记录需要核对的原因。')
                self.emit('完成当前文件：'+path.name,done=number,total=len(sources))
                if paused or cancelled:
                    break
            rows=store.records()
            if self.kind=='发票':
                by_number={}
                for row in rows:
                    number=row.get('fields',{}).get('发票号码')
                    if number:
                        by_number.setdefault(str(number),[]).append(row)
                for group in by_number.values():
                    if len({r['source_file'] for r in group})>1:
                        for row in group:
                            row['status']='review'
                            row['reasons']=list(dict.fromkeys(row.get('reasons',[])+['不同文件出现相同票号，需核对是否重复票据。']))
                for source in store.sources():
                    own=[r for r in rows if r['source_file']==source['path']]
                    store.replace_source(source['source_id'],source['path'],own,complete=source['complete'],checkpoint=store.get_checkpoint(source['source_id']))
            if rows:
                self.emit('生成Excel结果。',phase='导出')
                try:
                    output=export_results(rows,self.output_dir,self.batch_dir.name,fields,template=self.template)
                    store.set_export(str(output))
                except Exception as error:
                    self.emit('结果已保存，但Excel尚未导出：'+str(error),phase='导出未完成')
            normal=sum(r.get('status')=='normal' for r in rows)
            review=len(rows)-normal
            complete=not cancelled and not paused and not cancel.is_set() and all(s.get('complete') for s in store.sources()) and bool(output)
            return {'output':str(output) if output else None,'normal':normal,'review':review,'rows':rows,'complete':complete,'cancelled':cancelled or (cancel.is_set() and not paused),'paused':paused,'batch_dir':str(self.batch_dir)}
        finally:
            store.close()