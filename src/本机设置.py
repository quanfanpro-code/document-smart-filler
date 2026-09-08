"""只在本机保存连接设置，公开版本不包含私人服务信息。"""
from pathlib import Path
from datetime import datetime
import json
import os
import shutil
import uuid

DEFAULTS = {'base_url':'', 'api_key':'', 'model':'Qwen3.8-27B', 'remember_key':False}


def load_settings(root):
    path=Path(root)/'.local'/'设置.json'
    settings=DEFAULTS.copy()
    if path.exists():
        try:
            data=json.loads(path.read_text(encoding='utf-8-sig'))
            settings.update({k:data[k] for k in DEFAULTS if k in data})
        except (ValueError,OSError) as error:
            raise ValueError('本机设置无法读取，请保留原文件后重新填写设置。') from error
    if not settings['remember_key']:
        settings['api_key']=''
    return settings


def save_settings(root,settings):
    folder=Path(root)/'.local';folder.mkdir(parents=True,exist_ok=True)
    path=folder/'设置.json'
    data={k:settings.get(k,v) for k,v in DEFAULTS.items()}
    if not data['remember_key']:
        data['api_key']=''
    text=json.dumps(data,ensure_ascii=False,indent=2)
    if path.exists() and path.read_text(encoding='utf-8-sig')==text:
        return
    if path.exists():
        backup=Path.home()/'BackUp'/('文档智能填表系统_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f'))/'本机设置'/'设置.json'
        backup.parent.mkdir(parents=True,exist_ok=True)
        shutil.copy2(path,backup)
        if backup.read_bytes()!=path.read_bytes():
            raise OSError('本机设置备份校验失败，未修改原文件。')
    temp=folder/('设置_'+uuid.uuid4().hex+'.tmp')
    temp.write_text(text,encoding='utf-8-sig')
    os.replace(temp,path)