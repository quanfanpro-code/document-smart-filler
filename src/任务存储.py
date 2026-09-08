"""以单个来源为事务单位保存结果；Excel不参与恢复状态判断。"""
from __future__ import annotations

import json
import os
import sqlite3
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from uuid import uuid4


def _encode(value):
    if isinstance(value, Decimal):
        return {"__decimal__": str(value)}
    if isinstance(value, Path):
        return str(value)
    raise TypeError(f"无法保存的数据类型：{type(value).__name__}")


def _decode(value):
    if set(value) == {"__decimal__"}:
        return Decimal(value["__decimal__"])
    return value


def _dump(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, default=_encode, allow_nan=False)


def _load(value):
    return json.loads(value, object_hook=_decode)


def _without_key(value):
    if isinstance(value, dict):
        return {key: _without_key(item) for key, item in value.items()
                if str(key).lower() not in {"api_key", "apikey", "authorization"}}
    if isinstance(value, (list, tuple)):
        return [_without_key(item) for item in value]
    return value


class Store:
    """由流程协调线程串行调用；同一来源的结果和片段进度同时提交。"""

    def __init__(self, batch_dir):
        self.batch_dir = Path(batch_dir)
        self.batch_dir.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(self.batch_dir / "任务.sqlite3")
        self.db.execute("PRAGMA synchronous=FULL")
        self.db.executescript("""
            CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS sources (
                source_id TEXT PRIMARY KEY, path TEXT NOT NULL,
                complete INTEGER NOT NULL, ordinal INTEGER NOT NULL,
                records_json TEXT NOT NULL, checkpoint_json TEXT
            );
        """)
        self.db.commit()

    def configure(self, config):
        serialized = _dump(_without_key(config))
        with self.db:
            current = self.db.execute("SELECT value FROM meta WHERE key='config'").fetchone()
            if current is not None and current[0] != serialized:
                raise ValueError("批次配置与保存时不同，请恢复原配置或建立新批次")
            self.db.execute("INSERT OR IGNORE INTO meta(key,value) VALUES('config',?)", (serialized,))

    def get_config(self):
        row = self.db.execute("SELECT value FROM meta WHERE key='config'").fetchone()
        return _load(row[0]) if row else None

    def replace_source(self, source_id, path, records, complete=True, checkpoint=None):
        if not source_id:
            raise ValueError("来源标识不能为空")
        serialized = _dump(list(records))
        checkpoint_json = _dump(checkpoint) if checkpoint is not None else None
        with self.db:
            self.db.execute("""
                INSERT INTO sources(source_id,path,complete,ordinal,records_json,checkpoint_json)
                VALUES(?,?,?,(SELECT COALESCE(MAX(ordinal),-1)+1 FROM sources),?,?)
                ON CONFLICT(source_id) DO UPDATE SET path=excluded.path,
                    complete=excluded.complete, records_json=excluded.records_json,
                    checkpoint_json=excluded.checkpoint_json
            """, (str(source_id), str(path), bool(complete), serialized, checkpoint_json))

    def source_done(self, source_id):
        row = self.db.execute("SELECT complete FROM sources WHERE source_id=?", (str(source_id),)).fetchone()
        return bool(row[0]) if row else False

    def get_checkpoint(self, source_id):
        row = self.db.execute("SELECT checkpoint_json FROM sources WHERE source_id=?", (str(source_id),)).fetchone()
        return _load(row[0]) if row and row[0] else None

    def records(self):
        result = []
        for path, serialized in self.db.execute("SELECT path,records_json FROM sources ORDER BY ordinal,source_id"):
            for record in _load(serialized):
                result.append({**record, "source_file": path})
        return result

    def sources(self):
        return [{"source_id": source_id, "path": path, "complete": bool(complete)}
                for source_id, path, complete in self.db.execute(
                    "SELECT source_id,path,complete FROM sources ORDER BY ordinal,source_id")]

    def save_attempt(self, source_id, raw, usage=None, error=None):
        folder = self.batch_dir / "原始回答"
        folder.mkdir(exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        path = folder / f"{stamp}_{uuid4().hex}.json"
        payload = {"source_id": str(source_id), "raw": raw, "usage": usage, "error": error}
        with path.open("x", encoding="utf-8-sig") as stream:
            stream.write(_dump(payload))
            stream.flush()
            os.fsync(stream.fileno())
        return path

    def set_export(self, path):
        with self.db:
            self.db.execute("INSERT INTO meta(key,value) VALUES('export',?) "
                            "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (_dump(str(path)),))

    def get_export(self):
        row = self.db.execute("SELECT value FROM meta WHERE key='export'").fetchone()
        return _load(row[0]) if row else None

    def close(self):
        self.db.close()