"""Small SQLite transactions; never hold a transaction during network I/O."""

import json
import sqlite3
import time
import uuid
from contextlib import contextmanager
from pathlib import Path


def encode(value):
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), sort_keys=True)


def identifier():
    return uuid.uuid4().hex


class Store:
    def __init__(self, root: Path):
        self.root = root
        root.mkdir(parents=True, exist_ok=True)
        self.path = root / "creator.sqlite3"
        with self.connect() as db:
            db.execute("PRAGMA journal_mode=WAL")
            version = db.execute("PRAGMA user_version").fetchone()[0]
            if version > 1:
                raise ValueError("数据库版本高于当前程序版本")
            if version == 0:
                # Existing databases must be backed up before future migrations.
                db.executescript("""
                    CREATE TABLE users (
                        id TEXT PRIMARY KEY, username TEXT NOT NULL UNIQUE,
                        display_name TEXT NOT NULL, password_hash TEXT NOT NULL,
                        role TEXT NOT NULL, active INTEGER NOT NULL DEFAULT 1,
                        must_change INTEGER NOT NULL DEFAULT 1, created_at REAL NOT NULL
                    );
                    CREATE TABLE sessions (
                        token_hash TEXT PRIMARY KEY, user_id TEXT NOT NULL REFERENCES users(id),
                        csrf TEXT NOT NULL, expires_at REAL NOT NULL
                    );
                    CREATE TABLE records (
                        id TEXT PRIMARY KEY, owner TEXT NOT NULL, kind TEXT NOT NULL,
                        name TEXT NOT NULL, data TEXT NOT NULL, secret TEXT NOT NULL DEFAULT '',
                        version INTEGER NOT NULL DEFAULT 1, updated_at REAL NOT NULL
                    );
                    CREATE INDEX records_owner ON records(owner, kind);
                    CREATE TABLE jobs (
                        id TEXT PRIMARY KEY, owner TEXT NOT NULL REFERENCES users(id),
                        kind TEXT NOT NULL, status TEXT NOT NULL, payload TEXT NOT NULL,
                        snapshot TEXT NOT NULL DEFAULT '', result TEXT NOT NULL DEFAULT '{}',
                        error TEXT NOT NULL DEFAULT '', idem TEXT, digest TEXT,
                        created_at REAL NOT NULL, updated_at REAL NOT NULL,
                        UNIQUE(owner, idem)
                    );
                    CREATE INDEX jobs_owner ON jobs(owner, created_at);
                    CREATE TABLE items (
                        job_id TEXT NOT NULL REFERENCES jobs(id), owner TEXT NOT NULL,
                        item_id TEXT NOT NULL, state TEXT NOT NULL, data TEXT NOT NULL,
                        issue_key TEXT NOT NULL DEFAULT '', issue_url TEXT NOT NULL DEFAULT '',
                        error TEXT NOT NULL DEFAULT '', attempts INTEGER NOT NULL DEFAULT 0,
                        PRIMARY KEY(job_id, item_id)
                    );
                    CREATE INDEX items_identity ON items(owner, item_id, state);
                    CREATE TABLE login_failures (bucket TEXT NOT NULL, occurred_at REAL NOT NULL);
                    CREATE INDEX login_bucket ON login_failures(bucket, occurred_at);
                    CREATE TABLE settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
                    PRAGMA user_version=1;
                """)

    @contextmanager
    def connect(self):
        db = sqlite3.connect(self.path, timeout=5)
        db.row_factory = sqlite3.Row
        db.execute("PRAGMA foreign_keys=ON")
        try:
            with db:
                yield db
        finally:
            db.close()

    def record(self, owner, kind, record_id):
        with self.connect() as db:
            row = db.execute("SELECT * FROM records WHERE id=? AND owner=? AND kind=?",
                             (record_id, owner, kind)).fetchone()
        return self.unpack(row) if row else None

    @staticmethod
    def unpack(row):
        result = dict(row)
        result["data"] = json.loads(result["data"])
        return result

    def records(self, owner, kind):
        with self.connect() as db:
            return [self.unpack(row) for row in db.execute(
                "SELECT * FROM records WHERE owner=? AND kind=? ORDER BY updated_at DESC", (owner, kind))]

    def save(self, owner, kind, name, data, record_id=None, secret="", version=None):
        record_id = record_id or identifier()
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            old = db.execute("SELECT * FROM records WHERE id=?", (record_id,)).fetchone()
            if old:
                if old["owner"] != owner or old["kind"] != kind:
                    raise ValueError("配置不存在")
                if version != old["version"]:
                    raise ValueError("内容已更新，请重新加载后保存")
                db.execute("UPDATE records SET name=?,data=?,secret=?,version=version+1,updated_at=? WHERE id=?",
                           (name, encode(data), secret, time.time(), record_id))
            else:
                db.execute("INSERT INTO records(id,owner,kind,name,data,secret,updated_at) VALUES(?,?,?,?,?,?,?)",
                           (record_id, owner, kind, name, encode(data), secret, time.time()))
        return self.record(owner, kind, record_id)

    def job(self, owner, job_id):
        with self.connect() as db:
            row = db.execute("SELECT * FROM jobs WHERE id=? AND owner=?", (job_id, owner)).fetchone()
            if not row:
                return None
            value = dict(row)
            value["payload"] = json.loads(value["payload"])
            value["result"] = json.loads(value["result"])
            value["items"] = [dict(item) for item in db.execute(
                "SELECT * FROM items WHERE job_id=? ORDER BY rowid", (job_id,))]
            for item in value["items"]:
                item["data"] = json.loads(item["data"])
            return value

    def update_job(self, job_id, status, result=None, error=""):
        with self.connect() as db:
            db.execute("UPDATE jobs SET status=?,result=?,error=?,updated_at=? WHERE id=?",
                       (status, encode(result or {}), error, time.time(), job_id))

    def item_state(self, job_id, item_id, state, error="", key="", url=""):
        with self.connect() as db:
            db.execute("UPDATE items SET state=?,error=?,issue_key=?,issue_url=?,attempts=attempts+? "
                       "WHERE job_id=? AND item_id=?",
                       (state, error, key, url, int(state == "creating"), job_id, item_id))
            db.execute("UPDATE jobs SET updated_at=? WHERE id=?", (time.time(), job_id))

    def recover(self):
        with self.connect() as db:
            db.execute("UPDATE items SET state='unknown',error='服务重启，创建结果待核实' WHERE state='creating'")
            db.execute("UPDATE items SET state='failed',error='服务重启，尚未发送到 Jira' WHERE state='queued'")
            db.execute("UPDATE jobs SET status='interrupted',error='服务重启，任务已中断',updated_at=? "
                       "WHERE status IN ('queued','running')", (time.time(),))

    def seed_defaults(self, owner):
        """One-time account defaults; never recreate a user's intentional deletions."""
        defaults = {"project_key": "KECC", "issuetype_id": "", "priority_id": "",
                    "assignee": None, "due_date": None, "components": [], "labels": []}
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")

            def insert_record(record_owner, kind, name, data):
                record_id = identifier()
                db.execute("INSERT INTO records(id,owner,kind,name,data,updated_at) VALUES(?,?,?,?,?,?)",
                           (record_id, record_owner, kind, name, encode(data), time.time()))
                return record_id

            public_marker = "seed:public-template:v1"
            if not db.execute("SELECT 1 FROM settings WHERE key=?", (public_marker,)).fetchone():
                if not db.execute("SELECT 1 FROM records WHERE owner='public' AND kind='template'").fetchone():
                    insert_record("public", "template", "企业 Jira + DeepSeek 默认模板", {
                        "jira_base_url": "http://jira.kingsware.cn",
                        "llm_base_url": "https://api.deepseek.com/v1", "model": "deepseek-chat",
                        "output_mode": "json_object", "defaults": defaults, "members": {}, "mappings": [],
                    })
                db.execute("INSERT INTO settings VALUES(?,?)", (public_marker, "done"))

            marker = "seed:account-defaults:v1:" + owner
            if db.execute("SELECT 1 FROM settings WHERE key=?", (marker,)).fetchone():
                return
            existing = db.execute("SELECT id,data FROM records WHERE owner=? AND kind='connection' ORDER BY updated_at",
                                  (owner,)).fetchall()
            connections = {}
            for row in existing:
                connections.setdefault(json.loads(row["data"])["kind"], row["id"])
            for kind, name, url in (("jira", "企业 Jira", "http://jira.kingsware.cn"),
                                    ("llm", "DeepSeek", "https://api.deepseek.com/v1")):
                if kind not in connections:
                    connections[kind] = insert_record(owner, "connection", name, {
                        "kind": kind, "base_url": url, "username": "",
                        "model": "deepseek-chat" if kind == "llm" else "", "output_mode": "json_object",
                    })
            if not db.execute("SELECT 1 FROM records WHERE owner=? AND kind='profile'", (owner,)).fetchone():
                profile_id = insert_record(owner, "profile", "默认建单方案", {
                    "jira_connection_id": connections["jira"], "llm_connection_id": connections["llm"],
                    "defaults": defaults, "members": {}, "mappings": [],
                })
                db.execute("INSERT OR IGNORE INTO settings VALUES(?,?)", ("default:" + owner, profile_id))
            db.execute("INSERT INTO settings VALUES(?,?)", (marker, "done"))
