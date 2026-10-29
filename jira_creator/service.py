import asyncio
from concurrent.futures import ProcessPoolExecutor
import hashlib
import json
import multiprocessing
import os
import re
from pathlib import Path
import shutil
import sqlite3
import time

from argon2 import PasswordHasher
from cryptography.fernet import Fernet, InvalidToken
from pydantic import ValidationError

from .clients import Jira, RemoteError, extract
from .config import ConfigurationError
from .documents import parse_sources
from .progress import report_progress
from .security import checked_url, normalize_url
from .storage import Store, encode, identifier


class Service:
    def __init__(self):
        self.root = Path(os.environ.get("JIRA_CREATOR_DATA_DIR", "data/jira-creator")).resolve()
        self.root.mkdir(parents=True, exist_ok=True)
        self.lock_file = (self.root / "instance.lock").open("a+b")
        try:
            if os.name == "nt":
                import msvcrt
                self.lock_file.seek(0)
                if not self.lock_file.read(1):
                    self.lock_file.write(b"0")
                    self.lock_file.flush()
                self.lock_file.seek(0)
                try:
                    msvcrt.locking(self.lock_file.fileno(), msvcrt.LK_NBLCK, 1)
                except OSError as exc:
                    raise ConfigurationError("无法取得 Jira 批量建单实例锁，请停止使用同一数据目录的其他服务，并检查锁文件权限") from exc
            else:
                import fcntl
                try:
                    fcntl.flock(self.lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                except BlockingIOError as exc:
                    raise ConfigurationError("Jira 批量建单数据目录已被其他服务占用，请保持单实例运行") from exc
            self._initialize()
        except Exception:
            self.lock_file.close()
            raise

    def _initialize(self):
        key = os.environ.get("JIRA_CREATOR_MASTER_KEY", "")
        if not key:
            raise ConfigurationError("缺少 JIRA_CREATOR_MASTER_KEY，请在项目根目录 .env 中填写加密主密钥")
        try:
            self.crypto = Fernet(key.encode())
        except (ValueError, TypeError) as exc:
            raise ConfigurationError("JIRA_CREATOR_MASTER_KEY 格式无效，请填写 Fernet.generate_key() 生成的完整密钥") from exc
        self.store = Store(self.root)
        with self.store.connect() as db:
            sentinel = db.execute("SELECT value FROM settings WHERE key='key_check'").fetchone()
            if sentinel:
                try:
                    valid_key = self.crypto.decrypt(sentinel["value"].encode()) == b"jira-creator-v1"
                except InvalidToken:
                    valid_key = False
                if not valid_key:
                    raise ConfigurationError("主密钥与现有数据库不匹配，请恢复创建数据库时使用的 JIRA_CREATOR_MASTER_KEY，不要重新生成")
            else:
                db.execute("INSERT INTO settings VALUES('key_check',?)", (self.crypto.encrypt(b"jira-creator-v1").decode(),))
        self.allowed_origins = set()
        allowed_origins = os.environ.get("JIRA_CREATOR_ALLOWED_ORIGINS", "").strip() or "http://jira.kingsware.cn,https://api.deepseek.com"
        for value in allowed_origins.split(","):
            if value.strip():
                try:
                    _, origin = normalize_url(value)
                except ValueError as exc:
                    raise ConfigurationError("JIRA_CREATOR_ALLOWED_ORIGINS 格式无效，请填写逗号分隔的 HTTP(S) 服务来源") from exc
                self.allowed_origins.add(origin)
        self.public_origin = os.environ.get("JIRA_CREATOR_PUBLIC_ORIGIN", "").rstrip("/")
        if self.public_origin:
            try:
                _, self.public_origin = normalize_url(self.public_origin)
            except ValueError as exc:
                raise ConfigurationError("JIRA_CREATOR_PUBLIC_ORIGIN 格式无效，本机运行时可留空") from exc
        self.secure_cookie = os.environ.get("JIRA_CREATOR_COOKIE_SECURE", "false").lower() == "true"
        self.hasher = PasswordHasher()
        self.dummy_hash = self.hasher.hash(identifier())
        with self.store.connect() as db:
            if not db.execute("SELECT 1 FROM users LIMIT 1").fetchone():
                username = os.environ.get("JIRA_CREATOR_ADMIN_USERNAME", "").strip()
                password = os.environ.get("JIRA_CREATOR_ADMIN_PASSWORD", "")
                from .models import UserCreate
                try:
                    initial = UserCreate(username=username, password=password, display_name="管理员", role="admin")
                except ValidationError as exc:
                    fields = {error["loc"][0] for error in exc.errors()}
                    messages = []
                    if "username" in fields:
                        messages.append("JIRA_CREATOR_ADMIN_USERNAME 须为 2 至 80 位字母、数字或 _.@- 字符")
                    if "password" in fields:
                        messages.append("JIRA_CREATOR_ADMIN_PASSWORD 配置无法读取")
                    raise ConfigurationError("初始管理员配置无效：" + "；".join(messages)) from exc
                db.execute("INSERT INTO users(id,username,display_name,password_hash,role,created_at,must_change) VALUES(?,?,?,?,?,?,0)",
                           (identifier(), initial.username.casefold(), initial.display_name,
                            self.hasher.hash(initial.password), "admin", time.time()))
        with self.store.connect() as db:
            owners = [row[0] for row in db.execute("SELECT id FROM users")]
        for owner in owners:
            self.store.seed_defaults(owner)
        self.store.recover()
        self.upload_root = self.root / "sources"
        self.upload_root.mkdir(exist_ok=True)
        self.ocr_pool = ProcessPoolExecutor(max_workers=1, mp_context=multiprocessing.get_context("spawn"))
        self.ocr_slots = asyncio.Semaphore(1)
        self.ai_slots = asyncio.Semaphore(2)
        self.jira_slots = asyncio.Semaphore(3)
        self.tasks = set()
        self.cleanup_task = asyncio.create_task(self.cleanup_loop())

    def encrypt(self, value):
        return self.crypto.encrypt(value.encode()).decode() if value else ""

    def decrypt(self, value):
        return self.crypto.decrypt(value.encode()).decode() if value else ""

    def connection(self, owner, connection_id, kind):
        record = self.store.record(owner, "connection", connection_id)
        if not record or record["data"]["kind"] != kind:
            raise ValueError(f"请选择有效的 {kind} 连接")
        if not record["secret"]:
            raise ValueError(f"请先保存 {kind} 连接的凭据")
        checked_url(self, record["data"]["base_url"])
        return {**record["data"], "secret": self.decrypt(record["secret"]),
                "id": record["id"], "version": record["version"]}

    def snapshot(self, owner, profile_id, needs_model=False):
        record = self.store.record(owner, "profile", profile_id)
        if not record:
            raise ValueError("建单方案不存在")
        profile = record["data"]
        result = {"profile": profile, "profile_name": record["name"], "profile_version": record["version"]}
        if needs_model:
            result["llm"] = self.connection(owner, profile.get("llm_connection_id"), "llm")
        else:
            result["jira"] = self.connection(owner, profile.get("jira_connection_id"), "jira")
        return result

    def submit(self, owner, kind, payload, snapshot=None, idem=None):
        fingerprint = hashlib.sha256(encode(payload).encode()).hexdigest()
        job_id = identifier()
        with self.store.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            if idem:
                existing = db.execute("SELECT id,digest FROM jobs WHERE owner=? AND idem=?", (owner, idem)).fetchone()
                if existing:
                    if existing["digest"] != fingerprint:
                        raise ValueError("同一提交标识的内容发生变化，请重新确认后提交")
                    return existing["id"], False
            if db.execute("SELECT COUNT(*) FROM jobs WHERE status IN ('queued','running')").fetchone()[0] >= 50:
                raise ValueError("后台任务队列已满，请稍后重试")
            if kind == "batch":
                if db.execute("SELECT 1 FROM jobs WHERE owner=? AND kind='batch' AND status IN ('queued','running')", (owner,)).fetchone():
                    raise ValueError("你已有正在执行的创建批次，请等待完成")
                ids = [item["item_id"] for item in payload["issues"]]
                if len(ids) != len(set(ids)):
                    raise ValueError("事项标识重复，请重新加载草稿")
                for item_id in ids:
                    if db.execute("SELECT 1 FROM items WHERE owner=? AND item_id=? "
                                  "AND state IN ('queued','creating','success','unknown')", (owner, item_id)).fetchone():
                        raise ValueError("包含已成功、待核实或正在处理的事项，请从历史记录处理")
            now = time.time()
            db.execute("INSERT INTO jobs(id,owner,kind,status,payload,snapshot,idem,digest,created_at,updated_at) "
                       "VALUES(?,?,?,'queued',?,?,?,?,?,?)",
                       (job_id, owner, kind, encode(payload), self.encrypt(encode(snapshot or {})), idem, fingerprint, now, now))
            if kind == "batch":
                for item in payload["issues"]:
                    db.execute("INSERT INTO items(job_id,owner,item_id,state,data) VALUES(?,?,?,'queued',?)",
                               (job_id, owner, item["item_id"], encode(item)))
        return job_id, True

    def launch(self, owner, job_id):
        task = asyncio.create_task(self.run(owner, job_id))
        self.tasks.add(task)
        task.add_done_callback(self.tasks.discard)

    def active(self, owner):
        with self.store.connect() as db:
            row = db.execute("SELECT active FROM users WHERE id=?", (owner,)).fetchone()
            return bool(row and row["active"])

    async def run(self, owner, job_id):
        job = self.store.job(owner, job_id)
        self.store.update_job(job_id, "running")
        try:
            if not self.active(owner):
                raise ValueError("账号已停用，任务停止")
            if job["kind"] == "source":
                async with self.ocr_slots:
                    if not self.active(owner):
                        raise ValueError("账号已停用，任务停止")
                    result = await asyncio.get_running_loop().run_in_executor(
                        self.ocr_pool, parse_sources, job["payload"]["files"], str(self.store.path), job_id)
            elif job["kind"] == "extract":
                snapshot = json.loads(self.decrypt(job["snapshot"]))
                defaults = merge_defaults(snapshot["profile"]["defaults"], job["payload"].get("defaults", {}))
                async with self.ai_slots:
                    if not self.active(owner):
                        raise ValueError("账号已停用，任务停止")
                    result = await extract(self, snapshot["llm"], job["payload"]["raw_text"], defaults,
                                           snapshot["profile"].get("members", {}),
                                           progress=lambda detail, done, total: report_progress(self.store.path, job_id, detail, done, total))
                for issue in result["issues"]:
                    issue["item_id"] = identifier()
                    issue.update({key: value for key, value in defaults.items() if key not in issue})
                    issue["custom_fields"] = {}
                    members = snapshot["profile"].get("members", {})
                    if issue.get("assignee") in members:
                        issue["assignee"] = members[issue["assignee"]]
            else:
                await self.run_batch(owner, job)
                return
            self.store.update_job(job_id, "done", result)
        except asyncio.CancelledError:
            raise
        except (ValueError, RemoteError) as exc:
            self.fail_job(job, str(exc))
        except Exception:
            self.fail_job(job, "任务处理失败，请检查配置及输入材料后重试")

    def fail_job(self, job, error):
        with self.store.connect() as db:
            db.execute("UPDATE items SET state='failed',error=? WHERE job_id=? AND state='queued'", (error, job["id"]))
            db.execute("UPDATE items SET state='unknown',error='执行中断，创建结果待核实' WHERE job_id=? AND state='creating'", (job["id"],))
        self.store.update_job(job["id"], "failed", error=error)

    async def run_batch(self, owner, job):
        snapshot = json.loads(self.decrypt(job["snapshot"]))
        defaults = merge_defaults(snapshot["profile"]["defaults"], job["payload"].get("defaults", {}))
        async with self.jira_slots:
            async with Jira(self, snapshot["jira"]) as jira:
                metadata = {}
                ready = []
                # Validate the complete batch before the first write to Jira.
                for entry in job["items"]:
                    if entry["state"] != "queued":
                        continue
                    issue = merge_defaults(defaults, entry["data"])
                    try:
                        key = (issue["project_key"], issue["issuetype_id"])
                        if key not in metadata:
                            metadata[key] = await jira.metadata(*key)
                        payload = await jira.validate(issue, snapshot["profile"].get("mappings", []), metadata[key])
                        ready.append((entry, payload))
                    except (ValueError, RemoteError) as exc:
                        self.store.item_state(job["id"], entry["item_id"], "failed", str(exc))
                for entry, payload in ready:
                    if not self.active(owner):
                        self.store.item_state(job["id"], entry["item_id"], "failed", "账号已停用，尚未发送到 Jira")
                        continue
                    self.store.item_state(job["id"], entry["item_id"], "creating")
                    try:
                        result = await jira.request("POST", "issue", json=payload)
                        key = result.get("key", "")
                        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*-\d+", key):
                            raise RemoteError("创建响应缺少有效 Issue Key，结果待核实", True)
                        self.store.item_state(job["id"], entry["item_id"], "success", key=key, url=jira.base + "/browse/" + key)
                    except RemoteError as exc:
                        self.store.item_state(job["id"], entry["item_id"], "unknown" if exc.uncertain else "failed", str(exc))
        self.store.update_job(job["id"], "done")

    async def cleanup_loop(self):
        while True:
            try:
                cutoff = time.time() - 86400
                with self.store.connect() as db:
                    active_dirs = {str(Path(file["path"]).parent) for row in db.execute(
                        "SELECT payload FROM jobs WHERE kind='source' AND status IN ('queued','running')")
                        for file in json.loads(row[0]).get("files", [])}
                    db.execute("DELETE FROM sessions WHERE expires_at<?", (time.time(),))
                    db.execute("DELETE FROM login_failures WHERE occurred_at<?", (time.time() - 900,))
                    db.execute("UPDATE jobs SET result='{}',payload='{}' WHERE kind='source' AND created_at<? "
                               "AND status NOT IN ('queued','running')", (cutoff,))
                    db.execute("UPDATE jobs SET payload='{}' WHERE kind='extract' AND created_at<? "
                               "AND status NOT IN ('queued','running')", (cutoff,))
                for path in self.upload_root.iterdir():
                    if (path.is_dir() and not path.is_symlink() and str(path) not in active_dirs
                            and path.stat().st_mtime < cutoff and path.resolve().parent == self.upload_root.resolve()):
                        await asyncio.to_thread(shutil.rmtree, path)
            except (OSError, sqlite3.Error):
                # A transient filesystem/database lock must not kill future cleanup passes.
                pass
            await asyncio.sleep(600)

    async def close(self):
        self.cleanup_task.cancel()
        for task in self.tasks:
            task.cancel()
        await asyncio.gather(self.cleanup_task, *self.tasks, return_exceptions=True)
        self.store.recover()
        self.ocr_pool.shutdown(wait=False, cancel_futures=True)
        self.lock_file.close()


def merge_defaults(base, overrides):
    # Only omitted fields inherit. An explicit null/empty value intentionally clears a default.
    return {**base, **overrides}
