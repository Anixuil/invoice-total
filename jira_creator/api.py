import asyncio
import json
import logging
import os
from pathlib import Path
import secrets
import shutil
import sqlite3
import time

from argon2.exceptions import VerificationError
from fastapi import APIRouter, Depends, File, HTTPException, Query, Request, Response, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse
from fastapi.routing import APIRoute

from . import models
from .clients import Jira, RemoteError, extract
from .config import ConfigurationError, load_environment
from .security import COOKIE, PREFIX, admin, authenticated, checked_url, digest, public_user, runtime, same_origin, user
from .service import Service, merge_defaults
from .storage import encode, identifier


class PrivateRoute(APIRoute):
    def get_route_handler(self):
        handler = super().get_route_handler()

        async def private(request):
            try:
                runtime(request)
                same_origin(request)
                if request.url.path != PREFIX + "/auth/login":
                    authenticated(request, allow_change=request.url.path in {
                        PREFIX + "/auth/me", PREFIX + "/auth/logout", PREFIX + "/auth/password"})
                length = int(request.headers.get("content-length", "0"))
                limit = 55 * 1024 * 1024 if request.url.path.endswith("/sources") else 12 * 1024 * 1024
                if length > limit:
                    raise HTTPException(413, "请求体超过限制")
                received = 0
                receive = request.receive

                async def bounded_receive():
                    nonlocal received
                    message = await receive()
                    received += len(message.get("body", b""))
                    if received > limit:
                        raise HTTPException(413, "请求体超过限制")
                    return message

                response = await handler(Request(request.scope, receive=bounded_receive))
            except RequestValidationError as exc:
                # FastAPI's default validation payload includes raw input, including passwords.
                details = [f"{'.'.join(map(str, error['loc']))}: {error['msg']}" for error in exc.errors()]
                response = JSONResponse({"detail": "；".join(details)}, status_code=422)
            except (ValueError, RemoteError) as exc:
                response = JSONResponse({"detail": str(exc)}, status_code=422)
            except HTTPException as exc:
                response = JSONResponse({"detail": exc.detail}, status_code=exc.status_code, headers=exc.headers)
            except sqlite3.IntegrityError:
                response = JSONResponse({"detail": "账号名称或提交标识已存在，请重新加载"}, status_code=409)
            except Exception:
                response = JSONResponse({"detail": "操作未完成，请稍后重试或联系管理员检查新模块数据目录"}, status_code=500)
            response.headers["Cache-Control"] = "no-store"
            if getattr(request.state, "creator_user_id", None):
                response.headers["X-Jira-Creator-User"] = request.state.creator_user_id
            return response
        return private


router = APIRouter(prefix=PREFIX, route_class=PrivateRoute)


async def start(app):
    try:
        load_environment()
        if os.environ.get("JIRA_CREATOR_ENABLED", "true").strip().lower() != "true":
            app.state.jira_creator_error = "Jira 批量建单已关闭，请将 JIRA_CREATOR_ENABLED 设置为 true"
            return
        app.state.jira_creator = Service()
        app.state.jira_creator_error = ""
        return
    except ConfigurationError as exc:
        message = str(exc)
    except PermissionError:
        message = "Jira 批量建单数据目录不可写，请检查 JIRA_CREATOR_DATA_DIR 及目录权限"
    except sqlite3.Error:
        message = "Jira 批量建单数据库无法打开或结构不兼容，请检查数据目录权限和数据库备份"
    except Exception as exc:
        # Report the exception class only; arbitrary payloads can contain secrets.
        message = f"Jira 批量建单初始化失败（{type(exc).__name__}），请检查模块依赖和数据目录"
    app.state.jira_creator_error = message
    logging.getLogger(__name__).warning("Jira creator unavailable: %s", message)


async def stop(app):
    if app.state.jira_creator:
        await app.state.jira_creator.close()


def record_or_404(service, owner, kind, record_id):
    record = service.store.record(owner, kind, record_id)
    if not record:
        raise HTTPException(404, "记录不存在")
    return record


def public_record(record):
    return {"id": record["id"], "name": record["name"], "data": record["data"],
            "version": record["version"], "updated_at": record["updated_at"], "has_secret": bool(record["secret"])}


def public_job(job):
    counts = {key: sum(item["state"] == key for item in job["items"])
              for key in ("queued", "creating", "success", "failed", "unknown")}
    return {"id": job["id"], "kind": job["kind"], "status": job["status"], "created_at": job["created_at"],
            "updated_at": job["updated_at"], "result": job["result"], "error": job["error"],
            "counts": counts, "total": len(job["items"]), "items": [
                {key: item[key] for key in ("item_id", "state", "data", "issue_key", "issue_url", "error", "attempts")}
                for item in job["items"]]}


@router.post("/auth/login")
async def login(body: models.Login, request: Request, response: Response):
    service = runtime(request)
    username = body.username.strip().casefold()
    now = time.time()
    with service.store.connect() as db:
        account = db.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
    try:
        await asyncio.to_thread(service.hasher.verify, account["password_hash"] if account else service.dummy_hash, body.password)
        valid = bool(account and account["active"])
    except VerificationError:
        valid = False
    if not valid:
        raise HTTPException(401, "账号或密码不正确，或账号已停用")
    token, csrf = secrets.token_urlsafe(32), secrets.token_urlsafe(32)
    with service.store.connect() as db:
        current = db.execute("SELECT * FROM users WHERE id=?", (account["id"],)).fetchone()
        if not current["active"] or current["password_hash"] != account["password_hash"]:
            raise HTTPException(401, "账号状态已变化，请重新登录")
        db.execute("DELETE FROM sessions WHERE token_hash=?", (digest(request.cookies.get(COOKIE, "")),))
        db.execute("INSERT INTO sessions VALUES(?,?,?,?)", (digest(token), account["id"], csrf, now + 8 * 3600))
    response.set_cookie(COOKIE, token, max_age=8 * 3600, httponly=True, secure=service.secure_cookie, samesite="strict", path=PREFIX)
    return {"user": public_user(account), "csrf": csrf}


@router.get("/auth/me")
def me(request: Request):
    account = authenticated(request, allow_change=True)
    return {"user": public_user(account), "csrf": account["csrf"]}


@router.post("/auth/logout")
def logout(request: Request, response: Response):
    authenticated(request, allow_change=True)
    with runtime(request).store.connect() as db:
        db.execute("DELETE FROM sessions WHERE token_hash=?", (digest(request.cookies.get(COOKIE, "")),))
    response.delete_cookie(COOKIE, path=PREFIX)
    return {"ok": True}


@router.post("/auth/password")
async def password(body: models.PasswordChange, request: Request, response: Response):
    account = authenticated(request, allow_change=True)
    service = runtime(request)
    try:
        await asyncio.to_thread(service.hasher.verify, account["password_hash"], body.current_password)
    except VerificationError:
        raise HTTPException(422, "当前密码不正确")
    hashed = await asyncio.to_thread(service.hasher.hash, body.password)
    authenticated(request, allow_change=True)
    with service.store.connect() as db:
        changed = db.execute("UPDATE users SET password_hash=?,must_change=0 WHERE id=? AND password_hash=? AND active=1",
                             (hashed, account["id"], account["password_hash"])).rowcount
        if not changed:
            raise HTTPException(409, "账号密码已变化，请重新登录")
        db.execute("DELETE FROM sessions WHERE user_id=?", (account["id"],))
    response.delete_cookie(COOKIE, path=PREFIX)
    return {"ok": True, "message": "密码已修改，请重新登录"}


@router.get("/admin/users")
def users(request: Request, account=Depends(admin)):
    with runtime(request).store.connect() as db:
        return [public_user(row) for row in db.execute("SELECT * FROM users ORDER BY created_at")]


@router.post("/admin/users")
async def create_user(body: models.UserCreate, request: Request, account=Depends(admin)):
    service = runtime(request)
    hashed = await asyncio.to_thread(service.hasher.hash, body.password)
    admin(request)
    user_id = identifier()
    with service.store.connect() as db:
        db.execute("INSERT INTO users(id,username,display_name,password_hash,role,created_at,must_change) VALUES(?,?,?,?,?,?,0)",
                   (user_id, body.username.casefold(), body.display_name, hashed, body.role, time.time()))
    service.store.seed_defaults(user_id)
    return {"id": user_id}


@router.put("/admin/users/{user_id}")
async def edit_user(user_id: str, body: models.UserEdit, request: Request, account=Depends(admin)):
    service = runtime(request)
    hashed = await asyncio.to_thread(service.hasher.hash, body.password) if body.password else None
    admin(request)
    with service.store.connect() as db:
        db.execute("BEGIN IMMEDIATE")
        previous = db.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
        if not previous:
            raise HTTPException(404, "账号不存在")
        if previous["role"] == "admin" and previous["active"] and (body.role != "admin" or not body.active):
            if db.execute("SELECT COUNT(*) FROM users WHERE active=1 AND role='admin'").fetchone()[0] <= 1:
                raise HTTPException(422, "不能停用或降级最后一个有效管理员")
        db.execute("UPDATE users SET display_name=?,role=?,active=? WHERE id=?", (body.display_name, body.role, body.active, user_id))
        if hashed:
            db.execute("UPDATE users SET password_hash=?,must_change=0 WHERE id=?", (hashed, user_id))
        if hashed or not body.active or body.role != previous["role"]:
            db.execute("DELETE FROM sessions WHERE user_id=?", (user_id,))
    return {"ok": True}


@router.get("/connections")
def connections(request: Request, account=Depends(user)):
    return [public_record(row) for row in runtime(request).store.records(account["id"], "connection")]


def save_connection(request, account, body, record_id=None):
    service = runtime(request)
    data = body.model_dump(exclude={"secret", "clear_secret", "version", "name"})
    data["base_url"] = checked_url(service, data["base_url"])
    if body.kind == "jira" and not body.username:
        raise ValueError("Jira 用户名不能为空")
    if body.kind == "llm" and not body.model:
        raise ValueError("模型名称不能为空")
    old = record_or_404(service, account["id"], "connection", record_id) if record_id else None
    if old and old["data"]["kind"] != body.kind:
        raise ValueError("不能修改连接类型，请新增连接")
    secret = "" if body.clear_secret else service.encrypt(body.secret) if body.secret else old["secret"] if old else ""
    result = service.store.save(account["id"], "connection", body.name, data, record_id, secret, body.version)
    return public_record(result)


@router.post("/connections")
def create_connection(body: models.Connection, request: Request, account=Depends(user)):
    return save_connection(request, account, body)


@router.put("/connections/{record_id}")
def edit_connection(record_id: str, body: models.Connection, request: Request, account=Depends(user)):
    return save_connection(request, account, body, record_id)


@router.delete("/connections/{record_id}")
def delete_connection(record_id: str, request: Request, account=Depends(user)):
    service = runtime(request)
    record_or_404(service, account["id"], "connection", record_id)
    for profile in service.store.records(account["id"], "profile"):
        if record_id in (profile["data"].get("jira_connection_id"), profile["data"].get("llm_connection_id")):
            raise HTTPException(409, "连接仍被建单方案引用，请先解除引用")
    with service.store.connect() as db:
        db.execute("DELETE FROM records WHERE id=? AND owner=?", (record_id, account["id"]))
    return {"ok": True}


@router.post("/connections/{record_id}/check")
async def check_connection(record_id: str, request: Request, account=Depends(user)):
    service = runtime(request)
    record = record_or_404(service, account["id"], "connection", record_id)
    connection = service.connection(account["id"], record_id, record["data"]["kind"])
    if connection["kind"] == "jira":
        async with service.jira_slots:
            async with Jira(service, connection) as jira:
                person = await jira.request("GET", "myself")
                meta = await jira.metadata()
        return {"message": f"连接成功：{person.get('displayName', '')} / Jira {meta['version']}", **meta}
    async with service.ai_slots:
        await extract(service, connection, "连接检查，没有待办事项，请返回空 issues 数组。", {}, {}, probe=True)
    return {"message": "模型连接及结构化输出检查通过"}


@router.get("/profiles")
def profiles(request: Request, account=Depends(user)):
    service = runtime(request)
    with service.store.connect() as db:
        row = db.execute("SELECT value FROM settings WHERE key=?", ("default:" + account["id"],)).fetchone()
    return {"items": [public_record(row) for row in service.store.records(account["id"], "profile")],
            "default_id": row["value"] if row else None}


def save_profile(request, account, body, record_id=None):
    service = runtime(request)
    if record_id:
        record_or_404(service, account["id"], "profile", record_id)
    for field, kind in (("jira_connection_id", "jira"), ("llm_connection_id", "llm")):
        value = getattr(body, field)
        if value:
            connection = record_or_404(service, account["id"], "connection", value)
            if connection["data"]["kind"] != kind:
                raise ValueError("关联的连接类型不正确")
    keys = [(m.project_key, m.issuetype_id, m.field_id) for m in body.mappings]
    if len(keys) != len(set(keys)):
        raise ValueError("同一项目及类型的字段映射不能重复")
    return public_record(service.store.save(account["id"], "profile", body.name,
                         body.model_dump(exclude={"name", "version"}), record_id, version=body.version))


@router.post("/profiles")
def create_profile(body: models.Profile, request: Request, account=Depends(user)):
    return save_profile(request, account, body)


@router.put("/profiles/{record_id}")
def edit_profile(record_id: str, body: models.Profile, request: Request, account=Depends(user)):
    return save_profile(request, account, body, record_id)


@router.post("/profiles/{record_id}/default")
def default_profile(record_id: str, request: Request, account=Depends(user)):
    service = runtime(request)
    record_or_404(service, account["id"], "profile", record_id)
    with service.store.connect() as db:
        db.execute("INSERT OR REPLACE INTO settings VALUES(?,?)", ("default:" + account["id"], record_id))
    return {"ok": True}


@router.delete("/profiles/{record_id}")
def delete_profile(record_id: str, request: Request, account=Depends(user)):
    service = runtime(request)
    record_or_404(service, account["id"], "profile", record_id)
    with service.store.connect() as db:
        db.execute("DELETE FROM records WHERE id=? AND owner=?", (record_id, account["id"]))
        db.execute("DELETE FROM settings WHERE key=? AND value=?", ("default:" + account["id"], record_id))
    return {"ok": True}


@router.get("/templates")
def templates(request: Request, account=Depends(user)):
    return [public_record(row) for row in runtime(request).store.records("public", "template")]


@router.post("/admin/templates")
def create_template(body: models.Template, request: Request, account=Depends(admin)):
    service = runtime(request)
    for value in (body.jira_base_url, body.llm_base_url):
        if value:
            checked_url(service, value)
    return public_record(service.store.save("public", "template", body.name,
                         body.model_dump(exclude={"name", "version"})))


@router.put("/admin/templates/{record_id}")
def edit_template(record_id: str, body: models.Template, request: Request, account=Depends(admin)):
    service = runtime(request)
    record_or_404(service, "public", "template", record_id)
    for value in (body.jira_base_url, body.llm_base_url):
        if value:
            checked_url(service, value)
    return public_record(service.store.save("public", "template", body.name,
                         body.model_dump(exclude={"name", "version"}), record_id, version=body.version))


@router.delete("/admin/templates/{record_id}")
def delete_template(record_id: str, request: Request, account=Depends(admin)):
    with runtime(request).store.connect() as db:
        db.execute("DELETE FROM records WHERE owner='public' AND kind='template' AND id=?", (record_id,))
    return {"ok": True}


@router.post("/templates/{record_id}/copy")
def copy_template(record_id: str, request: Request, account=Depends(user)):
    service = runtime(request)
    template = record_or_404(service, "public", "template", record_id)
    data = template["data"]
    profile = {"defaults": data["defaults"], "members": data["members"], "mappings": data["mappings"],
               "jira_connection_id": None, "llm_connection_id": None}
    with service.store.connect() as db:
        for kind in ("jira", "llm"):
            if not data.get(kind + "_base_url"):
                continue
            connection_id = identifier()
            connection = {"kind": kind, "base_url": checked_url(service, data[kind + "_base_url"]),
                          "username": "", "model": data["model"] if kind == "llm" else "",
                          "output_mode": data["output_mode"]}
            db.execute("INSERT INTO records(id,owner,kind,name,data,updated_at) VALUES(?,?,'connection',?,?,?)",
                       (connection_id, account["id"], template["name"] + " · " + kind, encode(connection), time.time()))
            profile[kind + "_connection_id"] = connection_id
        profile_id = identifier()
        db.execute("INSERT INTO records(id,owner,kind,name,data,updated_at) VALUES(?,?,'profile',?,?,?)",
                   (profile_id, account["id"], template["name"] + " · 我的副本", encode(profile), time.time()))
    return {"id": profile_id, "message": "模板已复制，请填写个人凭据并检查字段映射"}


@router.get("/metadata")
async def metadata(request: Request, connection_id: str, project: str = "", issuetype: str = "", account=Depends(user)):
    service = runtime(request)
    connection = service.connection(account["id"], connection_id, "jira")
    async with service.jira_slots:
        async with Jira(service, connection) as jira:
            return await jira.metadata(project, issuetype)


@router.get("/metadata/users")
async def metadata_users(request: Request, connection_id: str, project: str, query: str = "", account=Depends(user)):
    service = runtime(request)
    async with service.jira_slots:
        async with Jira(service, service.connection(account["id"], connection_id, "jira")) as jira:
            return await jira.users(project, query)


@router.post("/sources", status_code=202)
async def sources(request: Request, files: list[UploadFile] = File(...), account=Depends(user)):
    service = runtime(request)
    directory = service.upload_root / identifier()
    try:
        if not 1 <= len(files) <= 10:
            raise HTTPException(422, "每批上传 1 至 10 个文件")
        directory.mkdir()
        entries = []
        total = 0
        for upload in files:
            name = Path((upload.filename or "材料").replace("\\", "/")).name
            suffix = Path(name).suffix.lower()
            if suffix not in {".txt", ".md", ".docx", ".xlsx", ".pdf", ".png", ".jpg", ".jpeg", ".webp"}:
                raise HTTPException(422, f"不支持文件格式：{name}")
            path = directory / (identifier() + suffix)
            size = 0
            with path.open("wb") as target:
                while chunk := await upload.read(1024 * 1024):
                    size += len(chunk)
                    total += len(chunk)
                    if size > 20 * 1024 * 1024 or total > 50 * 1024 * 1024:
                        raise HTTPException(413, "单文件不能超过 20 MB，每批总计不能超过 50 MB")
                    target.write(chunk)
            if not size:
                raise HTTPException(422, f"文件为空：{name}")
            entries.append({"name": name, "path": str(path)})
        job_id, _ = service.submit(account["id"], "source", {"files": entries})
        service.launch(account["id"], job_id)
        return {"job_id": job_id}
    except Exception:
        if directory.exists():
            shutil.rmtree(directory)
        raise
    finally:
        for upload in files:
            await upload.close()


@router.post("/extract", status_code=202)
async def extraction(body: models.ExtractRequest, request: Request, account=Depends(user)):
    service = runtime(request)
    snapshot = service.snapshot(account["id"], body.profile_id, needs_model=True)
    job_id, _ = service.submit(account["id"], "extract", body.model_dump(exclude_unset=True), snapshot)
    service.launch(account["id"], job_id)
    return {"job_id": job_id}


@router.get("/drafts")
def drafts(request: Request, account=Depends(user)):
    return [public_record(row) for row in runtime(request).store.records(account["id"], "draft")]


@router.post("/drafts")
def create_draft(body: models.Draft, request: Request, account=Depends(user)):
    return public_record(runtime(request).store.save(account["id"], "draft", body.name,
                         body.model_dump(exclude={"name", "version"})))


@router.put("/drafts/{record_id}")
def edit_draft(record_id: str, body: models.Draft, request: Request, account=Depends(user)):
    service = runtime(request)
    record_or_404(service, account["id"], "draft", record_id)
    return public_record(service.store.save(account["id"], "draft", body.name,
                         body.model_dump(exclude={"name", "version"}), record_id, version=body.version))


@router.delete("/drafts/{record_id}")
def delete_draft(record_id: str, request: Request, account=Depends(user)):
    with runtime(request).store.connect() as db:
        db.execute("DELETE FROM records WHERE id=? AND owner=? AND kind='draft'", (record_id, account["id"]))
    return {"ok": True}


@router.post("/validate")
async def validate_batch(body: models.Batch, request: Request, account=Depends(user)):
    service = runtime(request)
    snapshot = service.snapshot(account["id"], body.profile_id)
    defaults = merge_defaults(snapshot["profile"]["defaults"], body.defaults.model_dump(exclude_unset=True))
    results, cache = [], {}
    async with service.jira_slots:
        async with Jira(service, snapshot["jira"]) as jira:
            for item in body.issues:
                issue = merge_defaults(defaults, item.model_dump(exclude_unset=True))
                try:
                    key = (issue["project_key"], issue["issuetype_id"])
                    if key not in cache:
                        cache[key] = await jira.metadata(*key)
                    await jira.validate(issue, snapshot["profile"]["mappings"], cache[key])
                    results.append({"item_id": item.item_id, "valid": True, "error": ""})
                except (ValueError, RemoteError) as exc:
                    results.append({"item_id": item.item_id, "valid": False, "error": str(exc)})
    return {"items": results}


@router.post("/batch-create", status_code=202)
async def batch_create(body: models.Batch, request: Request, account=Depends(user)):
    service = runtime(request)
    payload = body.model_dump(exclude={"idempotency_key"}, exclude_unset=True)
    # Idempotent replay must work even when the original profile was edited/deleted.
    with service.store.connect() as db:
        existing = db.execute("SELECT id,digest FROM jobs WHERE owner=? AND idem=?", (account["id"], body.idempotency_key)).fetchone()
    if existing:
        import hashlib
        if existing["digest"] != hashlib.sha256(encode(payload).encode()).hexdigest():
            raise HTTPException(409, "该提交标识已用于不同内容")
        return {"job_id": existing["id"]}
    snapshot = service.snapshot(account["id"], body.profile_id)
    job_id, created = service.submit(account["id"], "batch", payload, snapshot, body.idempotency_key)
    if created:
        service.launch(account["id"], job_id)
    return {"job_id": job_id}


@router.get("/jobs")
def jobs(request: Request, offset: int = Query(0, ge=0), account=Depends(user)):
    service = runtime(request)
    with service.store.connect() as db:
        rows = db.execute("SELECT id,kind,status,error,created_at,updated_at FROM jobs WHERE owner=? "
                          "ORDER BY created_at DESC LIMIT 30 OFFSET ?", (account["id"], offset)).fetchall()
    return [dict(row) for row in rows]


@router.get("/jobs/{job_id}")
def job(job_id: str, request: Request, account=Depends(user)):
    service = runtime(request)
    record = service.store.job(account["id"], job_id)
    if not record:
        raise HTTPException(404, "任务不存在")
    result = public_job(record)
    if record["kind"] == "batch":
        snapshot = json.loads(service.decrypt(record["snapshot"]))
        result["configuration"] = {
            "name": snapshot["profile_name"], "version": snapshot["profile_version"],
            "jira_base_url": snapshot["jira"]["base_url"],
            "defaults": merge_defaults(snapshot["profile"]["defaults"], record["payload"].get("defaults", {})),
            "mappings": snapshot["profile"].get("mappings", []),
        }
    return result


@router.post("/jobs/{job_id}/items/{item_id}/resolve")
async def resolve(job_id: str, item_id: str, body: models.Resolve, request: Request, account=Depends(user)):
    service = runtime(request)
    job = service.store.job(account["id"], job_id)
    if not job or job["kind"] != "batch":
        raise HTTPException(404, "批次不存在")
    entry = next((item for item in job["items"] if item["item_id"] == item_id), None)
    if not entry or entry["state"] not in {"unknown", "failed"} or job["status"] in {"queued", "running"}:
        raise HTTPException(409, "仅可处理已结束批次中的失败或待核实事项")
    snapshot = json.loads(service.decrypt(job["snapshot"]))
    if body.action == "link":
        if not body.issue_key:
            raise ValueError("请输入已有 Issue Key")
        async with service.jira_slots:
            async with Jira(service, snapshot["jira"]) as jira:
                existing = await jira.request("GET", "issue/" + body.issue_key, params={"fields": "project,summary"})
                defaults = merge_defaults(snapshot["profile"]["defaults"], job["payload"].get("defaults", {}))
                expected = merge_defaults(defaults, entry["data"])
                if (existing["fields"]["project"]["key"] != expected["project_key"] or
                        existing["fields"]["summary"] != expected["summary"]):
                    raise ValueError("该 Issue 的项目或标题与原事项不一致，请确认后关联")
                key = existing["key"]
                with service.store.connect() as db:
                    changed = db.execute("UPDATE items SET state='success',issue_key=?,issue_url=?,error='' "
                                         "WHERE job_id=? AND item_id=? AND state IN ('unknown','failed')",
                                         (key, jira.base + "/browse/" + key, job_id, item_id)).rowcount
                    if not changed:
                        raise HTTPException(409, "事项状态已变化，请刷新")
        return {"job_id": job_id}
    if entry["state"] == "unknown" and not body.confirmed_not_created:
        raise ValueError("请先到 Jira 核实，确认未创建后再重新提交")
    issue = body.issue.model_dump(exclude_unset=True) if body.issue else entry["data"]
    if issue["item_id"] != item_id:
        raise ValueError("不允许改变事项标识")
    with service.store.connect() as db:
        db.execute("BEGIN IMMEDIATE")
        if db.execute("SELECT 1 FROM jobs WHERE owner=? AND kind='batch' AND status IN ('queued','running')", (account["id"],)).fetchone():
            raise HTTPException(409, "已有创建批次正在执行")
        if db.execute("SELECT 1 FROM items WHERE owner=? AND item_id=? AND job_id!=? "
                      "AND state IN ('queued','creating','success','unknown')", (account["id"], item_id, job_id)).fetchone():
            raise HTTPException(409, "该事项已在其他批次处理")
        changed = db.execute("UPDATE items SET state='queued',data=?,error='' WHERE job_id=? AND item_id=? "
                             "AND state IN ('failed','unknown')", (encode(issue), job_id, item_id)).rowcount
        if not changed:
            raise HTTPException(409, "事项状态已变化，请刷新")
        db.execute("UPDATE jobs SET status='queued',error='',updated_at=? WHERE id=?", (time.time(), job_id))
    service.launch(account["id"], job_id)
    return {"job_id": job_id}
