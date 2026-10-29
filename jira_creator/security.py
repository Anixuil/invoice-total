import hashlib
import secrets
import time
from urllib.parse import urlsplit

from fastapi import HTTPException, Request

COOKIE = "jira_creator_session"
PREFIX = "/api/jira-create"


def digest(value):
    return hashlib.sha256(value.encode()).hexdigest()


def public_user(row):
    # Keep the response field for cached clients, but ignore legacy forced-change flags.
    return {**{key: row[key] for key in ("id", "username", "display_name", "role", "active")},
            "must_change": False}


def runtime(request: Request):
    service = getattr(request.app.state, "jira_creator", None)
    if service is None:
        raise HTTPException(503, getattr(request.app.state, "jira_creator_error", "新模块不可用"))
    return service


def same_origin(request: Request):
    if request.method in {"GET", "HEAD", "OPTIONS"}:
        return
    origin = request.headers.get("origin")
    expected = str(request.base_url).rstrip("/")
    service = runtime(request)
    if origin != (service.public_origin or expected):
        raise HTTPException(403, "请求来源不匹配，请通过本站页面操作")


def authenticated(request: Request, allow_change=False):
    service = runtime(request)
    token = request.cookies.get(COOKIE, "")
    with service.store.connect() as db:
        row = db.execute("SELECT u.*,s.csrf FROM sessions s JOIN users u ON u.id=s.user_id "
                         "WHERE s.token_hash=? AND s.expires_at>? AND u.active=1",
                         (digest(token), time.time())).fetchone()
    if not row:
        raise HTTPException(401, "请登录后使用 Jira 批量建单")
    request.state.creator_user_id = row["id"]
    if request.method not in {"GET", "HEAD", "OPTIONS"}:
        same_origin(request)
        if not secrets.compare_digest(request.headers.get("x-csrf-token", ""), row["csrf"]):
            raise HTTPException(403, "登录状态已变化，请刷新页面")
    return dict(row)


def user(request: Request):
    return authenticated(request)


def admin(request: Request):
    account = authenticated(request)
    if account["role"] != "admin":
        raise HTTPException(403, "此操作需要管理员权限")
    return account


def normalize_url(value):
    parsed = urlsplit(value.strip())
    if (parsed.scheme not in {"http", "https"} or not parsed.hostname or parsed.username or
            parsed.password or parsed.query or parsed.fragment or "\\" in value):
        raise ValueError("服务地址必须是无凭据、查询参数或锚点的 HTTP(S) URL")
    # Accessing port also validates malformed port input.
    port = parsed.port
    hostname = parsed.hostname.lower().encode("idna").decode("ascii")
    if ":" in hostname:
        hostname = f"[{hostname}]"
    origin = f"{parsed.scheme}://{hostname}"
    if port and port != (443 if parsed.scheme == "https" else 80):
        origin += f":{port}"
    return origin + parsed.path.rstrip("/"), origin


def checked_url(service, value):
    normalized, origin = normalize_url(value)
    if origin not in service.allowed_origins:
        raise ValueError("服务地址未列入部署白名单，请管理员配置 JIRA_CREATOR_ALLOWED_ORIGINS")
    return normalized
