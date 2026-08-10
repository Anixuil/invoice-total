#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
server.py — 本地发票核验与 Jira 数据处理 Web 服务（FastAPI）

启动: uvicorn server:app --host 0.0.0.0 --port 8000
API:
  GET  /             前端页面
  POST /api/upload   上传 PDF 或 ZIP，先返回金额核算结果和票面任务 ID
  GET  /api/jobs/{id} 轮询票面核对进度，完成后返回全部预览
  POST /api/jira/import  上传 Jira xlsx 并创建处理任务
  GET  /api/jira/jobs/{id} 轮询 Jira 处理进度
  POST /api/weekly-report/import  上传部门项目周报 ZIP
  GET  /api/weekly-report/jobs/{id} 轮询周报处理进度
"""

import base64
import stat
import shutil
import tempfile
import threading
import time
import uuid
import zipfile
from pathlib import Path
from pathlib import PurePosixPath

import fitz
from fastapi import BackgroundTasks, FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from invoice_total import Extractor, sum_money
from jira_processor import export_jira_xlsx, process_jira_workbook
from weekly_report_processor import (
    build_weekly_meeting_document,
    build_weekly_presentation,
    export_weekly_report_xlsx,
    process_weekly_report,
)

app = FastAPI(title="本地文件处理工具", version="1.1.0")

BASE = Path(__file__).resolve().parent
STATIC = BASE / "static"
MAX_SIZE = 20 * 1024 * 1024  # 20MB
MAX_FILES = 50
MAX_ARCHIVE_UNPACKED_SIZE = 200 * 1024 * 1024  # ZIP 展开总量 200MB
WEEKLY_MAX_SIZE = 500 * 1024 * 1024  # 部门周报 ZIP/PPTX 单文件 500MB
WEEKLY_MAX_ARCHIVE_UNPACKED_SIZE = 2 * 1024 * 1024 * 1024  # 部门周报 ZIP 展开总量 2GB
READ_CHUNK_SIZE = 1024 * 1024
MAX_PREVIEW_PAGES = 100
PREVIEW_SCALE = 1.5
PREVIEW_JPEG_QUALITY = 76
PREVIEW_JOB_TTL = 60 * 60
PREVIEW_JOB_ROOT = Path(tempfile.gettempdir()) / "invoice_total_preview_jobs"
PREVIEW_JOBS = {}
PREVIEW_JOBS_LOCK = threading.Lock()
JIRA_JOB_TTL = 60 * 60
JIRA_JOB_ROOT = Path(tempfile.gettempdir()) / "invoice_total_jira_jobs"
JIRA_JOBS = {}
JIRA_JOBS_LOCK = threading.Lock()
WEEKLY_JOB_TTL = 60 * 60
WEEKLY_JOB_ROOT = Path(tempfile.gettempdir()) / "invoice_total_weekly_jobs"
WEEKLY_JOBS = {}
WEEKLY_JOBS_LOCK = threading.Lock()

app.mount("/static", StaticFiles(directory=STATIC), name="static")


def build_page_previews(pdf_path: Path, page_numbers: list[int] | None = None, progress_callback=None):
    """将 PDF 页面转为轻量 JPEG Data URL，随本次响应返回。"""
    previews = []
    with fitz.open(str(pdf_path)) as doc:
        selected_pages = page_numbers or list(range(1, doc.page_count + 1))
        for page_number in selected_pages[:MAX_PREVIEW_PAGES]:
            if not 1 <= page_number <= doc.page_count:
                continue
            try:
                pixmap = doc[page_number - 1].get_pixmap(
                    matrix=fitz.Matrix(PREVIEW_SCALE, PREVIEW_SCALE), alpha=False
                )
                jpeg = pixmap.tobytes("jpeg", jpg_quality=PREVIEW_JPEG_QUALITY)
                encoded = base64.b64encode(jpeg).decode("ascii")
                previews.append({
                    "page": page_number,
                    "data_url": f"data:image/jpeg;base64,{encoded}",
                })
            except Exception:
                continue
            finally:
                if progress_callback:
                    progress_callback(page_number, min(doc.page_count, MAX_PREVIEW_PAGES))
    return previews


@app.get("/")
async def index():
    return FileResponse(STATIC / "index.html")


@app.get("/jira")
async def jira_index():
    return FileResponse(STATIC / "jira.html")


@app.get("/weekly-report")
async def weekly_report_index():
    return FileResponse(STATIC / "weekly-report.html")


def _safe_name(raw_name: str) -> str:
    """只保留用户可见的文件名，避免把上传路径带入结果。"""
    return Path(raw_name.replace("\\", "/")).name.strip() or "unnamed"


def _safe_archive_member_name(raw_name: str) -> str:
    """规范化 ZIP 成员名；拒绝绝对路径和 ..，防止路径穿越。"""
    normalized = raw_name.replace("\\", "/")
    path = PurePosixPath(normalized)
    if path.is_absolute() or any(part == ".." for part in path.parts):
        raise ValueError("压缩包包含不安全的目录路径")
    clean = "/".join(part for part in path.parts if part not in ("", "."))
    return clean or "unnamed.pdf"


async def _save_upload(upload: UploadFile, path: Path, max_size: int = MAX_SIZE) -> tuple[int, bytes]:
    """将上传流写入临时文件，并返回大小与文件头。"""
    size = 0
    prefix = bytearray()
    with path.open("wb") as output:
        while chunk := await upload.read(READ_CHUNK_SIZE):
            size += len(chunk)
            if size > max_size:
                return size, bytes(prefix)
            if len(prefix) < 1024:
                prefix.extend(chunk[: 1024 - len(prefix)])
            output.write(chunk)
    return size, bytes(prefix)


def _result_for_pdf(path: Path, display_name: str):
    """复用现有发票引擎处理一个 PDF；票面预览交给后台任务。"""
    try:
        with path.open("rb") as source:
            prefix = source.read(1024)
        if b"%PDF-" not in prefix:
            return {"file": display_name, "ok": False, "error": "文件内容不是有效的 PDF"}
        ex = Extractor(path)
        ex.load()
        result = ex.run()
        result["file"] = display_name
        result["previews"] = []
        result["preview_status"] = "pending"
        result["preview_pages_total"] = getattr(ex, "page_count", 0)
        result["_preview_path"] = str(path)
        return result
    except Exception as exc:
        return {"file": display_name, "ok": False, "error": f"解析失败: {exc}"}


def _archive_pdf_infos(archive: zipfile.ZipFile):
    """返回 ZIP 中的 PDF 成员和目录清单，并校验归档安全限制。"""
    infos = []
    manifest = []
    expanded_size = 0
    for info in archive.infolist():
        member_name = _safe_archive_member_name(info.filename)
        mode = (info.external_attr >> 16) & 0xFFFF
        if info.is_dir() or stat.S_ISDIR(mode):
            manifest.append({"path": member_name, "kind": "directory"})
            continue
        if not member_name.lower().endswith(".pdf"):
            manifest.append({"path": member_name, "kind": "ignored", "size": info.file_size})
            continue
        if stat.S_ISLNK(mode):
            raise ValueError(f"压缩包包含不支持的软链接: {member_name}")
        if info.flag_bits & 0x1:
            raise ValueError(f"压缩包包含加密文件: {member_name}")
        if info.file_size > MAX_SIZE:
            raise ValueError(f"压缩包内文件超过 20MB: {member_name}")
        expanded_size += info.file_size
        if expanded_size > MAX_ARCHIVE_UNPACKED_SIZE:
            raise ValueError("压缩包展开后超过 200MB 限制")
        infos.append((info, member_name))
        manifest.append({"path": member_name, "kind": "pdf", "size": info.file_size})
    return infos, manifest


def _extract_archive_member(
    archive: zipfile.ZipFile,
    info: zipfile.ZipInfo,
    path: Path,
    max_size: int = MAX_SIZE,
) -> None:
    """将一个 ZIP 成员写入临时路径，并再次限制实际展开大小。"""
    written = 0
    with archive.open(info, "r") as source, path.open("wb") as output:
        while chunk := source.read(READ_CHUNK_SIZE):
            written += len(chunk)
            if written > max_size:
                raise ValueError(f"压缩包内文件超过 {max_size // (1024 * 1024)}MB")
            output.write(chunk)


def _public_result(result):
    """移除后台任务使用的临时路径，避免把服务器路径返回给浏览器。"""
    public = dict(result)
    public.pop("_preview_path", None)
    public["previews"] = list(result.get("previews") or [])
    return public


def _cleanup_preview_jobs() -> None:
    """清理已完成且超过 TTL 的票面任务，避免临时文件长期占用磁盘。"""
    now = time.time()
    expired = []
    with PREVIEW_JOBS_LOCK:
        for job_id, job in PREVIEW_JOBS.items():
            if job.get("status") in {"done", "error"} and now - job.get("updated_at", now) > PREVIEW_JOB_TTL:
                expired.append((job_id, job.get("directory")))
        for job_id, _ in expired:
            PREVIEW_JOBS.pop(job_id, None)
    for _, directory in expired:
        if directory:
            shutil.rmtree(directory, ignore_errors=True)


def _preview_job_response(job, include_results=False):
    response = {
        "job_id": job["job_id"],
        "status": job["status"],
        "preview": dict(job["preview"]),
    }
    if include_results or job["status"] in {"done", "error"}:
        response.update({
            "grand_total": job["grand_total"],
            "files": job["files"],
            "ok_count": job["ok_count"],
            "low_confidence_count": job["low_confidence_count"],
            "sources": job["sources"],
            "results": [_public_result(result) for result in job["results"]],
        })
    return response


def _process_preview_job(job_id: str) -> None:
    """后台按 PDF 逐个渲染票面，并更新总进度和当前文件页进度。"""
    with PREVIEW_JOBS_LOCK:
        job = PREVIEW_JOBS.get(job_id)
        if not job:
            return
        job["status"] = "running"
        job["updated_at"] = time.time()
    try:
        preview_results = [result for result in job["results"] if result.get("_preview_path")]
        total_files = len(preview_results)
        with PREVIEW_JOBS_LOCK:
            job["preview"].update({"status": "running", "total_files": total_files})
        for file_index, result in enumerate(preview_results, start=1):
            page_total = min(int(result.get("preview_pages_total") or 0), MAX_PREVIEW_PAGES)
            with PREVIEW_JOBS_LOCK:
                result["preview_status"] = "running"
                job["preview"].update({
                    "completed_files": file_index - 1,
                    "current_file": result.get("file", ""),
                    "current_file_index": file_index,
                    "current_page": 0,
                    "current_pages": page_total,
                })

            def on_page(page_number, page_count):
                with PREVIEW_JOBS_LOCK:
                    job["preview"].update({"current_page": page_number, "current_pages": page_count})
                    job["updated_at"] = time.time()

            previews = build_page_previews(Path(result["_preview_path"]), progress_callback=on_page)
            with PREVIEW_JOBS_LOCK:
                result["previews"] = previews
                result["preview_status"] = "done"
                result["preview_truncated"] = int(result.get("preview_pages_total") or 0) > MAX_PREVIEW_PAGES
                job["preview"].update({"completed_files": file_index, "current_page": page_total})
                job["updated_at"] = time.time()

        with PREVIEW_JOBS_LOCK:
            job["status"] = "done"
            job["preview"].update({
                "status": "done",
                "completed_files": total_files,
                "current_file": "",
                "current_file_index": total_files,
            })
            job["updated_at"] = time.time()
    except Exception as exc:
        with PREVIEW_JOBS_LOCK:
            job["status"] = "error"
            job["preview"].update({"status": "error", "message": f"票面预览处理失败: {exc}"})
            job["updated_at"] = time.time()


@app.get("/api/jobs/{job_id}")
async def preview_job(job_id: str):
    _cleanup_preview_jobs()
    with PREVIEW_JOBS_LOCK:
        job = PREVIEW_JOBS.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="票面核对任务不存在或已过期")
        return _preview_job_response(job)


@app.post("/api/upload")
async def upload(background_tasks: BackgroundTasks, files: list[UploadFile] = File(...)):
    if not files:
        raise HTTPException(status_code=400, detail="未收到文件")
    if len(files) > MAX_FILES:
        raise HTTPException(status_code=400, detail=f"单次最多上传 {MAX_FILES} 个文件")

    _cleanup_preview_jobs()
    PREVIEW_JOB_ROOT.mkdir(parents=True, exist_ok=True)
    job_id = uuid.uuid4().hex
    job_directory = Path(tempfile.mkdtemp(prefix=f"{job_id}_", dir=PREVIEW_JOB_ROOT))
    results = []
    sources = []
    processed_pdf_count = 0

    def add_result(result, source_entry=None):
        result_index = len(results)
        results.append(result)
        if source_entry is not None:
            source_entry["result_index"] = result_index
            source_entry["ok"] = bool(result.get("ok"))
            source_entry["total"] = result.get("total")

    try:
        for index, f in enumerate(files):
            raw_name = f.filename or "unnamed.pdf"
            name = _safe_name(raw_name)
            source_type = "zip" if name.lower().endswith(".zip") else "pdf" if name.lower().endswith(".pdf") else "unknown"
            source = {"name": name, "type": source_type, "entries": []}
            try:
                path = job_directory / f"upload_{index}"
                size, prefix = await _save_upload(f, path)
                if size > MAX_SIZE:
                    entry = {"path": name, "kind": source_type}
                    source["entries"] = [entry]
                    add_result({"file": name, "ok": False, "error": "文件超过 20MB 限制"}, entry)
                    continue
                if size == 0:
                    entry = {"path": name, "kind": source_type}
                    source["entries"] = [entry]
                    add_result({"file": name, "ok": False, "error": "文件内容为空"}, entry)
                    continue

                if source_type == "pdf":
                    entry = {"path": name, "kind": "pdf"}
                    source["entries"] = [entry]
                    if b"%PDF-" not in prefix:
                        add_result({"file": name, "ok": False, "error": "文件内容不是有效的 PDF"}, entry)
                    else:
                        add_result(_result_for_pdf(path, name), entry)
                    processed_pdf_count += 1
                    continue

                if source_type != "zip":
                    entry = {"path": name, "kind": "unknown"}
                    source["entries"] = [entry]
                    add_result({"file": name, "ok": False, "error": "仅支持 PDF 或 ZIP 文件"}, entry)
                    continue

                with zipfile.ZipFile(path) as archive:
                    infos, manifest = _archive_pdf_infos(archive)
                    source["entries"] = manifest
                    source["pdf_count"] = len(infos)
                    if not infos:
                        add_result({"file": name, "ok": False, "error": "压缩包内没有 PDF 文件"})
                        continue
                    if processed_pdf_count + len(infos) > MAX_FILES:
                        add_result({"file": name, "ok": False, "error": f"单次最多解析 {MAX_FILES} 个 PDF 文件"})
                        continue
                    for member_index, (info, member_name) in enumerate(infos):
                        member_path = job_directory / f"archive_{index}_{member_index}.pdf"
                        _extract_archive_member(archive, info, member_path)
                        source_entry = next(
                            entry for entry in manifest
                            if entry.get("path") == member_name and "result_index" not in entry
                        )
                        add_result(_result_for_pdf(member_path, f"{name} / {member_name}"), source_entry)
                    processed_pdf_count += len(infos)
            except zipfile.BadZipFile:
                add_result({"file": name, "ok": False, "error": "压缩包损坏或格式不受支持"})
            except Exception as e:
                add_result({"file": name, "ok": False, "error": f"解析失败: {e}"})
            finally:
                sources.append(source)
                await f.close()
    except Exception:
        shutil.rmtree(job_directory, ignore_errors=True)
        raise

    ok = [r for r in results if r.get("ok") and r.get("total") is not None]
    grand = sum_money(r["total"] for r in ok) if ok else None
    n_low = sum(1 for r in results if r.get("ok") and r.get("confidence") == "low")
    preview_files = sum(1 for result in results if result.get("_preview_path"))
    job = {
        "job_id": job_id,
        "status": "queued",
        "directory": str(job_directory),
        "updated_at": time.time(),
        "grand_total": grand,
        "files": len(results),
        "ok_count": len(ok),
        "low_confidence_count": n_low,
        "sources": sources,
        "results": results,
        "preview": {
            "status": "queued",
            "total_files": preview_files,
            "completed_files": 0,
            "current_file": "",
            "current_file_index": 0,
            "current_page": 0,
            "current_pages": 0,
        },
    }
    with PREVIEW_JOBS_LOCK:
        PREVIEW_JOBS[job_id] = job
    background_tasks.add_task(_process_preview_job, job_id)
    return _preview_job_response(job, include_results=True)


def _cleanup_jira_jobs() -> None:
    """清理已完成且超过 TTL 的 Jira 任务及其导出文件。"""
    now = time.time()
    expired = []
    with JIRA_JOBS_LOCK:
        for job_id, job in JIRA_JOBS.items():
            if job.get("status") in {"done", "error"} and now - job.get("updated_at", now) > JIRA_JOB_TTL:
                expired.append((job_id, job.get("directory")))
        for job_id, _ in expired:
            JIRA_JOBS.pop(job_id, None)
    for _, directory in expired:
        if directory:
            shutil.rmtree(directory, ignore_errors=True)


def _public_jira_result(result: dict) -> dict:
    """页面只需要处理结果；原始矩阵留在服务端供导出使用。"""
    public = dict(result)
    public.pop("raw_rows", None)
    public.pop("raw_headers", None)
    return public


def _jira_job_response(job: dict, include_result: bool = False) -> dict:
    response = {
        "job_id": job["job_id"],
        "status": job["status"],
        "progress": dict(job["progress"]),
        "source": job.get("source", ""),
    }
    if job["status"] == "error":
        response["error"] = job.get("error", "Jira 文件处理失败")
    if include_result or job["status"] == "done":
        response["result"] = _public_jira_result(job["result"]) if job.get("result") else None
    return response


def _process_jira_job(job_id: str, path: Path, display_name: str) -> None:
    with JIRA_JOBS_LOCK:
        job = JIRA_JOBS.get(job_id)
        if not job:
            return
        job["status"] = "running"
        job["updated_at"] = time.time()

    def on_progress(update: dict) -> None:
        with JIRA_JOBS_LOCK:
            current = JIRA_JOBS.get(job_id)
            if current:
                current["progress"] = update
                current["updated_at"] = time.time()

    try:
        result = process_jira_workbook(path, progress_callback=on_progress)
        result["source"] = display_name
        with JIRA_JOBS_LOCK:
            job = JIRA_JOBS.get(job_id)
            if job:
                job["status"] = "done"
                job["result"] = result
                job["progress"] = {"stage": "完成", "percent": 100, "detail": "Jira 数据处理完成"}
                job["updated_at"] = time.time()
    except Exception as exc:
        with JIRA_JOBS_LOCK:
            job = JIRA_JOBS.get(job_id)
            if job:
                job["status"] = "error"
                job["error"] = str(exc)
                job["progress"] = {"stage": "失败", "percent": 100, "detail": "无法生成 Jira 处理结果"}
                job["updated_at"] = time.time()
    finally:
        path.unlink(missing_ok=True)


@app.post("/api/jira/import")
async def jira_import(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """上传 Jira 导出的 xlsx，异步执行 Word 中的清洗、筛选、排序和合并步骤。"""
    _cleanup_jira_jobs()
    display_name = _safe_name(file.filename or "jira-export.xlsx")
    if not display_name.lower().endswith(".xlsx"):
        await file.close()
        raise HTTPException(status_code=400, detail="目前仅支持 Jira 导出的 .xlsx 文件")

    JIRA_JOB_ROOT.mkdir(parents=True, exist_ok=True)
    job_id = uuid.uuid4().hex
    job_directory = Path(tempfile.mkdtemp(prefix=f"{job_id}_", dir=JIRA_JOB_ROOT))
    upload_path = job_directory / "source.xlsx"
    try:
        size, _ = await _save_upload(file, upload_path)
        if size > MAX_SIZE:
            raise HTTPException(status_code=413, detail="Excel 文件超过 20MB 限制")
        if size == 0:
            raise HTTPException(status_code=400, detail="Excel 文件内容为空")
    except HTTPException:
        shutil.rmtree(job_directory, ignore_errors=True)
        raise
    finally:
        await file.close()

    job = {
        "job_id": job_id,
        "status": "queued",
        "source": display_name,
        "directory": str(job_directory),
        "path": str(upload_path),
        "updated_at": time.time(),
        "progress": {"stage": "排队中", "percent": 0, "detail": "等待开始处理"},
        "result": None,
        "error": "",
    }
    with JIRA_JOBS_LOCK:
        JIRA_JOBS[job_id] = job
    background_tasks.add_task(_process_jira_job, job_id, upload_path, display_name)
    return _jira_job_response(job)


@app.get("/api/jira/jobs/{job_id}")
async def jira_job(job_id: str):
    _cleanup_jira_jobs()
    with JIRA_JOBS_LOCK:
        job = JIRA_JOBS.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Jira 处理任务不存在或已过期")
        return _jira_job_response(job)


@app.get("/api/jira/jobs/{job_id}/export")
async def jira_export(job_id: str):
    _cleanup_jira_jobs()
    with JIRA_JOBS_LOCK:
        job = JIRA_JOBS.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Jira 处理任务不存在或已过期")
        if job["status"] != "done" or not job.get("result"):
            raise HTTPException(status_code=409, detail="Jira 数据尚未处理完成")
        target = Path(job["directory"]) / "jira-report.xlsx"
        result = job["result"]
        source = job.get("source", "jira-export.xlsx")
    if not target.exists():
        export_jira_xlsx(result, target)
    return FileResponse(
        target,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"{Path(source).stem}-处理结果.xlsx",
    )


def _cleanup_weekly_jobs() -> None:
    """清理过期的周报解析材料和导出文件。"""
    now = time.time()
    expired = []
    with WEEKLY_JOBS_LOCK:
        for job_id, job in WEEKLY_JOBS.items():
            if job.get("status") in {"done", "error"} and now - job.get("updated_at", now) > WEEKLY_JOB_TTL:
                expired.append((job_id, job.get("directory")))
        for job_id, _ in expired:
            WEEKLY_JOBS.pop(job_id, None)
    for _, directory in expired:
        if directory:
            shutil.rmtree(directory, ignore_errors=True)


def _weekly_job_response(job: dict, include_result: bool = False) -> dict:
    response = {
        "job_id": job["job_id"],
        "status": job["status"],
        "progress": dict(job["progress"]),
        "source": job.get("source", ""),
    }
    if job["status"] == "error":
        response["error"] = job.get("error", "周报处理失败")
    if include_result or job["status"] == "done":
        response["result"] = job.get("result")
    return response


def _extract_weekly_archive(
    archive_path: Path,
    archive_name: str,
    target_directory: Path,
    upload_index: int,
) -> tuple[list[tuple[Path, str]], list[dict]]:
    """安全展开 ZIP 中的项目 PPTX，并保留完整目录清单。"""
    presentation_sources = []
    manifest = []
    expanded_size = 0
    supported_index = 0
    # 国内常见压缩软件生成的 ZIP 未设置 UTF-8 标记，文件名通常使用 GBK。
    with zipfile.ZipFile(archive_path, metadata_encoding="gbk") as archive:
        for info in archive.infolist():
            member_name = _safe_archive_member_name(info.filename)
            mode = (info.external_attr >> 16) & 0xFFFF
            if info.is_dir() or stat.S_ISDIR(mode):
                manifest.append({"path": f"{archive_name} / {member_name}", "kind": "directory", "status": "已扫描"})
                continue
            extension = Path(member_name).suffix.lower()
            kind = "ppt" if extension == ".pptx" else "historical" if extension == ".docx" else "ignored"
            entry = {
                "path": f"{archive_name} / {member_name}",
                "kind": kind,
                "size": info.file_size,
                "status": "待解析" if kind == "ppt" else "历史成品" if kind == "historical" else "已忽略",
            }
            manifest.append(entry)
            if kind != "ppt":
                continue
            if stat.S_ISLNK(mode):
                raise ValueError(f"压缩包包含不支持的软链接: {member_name}")
            if info.flag_bits & 0x1:
                raise ValueError(f"压缩包包含加密文件: {member_name}")
            if info.file_size > WEEKLY_MAX_SIZE:
                raise ValueError(f"压缩包内文件超过 500MB: {member_name}")
            expanded_size += info.file_size
            if expanded_size > WEEKLY_MAX_ARCHIVE_UNPACKED_SIZE:
                raise ValueError("周报压缩包展开后超过 2GB 限制")
            target = target_directory / f"archive_{upload_index}_{supported_index}{extension}"
            supported_index += 1
            _extract_archive_member(archive, info, target, WEEKLY_MAX_SIZE)
            display_name = f"{archive_name} / {member_name}"
            entry["status"] = "已发现"
            presentation_sources.append((target, display_name))
    return presentation_sources, manifest


def _process_weekly_job(job_id: str) -> None:
    with WEEKLY_JOBS_LOCK:
        job = WEEKLY_JOBS.get(job_id)
        if not job:
            return
        job["status"] = "running"
        job["updated_at"] = time.time()

    def on_progress(update: dict) -> None:
        with WEEKLY_JOBS_LOCK:
            current = WEEKLY_JOBS.get(job_id)
            if current:
                current["progress"] = update
                current["updated_at"] = time.time()

    try:
        result = process_weekly_report(
            [(Path(path), name) for path, name in job["presentation_sources"]],
            job["manifest"],
            progress_callback=on_progress,
        )
        directory = Path(job["directory"])
        source_lookup = {name: Path(path) for path, name in job["presentation_sources"]}
        output_stem = result["output_stem"]
        suffix = output_stem.removeprefix("项目周报")
        ppt_target = directory / f"{output_stem}.pptx"
        word_target = directory / f"部门周例会{suffix}.docx"
        report_target = directory / f"{output_stem}-审核报告.xlsx"
        zip_target = directory / f"{output_stem}-处理结果.zip"
        on_progress({"stage": "整合总周报", "percent": 91, "detail": f"正在生成 {ppt_target.name}"})
        build_weekly_presentation(result, source_lookup, ppt_target, progress_callback=on_progress)
        on_progress({"stage": "生成周例会", "percent": 96, "detail": f"正在填充 {word_target.name}"})
        build_weekly_meeting_document(result, word_target)
        export_weekly_report_xlsx(result, report_target)
        with zipfile.ZipFile(zip_target, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.write(ppt_target, ppt_target.name)
            archive.write(word_target, word_target.name)
            archive.write(report_target, report_target.name)
        with WEEKLY_JOBS_LOCK:
            current = WEEKLY_JOBS.get(job_id)
            if current:
                current["status"] = "done"
                current["result"] = result
                current["export_path"] = str(zip_target)
                current["progress"] = {"stage": "完成", "percent": 100, "detail": "周报已完成核验"}
                current["updated_at"] = time.time()
    except Exception as exc:
        with WEEKLY_JOBS_LOCK:
            current = WEEKLY_JOBS.get(job_id)
            if current:
                current["status"] = "error"
                current["error"] = str(exc)
                current["progress"] = {"stage": "失败", "percent": 100, "detail": "无法完成周报核验"}
                current["updated_at"] = time.time()


@app.post("/api/weekly-report/import")
async def weekly_report_import(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """上传部门项目周报 ZIP，异步审核并生成总 PPT 与周例会 Word。"""
    _cleanup_weekly_jobs()
    WEEKLY_JOB_ROOT.mkdir(parents=True, exist_ok=True)
    job_id = uuid.uuid4().hex
    job_directory = Path(tempfile.mkdtemp(prefix=f"{job_id}_", dir=WEEKLY_JOB_ROOT))
    display_name = _safe_name(file.filename or "部门项目周报.zip")
    if not display_name.lower().endswith(".zip"):
        await file.close()
        shutil.rmtree(job_directory, ignore_errors=True)
        raise HTTPException(status_code=400, detail="请上传包含部门所有项目周报的 ZIP 文件")
    upload_path = job_directory / "source.zip"
    try:
        size, _ = await _save_upload(file, upload_path, WEEKLY_MAX_SIZE)
        if size > WEEKLY_MAX_SIZE:
            raise HTTPException(status_code=413, detail="ZIP 文件超过 500MB 限制")
        if size == 0:
            raise HTTPException(status_code=400, detail="ZIP 文件内容为空")
        presentation_sources, archive_manifest = _extract_weekly_archive(
            upload_path, display_name, job_directory, 0
        )
        manifest = [{"path": display_name, "kind": "archive", "size": size, "status": "已扫描"}, *archive_manifest]
    except (HTTPException, zipfile.BadZipFile, ValueError) as exc:
        shutil.rmtree(job_directory, ignore_errors=True)
        if isinstance(exc, HTTPException):
            raise
        detail = "压缩包损坏或格式不受支持" if isinstance(exc, zipfile.BadZipFile) else str(exc)
        raise HTTPException(status_code=400, detail=detail) from exc
    finally:
        await file.close()
    if not presentation_sources:
        shutil.rmtree(job_directory, ignore_errors=True)
        raise HTTPException(status_code=400, detail="ZIP 中没有找到可处理的 PPTX 文件")
    if len(presentation_sources) > 50:
        shutil.rmtree(job_directory, ignore_errors=True)
        raise HTTPException(status_code=400, detail="单次最多解析 50 个 PPTX 文件")

    job = {
        "job_id": job_id,
        "status": "queued",
        "source": display_name,
        "directory": str(job_directory),
        "presentation_sources": [(str(path), name) for path, name in presentation_sources],
        "manifest": manifest,
        "updated_at": time.time(),
        "progress": {"stage": "排队中", "percent": 0, "detail": "等待开始审核项目周报"},
        "result": None,
        "error": "",
    }
    with WEEKLY_JOBS_LOCK:
        WEEKLY_JOBS[job_id] = job
    background_tasks.add_task(_process_weekly_job, job_id)
    return _weekly_job_response(job)


@app.get("/api/weekly-report/jobs/{job_id}")
async def weekly_report_job(job_id: str):
    _cleanup_weekly_jobs()
    with WEEKLY_JOBS_LOCK:
        job = WEEKLY_JOBS.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="周报任务不存在或已过期")
        return _weekly_job_response(job)


@app.get("/api/weekly-report/jobs/{job_id}/export")
async def weekly_report_export(job_id: str):
    _cleanup_weekly_jobs()
    with WEEKLY_JOBS_LOCK:
        job = WEEKLY_JOBS.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="周报任务不存在或已过期")
        if job["status"] != "done" or not job.get("result"):
            raise HTTPException(status_code=409, detail="周报数据尚未处理完成")
        result = job["result"]
        zip_target = Path(job.get("export_path", ""))
    if not zip_target.is_file():
        raise HTTPException(status_code=410, detail="周报结果文件不存在或已清理，请重新上传")
    return FileResponse(
        zip_target,
        media_type="application/zip",
        filename=zip_target.name,
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
