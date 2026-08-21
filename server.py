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

import asyncio
import base64
import stat
import shutil
import tempfile
import threading
import time
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from pathlib import PurePosixPath

import pymupdf as fitz
from fastapi import BackgroundTasks, Body, FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from invoice_total import Extractor, sum_money
from jira_processor import (
    export_daily_jira_xlsx,
    export_jira_xlsx,
    process_daily_jira_workbook,
    process_jira_workbook,
    discover_weekly_files,
    export_weekly_statistics_xlsx,
    process_weekly_statistics,
    process_new_task_statistics,
    process_screenshot_statistics,
    export_new_task_statistics_xlsx,
    export_combined_weekly_statistics_xlsx,
    preview_jira_workbook,
    validate_screenshot_metric,
)
from weekly_report_processor import (
    build_weekly_meeting_document,
    build_weekly_presentation,
    export_weekly_report_xlsx,
    process_weekly_report,
)
from reimbursement_generator import (
    parse_reimbursement_docx_many,
    render_reimbursement_pdf,
    validate_reimbursement_pdf,
)
try:
    from image_ppt_processor import build_image_presentation, validate_image_presentation
except ModuleNotFoundError:
    # 图片转 PPT 为可选功能，精简部署镜像中不包含其处理模块。
    IMAGE_PPT_AVAILABLE = False
else:
    IMAGE_PPT_AVAILABLE = True

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
WEEKLY_JIRA_JOB_TTL = 60 * 60
WEEKLY_JIRA_JOB_ROOT = Path(tempfile.gettempdir()) / "invoice_total_weekly_jira_jobs"
WEEKLY_JIRA_JOBS = {}
WEEKLY_JIRA_JOBS_LOCK = threading.Lock()
NEW_TASK_JOB_ROOT = Path(tempfile.gettempdir()) / "invoice_total_new_task_jobs"
NEW_TASK_JOBS = {}
NEW_TASK_JOBS_LOCK = threading.Lock()
WEEKLY_JOB_TTL = 60 * 60
WEEKLY_JOB_ROOT = Path(tempfile.gettempdir()) / "invoice_total_weekly_jobs"
WEEKLY_JOBS = {}
WEEKLY_JOBS_LOCK = threading.Lock()
IMAGE_PPT_JOB_TTL = 60 * 60
IMAGE_PPT_JOB_ROOT = Path(tempfile.gettempdir()) / "invoice_total_image_ppt_jobs"
IMAGE_PPT_JOBS = {}
IMAGE_PPT_JOBS_LOCK = threading.Lock()

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
    return FileResponse(STATIC / "jira.html", headers={"Cache-Control": "no-store, no-cache, must-revalidate"})


@app.get("/weekly-report")
async def weekly_report_index():
    return FileResponse(STATIC / "weekly-report.html")


@app.get("/reimbursement")
async def reimbursement_index():
    return FileResponse(STATIC / "reimbursement.html")


@app.get("/image-to-ppt")
async def image_to_ppt_index():
    if not IMAGE_PPT_AVAILABLE:
        raise HTTPException(status_code=404, detail="图片转 PPT 功能未部署")
    return FileResponse(STATIC / "image-to-ppt.html")


def _cleanup_image_ppt_jobs() -> None:
    now = time.time()
    expired = []
    with IMAGE_PPT_JOBS_LOCK:
        for job_id, job in IMAGE_PPT_JOBS.items():
            if now - job.get("updated_at", now) > IMAGE_PPT_JOB_TTL:
                expired.append((job_id, job.get("directory")))
        for job_id, _ in expired:
            IMAGE_PPT_JOBS.pop(job_id, None)
    for _, directory in expired:
        if directory:
            shutil.rmtree(directory, ignore_errors=True)


@app.post("/api/image-to-ppt/generate")
async def generate_image_presentation(file: UploadFile = File(...)):
    """Create a visually faithful one-slide PPTX from a raster reference image."""
    if not IMAGE_PPT_AVAILABLE:
        await file.close()
        raise HTTPException(status_code=503, detail="图片转 PPT 功能未部署")
    _cleanup_image_ppt_jobs()
    suffix = Path(file.filename or "").suffix.lower()
    if suffix not in {".png", ".jpg", ".jpeg", ".webp"}:
        await file.close()
        raise HTTPException(status_code=400, detail="请上传 PNG、JPG、JPEG 或 WEBP 图片")
    IMAGE_PPT_JOB_ROOT.mkdir(parents=True, exist_ok=True)
    job_id = uuid.uuid4().hex
    directory = Path(tempfile.mkdtemp(prefix=f"{job_id}_", dir=IMAGE_PPT_JOB_ROOT))
    source = directory / f"source{suffix}"
    normalised = directory / "source-normalised.png"
    output = directory / "图片转PPT.pptx"
    try:
        size, _ = await _save_upload(file, source, max_size=MAX_SIZE)
        if size == 0:
            raise HTTPException(status_code=400, detail="图片内容为空")
        if size > MAX_SIZE:
            raise HTTPException(status_code=413, detail="图片不能超过 20MB")
        expected = build_image_presentation(source, output, normalised)
        validation = validate_image_presentation(output, normalised, expected)
        if not validation["ok"]:
            raise HTTPException(status_code=422, detail=f"PPT 核验未通过：{'、'.join(validation['issues'])}")
        with IMAGE_PPT_JOBS_LOCK:
            IMAGE_PPT_JOBS[job_id] = {"directory": str(directory), "output": str(output), "updated_at": time.time()}
        return {"validation": validation, "download_url": f"/api/image-to-ppt/jobs/{job_id}/download"}
    except HTTPException:
        shutil.rmtree(directory, ignore_errors=True)
        raise
    except Exception as exc:
        shutil.rmtree(directory, ignore_errors=True)
        raise HTTPException(status_code=422, detail=f"无法处理图片：{exc}") from exc
    finally:
        await file.close()


@app.get("/api/image-to-ppt/jobs/{job_id}/download")
async def download_image_presentation(job_id: str):
    if not IMAGE_PPT_AVAILABLE:
        raise HTTPException(status_code=503, detail="图片转 PPT 功能未部署")
    _cleanup_image_ppt_jobs()
    with IMAGE_PPT_JOBS_LOCK:
        job = IMAGE_PPT_JOBS.get(job_id)
        output = Path(job.get("output", "")) if job else None
    if output is None or not output.is_file():
        raise HTTPException(status_code=404, detail="生成结果不存在或已清理，请重新上传图片")
    return FileResponse(output, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", filename="图片转PPT.pptx")


@app.post("/api/reimbursement/generate")
async def generate_reimbursement(file: UploadFile = File(...)):
    """Create a reimbursement PDF from the exported DOCX label/value document."""
    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="请上传 DOCX 格式的报销信息文件")
    directory = Path(tempfile.mkdtemp(prefix="reimbursement_"))
    try:
        source = directory / "source.docx"
        size, _ = await _save_upload(file, source, max_size=10 * 1024 * 1024)
        if size > 10 * 1024 * 1024:
            raise HTTPException(status_code=413, detail="报销信息文件不能超过 10MB")
        try:
            reimbursements = parse_reimbursement_docx_many(source)
            if not reimbursements:
                raise ValueError("未识别到报销单内容")
            generated_at = datetime.now()
            outputs: list[Path] = []
            validations = []
            for index, reimbursement in enumerate(reimbursements, start=1):
                claimant = reimbursement.fields.get("claimant", "")
                if not claimant:
                    raise ValueError(f"第 {index} 张报销单未识别到“报销人”字段")
                safe_claimant = "".join(char for char in claimant if char not in '\\/:*?\"<>|').strip()
                number = reimbursement.fields.get("reimbursement_number", "")
                safe_number = "".join(char for char in number if char not in '\\/:*?\"<>|').strip()
                if not safe_claimant or not safe_number:
                    raise ValueError(f"第 {index} 张报销单的报销人或报销编号不能作为文件名")
                output = directory / f"报销{safe_claimant}-{safe_number}.pdf"
                render_reimbursement_pdf(reimbursement, output, generated_at=generated_at)
                validation = validate_reimbursement_pdf(reimbursement, output, generated_at)
                if not validation["ok"]:
                    raise ValueError(f"第 {index} 张报销单整体核验未通过：{'、'.join(validation['errors'])}")
                outputs.append(output)
                validations.append(validation)
        except ValueError as exc:
            raise HTTPException(status_code=422, detail=f"无法解析报销信息：{exc}") from exc
        except Exception as exc:
            raise HTTPException(status_code=422, detail=f"生成报销单失败：{exc}") from exc
        cleanup = BackgroundTasks()
        cleanup.add_task(shutil.rmtree, directory, ignore_errors=True)
        if len(outputs) == 1:
            response = FileResponse(
                outputs[0],
                media_type="application/pdf",
                filename=outputs[0].name,
                background=cleanup,
            )
        else:
            archive = directory / "报销单.pdf.zip"
            with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as bundle:
                for output in outputs:
                    bundle.write(output, output.name)
            response = FileResponse(
                archive,
                media_type="application/zip",
                filename="报销单.zip",
                background=cleanup,
            )
        response.headers["X-Reimbursement-Validation"] = "passed"
        response.headers["X-Reimbursement-Validation-Checks"] = str(sum(len(item["checks"]) for item in validations))
        response.headers["X-Reimbursement-Count"] = str(len(outputs))
        return response
    except Exception:
        shutil.rmtree(directory, ignore_errors=True)
        raise


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


def _safe_upload_name(raw_name: str) -> str:
    """保留目录上传的安全相对路径，普通上传仍显示文件名。"""
    try:
        return _safe_archive_member_name(raw_name)
    except ValueError:
        return _safe_name(raw_name)


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
            name = _safe_upload_name(raw_name)
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


def _cleanup_weekly_jira_jobs() -> None:
    now = time.time()
    expired = []
    with WEEKLY_JIRA_JOBS_LOCK:
        for job_id, job in WEEKLY_JIRA_JOBS.items():
            if job.get("status") in {"done", "error"} and now - job.get("updated_at", now) > WEEKLY_JIRA_JOB_TTL:
                expired.append((job_id, job.get("directory")))
        for job_id, _ in expired:
            WEEKLY_JIRA_JOBS.pop(job_id, None)
    for _, directory in expired:
        if directory:
            shutil.rmtree(directory, ignore_errors=True)


def _weekly_jira_job_response(job: dict) -> dict:
    response = {"job_id": job["job_id"], "status": job["status"], "progress": dict(job["progress"]), "source": job.get("source", "")}
    if job["status"] == "error":
        response["error"] = job.get("error", "周报 Jira 处理失败")
    if job["status"] == "done":
        response["result"] = job.get("result")
    return response


def _extract_weekly_jira_zip(archive_path: Path, target_dir: Path) -> list[Path]:
    paths = []
    expanded_size = 0
    with zipfile.ZipFile(archive_path) as archive:
        for info in archive.infolist():
            if info.is_dir() or not info.filename.lower().endswith(".xlsx"):
                continue
            member_name = _safe_archive_member_name(info.filename)
            if Path(member_name).name.startswith("~$"):
                continue
            if info.flag_bits & 0x1:
                raise ValueError(f"压缩包包含加密文件: {member_name}")
            if info.file_size > MAX_SIZE:
                raise ValueError(f"压缩包内 Excel 超过 20MB: {member_name}")
            expanded_size += info.file_size
            if expanded_size > MAX_ARCHIVE_UNPACKED_SIZE:
                raise ValueError("周报压缩包展开后超过 200MB 限制")
            target = target_dir.joinpath(*PurePosixPath(member_name).parts)
            target.parent.mkdir(parents=True, exist_ok=True)
            _extract_archive_member(archive, info, target, MAX_SIZE)
            paths.append(target)
    return paths


def _process_weekly_jira_job(job_id: str) -> None:
    with WEEKLY_JIRA_JOBS_LOCK:
        job = WEEKLY_JIRA_JOBS.get(job_id)
        if not job:
            return
        job["status"] = "running"
        job["updated_at"] = time.time()

    def on_progress(update: dict) -> None:
        with WEEKLY_JIRA_JOBS_LOCK:
            current = WEEKLY_JIRA_JOBS.get(job_id)
            if current:
                current["progress"] = update
                current["updated_at"] = time.time()

    try:
        sources = discover_weekly_files([Path(item) for item in job["input_paths"]])
        result = process_weekly_statistics(sources, progress_callback=on_progress)
        target = Path(job["directory"]) / "weekly-jira-statistics.xlsx"
        export_weekly_statistics_xlsx(result, target)
        result["source_files"] = {key: path.name for key, path in sources.items()}
        with WEEKLY_JIRA_JOBS_LOCK:
            current = WEEKLY_JIRA_JOBS.get(job_id)
            if current:
                current.update({"status": "done", "result": result, "output": str(target), "progress": {"stage": "完成", "percent": 100, "detail": "当前周 Jira 统计已完成"}, "updated_at": time.time()})
    except Exception as exc:
        with WEEKLY_JIRA_JOBS_LOCK:
            current = WEEKLY_JIRA_JOBS.get(job_id)
            if current:
                current.update({"status": "error", "error": str(exc), "progress": {"stage": "失败", "percent": 100, "detail": "无法生成当前周 Jira 统计"}, "updated_at": time.time()})


@app.post("/api/jira/weekly-statistics/import")
async def weekly_jira_statistics_import(background_tasks: BackgroundTasks, files: list[UploadFile] = File(...)):
    """接收当前周 Jira ZIP 或目录文件列表。"""
    _cleanup_weekly_jira_jobs()
    if not files:
        raise HTTPException(status_code=400, detail="请上传周报 ZIP 或选择包含 Jira Excel 的目录")
    job_id = uuid.uuid4().hex
    WEEKLY_JIRA_JOB_ROOT.mkdir(parents=True, exist_ok=True)
    job_directory = Path(tempfile.mkdtemp(prefix=f"{job_id}_", dir=WEEKLY_JIRA_JOB_ROOT))
    input_dir = job_directory / "inputs"
    input_dir.mkdir()
    paths = []
    try:
        for index, upload in enumerate(files):
            name = _safe_upload_name(upload.filename or f"upload-{index}.xlsx")
            path = input_dir / name
            if path.exists():
                path = input_dir / f"{index:03d}-{Path(name).name}"
            path.parent.mkdir(parents=True, exist_ok=True)
            size, _ = await _save_upload(upload, path, WEEKLY_MAX_SIZE)
            if size > WEEKLY_MAX_SIZE:
                raise HTTPException(status_code=413, detail="周报上传文件超过 500MB 限制")
            if size == 0:
                continue
            if path.suffix.lower() == ".zip":
                extracted = _extract_weekly_jira_zip(path, input_dir / f"zip-{index}")
                paths.extend(extracted)
                path.unlink(missing_ok=True)
            elif path.suffix.lower() == ".xlsx":
                paths.append(path)
            await upload.close()
        if not paths:
            raise HTTPException(status_code=400, detail="没有找到有效的周报 Excel 文件")
        discover_weekly_files(paths)
    except (HTTPException, ValueError) as exc:
        shutil.rmtree(job_directory, ignore_errors=True)
        if isinstance(exc, HTTPException):
            raise
        raise HTTPException(status_code=400, detail=str(exc))
    finally:
        for upload in files:
            await upload.close()
    job = {"job_id": job_id, "status": "queued", "source": "周报 Jira ZIP/目录", "directory": str(job_directory), "input_paths": [str(path) for path in paths], "progress": {"stage": "排队中", "percent": 0, "detail": "等待处理当前周 Jira 数据"}, "result": None, "error": "", "updated_at": time.time()}
    with WEEKLY_JIRA_JOBS_LOCK:
        WEEKLY_JIRA_JOBS[job_id] = job
    background_tasks.add_task(_process_weekly_jira_job, job_id)
    return _weekly_jira_job_response(job)


@app.get("/api/jira/weekly-statistics/jobs/{job_id}")
async def weekly_jira_statistics_job(job_id: str):
    _cleanup_weekly_jira_jobs()
    with WEEKLY_JIRA_JOBS_LOCK:
        job = WEEKLY_JIRA_JOBS.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="周报 Jira 任务不存在或已过期")
        return _weekly_jira_job_response(job)


@app.get("/api/jira/weekly-statistics/jobs/{job_id}/export")
async def weekly_jira_statistics_export(job_id: str):
    _cleanup_weekly_jira_jobs()
    with WEEKLY_JIRA_JOBS_LOCK:
        job = WEEKLY_JIRA_JOBS.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="周报 Jira 任务不存在或已过期")
        if job["status"] != "done":
            raise HTTPException(status_code=409, detail="周报 Jira 数据尚未处理完成")
        target = Path(job["output"])
    if not target.exists():
        raise HTTPException(status_code=410, detail="周报统计文件不存在或已清理")
    return FileResponse(target, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="周报Jira统计结果.xlsx")


def _process_jira_job(job_id: str, path: Path, display_name: str, mode: str = "weekly") -> None:
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
        processor = process_daily_jira_workbook if mode == "daily" else process_jira_workbook
        result = processor(path, progress_callback=on_progress)
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


def _process_new_task_job(job_id: str) -> None:
    with NEW_TASK_JOBS_LOCK:
        job = NEW_TASK_JOBS.get(job_id)
        if not job:
            return
        job["status"] = "running"
        job["progress"] = {"stage": "正在识别截图", "percent": 25, "detail": "正在识别统计标题和合计数量"}
    def on_progress(update: dict) -> None:
        with NEW_TASK_JOBS_LOCK:
            current = NEW_TASK_JOBS.get(job_id)
            if current:
                current["progress"] = update

    try:
        def on_progress(update: dict) -> None:
            with NEW_TASK_JOBS_LOCK:
                current = NEW_TASK_JOBS.get(job_id)
                if current:
                    current["progress"] = update
        stat_key = job.get("stat_key", "new-tasks")
        metric_map = {
            "new-tasks": "本周新增任务数", "completed-tasks": "本周完成任务数",
            "new-defects": "本周新增缺陷数", "fixed-defects": "本周已修复缺陷数",
            "delayed-defects": "本周延期缺陷数", "pending-defects": "总挂起缺陷数",
            "delayed-tasks": "本周延期任务数", "pending-tasks": "总挂起任务数",
        }
        metric = metric_map.get(stat_key, metric_map["new-tasks"])
        result = process_new_task_statistics(job["xlsx"], job["screenshot"], metric=metric, progress_callback=on_progress) if stat_key in {"new-tasks", "completed-tasks"} else process_screenshot_statistics(job["screenshot"], metric, progress_callback=on_progress)
        with NEW_TASK_JOBS_LOCK:
            job["progress"] = {"stage": "正在生成 Excel", "percent": 85, "detail": f"正在整理{metric}统计结果"}
        target = Path(job["directory"]) / f"{metric}.xlsx"
        export_new_task_statistics_xlsx(result, target)
        with NEW_TASK_JOBS_LOCK:
            job.update({"status": "done", "result": result, "output": str(target), "progress": {"stage": "完成", "percent": 100, "detail": f"{metric}统计已完成"}})
    except Exception as exc:
        with NEW_TASK_JOBS_LOCK:
            job.update({"status": "error", "error": str(exc), "progress": {"stage": "失败", "percent": 100, "detail": f"无法生成{metric}统计"}})


async def _process_new_task_job_async(job_id: str) -> None:
    """Run OCR/Excel work off the event loop so job polling remains responsive."""
    await asyncio.to_thread(_process_new_task_job, job_id)


@app.post("/api/jira/weekly-new-tasks/import")
async def weekly_new_tasks_import(background_tasks: BackgroundTasks, jira_file: UploadFile = File(...), screenshot: UploadFile = File(...), stat_key: str = "new-tasks"):
    """Upload a Jira export and the assignee-total screenshot for new-task correction."""
    if not (jira_file.filename or "").lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="请上传 Jira .xlsx 文件")
    if Path(screenshot.filename or "").suffix.lower() not in {".png", ".jpg", ".jpeg", ".webp"}:
        raise HTTPException(status_code=400, detail="请上传 PNG、JPG、JPEG 或 WEBP 截图")
    NEW_TASK_JOB_ROOT.mkdir(parents=True, exist_ok=True)
    job_id = uuid.uuid4().hex
    directory = Path(tempfile.mkdtemp(prefix=f"{job_id}_", dir=NEW_TASK_JOB_ROOT))
    xlsx_path, image_path = directory / "source.xlsx", directory / f"source{Path(screenshot.filename).suffix.lower()}"
    try:
        xlsx_size, _ = await _save_upload(jira_file, xlsx_path, MAX_SIZE)
        image_size, _ = await _save_upload(screenshot, image_path, MAX_SIZE)
        if not xlsx_size or not image_size:
            raise HTTPException(status_code=400, detail="上传文件内容为空")
    except HTTPException:
        shutil.rmtree(directory, ignore_errors=True)
        raise
    finally:
        await jira_file.close()
        await screenshot.close()
    allowed = {"new-tasks", "completed-tasks"}
    if stat_key not in allowed:
        shutil.rmtree(directory, ignore_errors=True)
        raise HTTPException(status_code=400, detail="该统计项需要上传截图，不需要 Jira Excel")
    job = {"job_id": job_id, "status": "queued", "stat_key": stat_key, "directory": str(directory), "xlsx": str(xlsx_path), "screenshot": str(image_path), "output": "", "result": None, "error": "", "progress": {"stage": "排队中", "percent": 0, "detail": "等待开始处理"}}
    with NEW_TASK_JOBS_LOCK:
        NEW_TASK_JOBS[job_id] = job
    background_tasks.add_task(_process_new_task_job_async, job_id)
    return {"job_id": job_id, "status": "queued", "progress": job["progress"]}


@app.get("/api/jira/weekly-new-tasks/jobs/{job_id}")
async def weekly_new_tasks_job(job_id: str):
    with NEW_TASK_JOBS_LOCK:
        job = NEW_TASK_JOBS.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="统计任务不存在或已过期")
        response = {"job_id": job_id, "status": job["status"], "progress": job["progress"]}
        if job["status"] == "done":
            response["result"] = job["result"]
        if job["status"] == "error":
            response["error"] = job["error"]
        return response


@app.get("/api/jira/weekly-new-tasks/jobs/{job_id}/export")
async def weekly_new_tasks_export(job_id: str):
    with NEW_TASK_JOBS_LOCK:
        job = NEW_TASK_JOBS.get(job_id)
        if not job or job["status"] != "done":
            raise HTTPException(status_code=409, detail="统计结果尚未生成")
        target = Path(job["output"])
        metric = {"new-tasks": "本周新增任务数", "completed-tasks": "本周完成任务数", "new-defects": "本周新增缺陷数", "fixed-defects": "本周已修复缺陷数", "delayed-defects": "本周延期缺陷数", "pending-defects": "总挂起缺陷数", "delayed-tasks": "本周延期任务数", "pending-tasks": "总挂起任务数"}.get(job.get("stat_key", "new-tasks"), "周报统计")
    return FileResponse(target, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=f"{metric}.xlsx")


@app.post("/api/jira/weekly-summary/export")
async def weekly_summary_export(payload: dict = Body(default_factory=dict)):
    """Merge completed weekly-stat jobs from their server-side results.

    The browser only keeps files and rendered results in memory. Use its completed
    job IDs as the source of truth so a UI state refresh cannot create an empty
    workbook after an individual statistic has already been generated.
    """
    job_ids = payload.get("job_ids", {})
    stat_keys = payload.get("stat_keys", [])
    if not isinstance(job_ids, dict):
        raise HTTPException(status_code=400, detail="整合任务参数格式不正确")
    if not isinstance(stat_keys, list):
        raise HTTPException(status_code=400, detail="整合统计项参数格式不正确")
    statistics: dict[str, list[dict]] = {}
    with NEW_TASK_JOBS_LOCK:
        for stat_key, job_id in job_ids.items():
            job = NEW_TASK_JOBS.get(str(job_id))
            if not job or job.get("status") != "done" or job.get("stat_key") != stat_key:
                continue
            result = job.get("result") or {}
            summary = result.get("summary")
            if isinstance(summary, list):
                statistics[stat_key] = summary
        # Compatibility for a page that visually completed before it started
        # retaining the task ID on the action button. The newest completed job
        # for that visible statistic is the same result the user just generated.
        for stat_key in stat_keys:
            if stat_key in statistics:
                continue
            job = next((item for item in reversed(list(NEW_TASK_JOBS.values())) if item.get("status") == "done" and item.get("stat_key") == stat_key), None)
            summary = (job or {}).get("result", {}).get("summary")
            if isinstance(summary, list):
                statistics[stat_key] = summary
    if not statistics:
        raise HTTPException(status_code=409, detail="没有可整合的已完成统计，请先生成至少一项统计文件")
    directory = Path(tempfile.mkdtemp(prefix="weekly_summary_", dir=NEW_TASK_JOB_ROOT))
    target = directory / "Jira周报统计汇总.xlsx"
    export_combined_weekly_statistics_xlsx(statistics, target)
    return FileResponse(target, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="Jira周报统计汇总.xlsx")


@app.post("/api/jira/workbook-preview")
async def jira_workbook_preview(file: UploadFile = File(...)):
    if not (file.filename or "").lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="请上传 Jira .xlsx 文件")
    directory = Path(tempfile.mkdtemp(prefix="jira_preview_", dir=NEW_TASK_JOB_ROOT))
    target = directory / "preview.xlsx"
    try:
        size, _ = await _save_upload(file, target, MAX_SIZE)
        if not size:
            raise HTTPException(status_code=400, detail="Excel 文件内容为空")
        return preview_jira_workbook(target)
    finally:
        await file.close()
        shutil.rmtree(directory, ignore_errors=True)


@app.post("/api/jira/screenshot-title-check/{stat_key}")
async def screenshot_title_check(stat_key: str, screenshot: UploadFile = File(...)):
    metric_map = {"new-tasks": "本周新增任务数", "completed-tasks": "本周完成任务数", "new-defects": "本周新增缺陷数", "fixed-defects": "本周已修复缺陷数", "delayed-defects": "本周延期缺陷数", "pending-defects": "总挂起缺陷数", "delayed-tasks": "本周延期任务数", "pending-tasks": "总挂起任务数"}
    metric = metric_map.get(stat_key)
    if not metric:
        raise HTTPException(status_code=400, detail="不支持的统计项")
    if Path(screenshot.filename or "").suffix.lower() not in {".png", ".jpg", ".jpeg", ".webp"}:
        raise HTTPException(status_code=400, detail="请粘贴 PNG、JPG、JPEG 或 WEBP 截图")
    directory = Path(tempfile.mkdtemp(prefix="screenshot_check_", dir=NEW_TASK_JOB_ROOT))
    target = directory / f"source{Path(screenshot.filename).suffix.lower()}"
    try:
        size, _ = await _save_upload(screenshot, target, MAX_SIZE)
        if not size:
            raise HTTPException(status_code=400, detail="截图内容为空")
        await asyncio.to_thread(validate_screenshot_metric, target, metric)
        return {"ok": True, "metric": metric}
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    finally:
        await screenshot.close()
        shutil.rmtree(directory, ignore_errors=True)


@app.post("/api/jira/weekly-screenshot-stats/{stat_key}/import")
async def weekly_screenshot_stats_import(stat_key: str, background_tasks: BackgroundTasks, screenshot: UploadFile = File(...)):
    allowed = {"new-defects", "fixed-defects", "delayed-defects", "pending-defects", "delayed-tasks", "pending-tasks"}
    if stat_key not in allowed:
        raise HTTPException(status_code=400, detail="不支持的统计项")
    if Path(screenshot.filename or "").suffix.lower() not in {".png", ".jpg", ".jpeg", ".webp"}:
        raise HTTPException(status_code=400, detail="请上传 PNG、JPG、JPEG 或 WEBP 截图")
    metric_map = {"new-defects": "本周新增缺陷数", "fixed-defects": "本周已修复缺陷数", "delayed-defects": "本周延期缺陷数", "pending-defects": "总挂起缺陷数", "delayed-tasks": "本周延期任务数", "pending-tasks": "总挂起任务数"}
    root = NEW_TASK_JOB_ROOT / "screenshot-stats"
    root.mkdir(parents=True, exist_ok=True)
    job_id = uuid.uuid4().hex
    directory = Path(tempfile.mkdtemp(prefix=f"{job_id}_", dir=root))
    image_path = directory / f"source{Path(screenshot.filename).suffix.lower()}"
    try:
        size, _ = await _save_upload(screenshot, image_path, MAX_SIZE)
        if not size:
            raise HTTPException(status_code=400, detail="上传截图内容为空")
    finally:
        await screenshot.close()
    job = {"job_id": job_id, "status": "queued", "stat_key": stat_key, "directory": str(directory), "xlsx": "", "screenshot": str(image_path), "output": "", "result": None, "error": "", "progress": {"stage": "排队中", "percent": 0, "detail": "等待开始处理"}}
    with NEW_TASK_JOBS_LOCK:
        NEW_TASK_JOBS[job_id] = job
    background_tasks.add_task(_process_new_task_job_async, job_id)
    return {"job_id": job_id, "status": "queued", "progress": job["progress"]}


@app.post("/api/jira/import")
async def jira_import(background_tasks: BackgroundTasks, file: UploadFile = File(...), mode: str = "weekly"):
    """上传 Jira 导出的 xlsx，异步执行 Word 中的清洗、筛选、排序和合并步骤。"""
    _cleanup_jira_jobs()
    if mode not in {"weekly", "daily"}:
        raise HTTPException(status_code=400, detail="不支持的 Jira 处理模式")
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
        "mode": mode,
        "directory": str(job_directory),
        "path": str(upload_path),
        "updated_at": time.time(),
        "progress": {"stage": "排队中", "percent": 0, "detail": "等待开始处理"},
        "result": None,
        "error": "",
    }
    with JIRA_JOBS_LOCK:
        JIRA_JOBS[job_id] = job
    background_tasks.add_task(_process_jira_job, job_id, upload_path, display_name, mode)
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
        is_daily = job.get("mode") == "daily"
        target = Path(job["directory"]) / ("daily-jira-report.xlsx" if is_daily else "jira-report.xlsx")
        result = job["result"]
        source = job.get("source", "jira-export.xlsx")
    if not target.exists():
        exporter = export_daily_jira_xlsx if is_daily else export_jira_xlsx
        exporter(result, target)
    return FileResponse(
        target,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"{Path(source).stem}-{'每日Jira' if is_daily else '周报Jira'}处理结果.xlsx",
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


async def _collect_weekly_folder(uploads: list[UploadFile], target_directory: Path, source_name: str) -> tuple[list[tuple[Path, str]], list[dict], int]:
    """Save browser directory uploads and return PPTX sources plus a manifest."""
    presentation_sources, manifest = [], []
    total_size = 0
    supported_index = 0
    for upload in uploads:
        relative_name = _safe_upload_name(upload.filename or "unnamed")
        extension = Path(relative_name).suffix.lower()
        kind = "ppt" if extension == ".pptx" else "historical" if extension == ".docx" else "ignored"
        target = target_directory / f"folder_{supported_index}{extension or '.bin'}"
        size, _ = await _save_upload(upload, target, WEEKLY_MAX_SIZE)
        total_size += size
        if size > WEEKLY_MAX_SIZE:
            raise ValueError(f"File exceeds 500MB limit: {relative_name}")
        if total_size > WEEKLY_MAX_ARCHIVE_UNPACKED_SIZE:
            raise ValueError("Folder upload exceeds 2GB limit")
        entry = {"path": f"{source_name} / {relative_name}", "kind": kind, "size": size, "status": "待解析" if kind == "ppt" else "历史成品" if kind == "historical" else "已忽略"}
        manifest.append(entry)
        if kind == "ppt":
            entry["status"] = "已发现"
            presentation_sources.append((target, f"{source_name} / {relative_name}"))
            supported_index += 1
        else:
            target.unlink(missing_ok=True)
    return presentation_sources, manifest, total_size


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
                current["export_paths"] = {
                    "ppt": str(ppt_target),
                    "docx": str(word_target),
                    "xlsx": str(report_target),
                    "zip": str(zip_target),
                }
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
async def weekly_report_import(background_tasks: BackgroundTasks, file: UploadFile | None = File(None), files: list[UploadFile] | None = File(None)):
    """上传部门项目周报 ZIP 或文件夹内容，异步审核并生成结果。"""
    _cleanup_weekly_jobs()
    uploads = ([file] if file else []) + (files or [])
    if not uploads:
        raise HTTPException(status_code=400, detail="请选择 ZIP 文件或周报文件夹")
    WEEKLY_JOB_ROOT.mkdir(parents=True, exist_ok=True)
    job_id = uuid.uuid4().hex
    job_directory = Path(tempfile.mkdtemp(prefix=f"{job_id}_", dir=WEEKLY_JOB_ROOT))
    display_name = _safe_name(uploads[0].filename or "部门项目周报.zip")
    is_archive = len(uploads) == 1 and display_name.lower().endswith(".zip")
    try:
        if is_archive:
            upload_path = job_directory / "source.zip"
            size, _ = await _save_upload(uploads[0], upload_path, WEEKLY_MAX_SIZE)
            if size > WEEKLY_MAX_SIZE:
                raise HTTPException(status_code=413, detail="ZIP 文件超过 500MB 限制")
            if size == 0:
                raise HTTPException(status_code=400, detail="ZIP 文件内容为空")
            presentation_sources, archive_manifest = _extract_weekly_archive(upload_path, display_name, job_directory, 0)
            manifest = [{"path": display_name, "kind": "archive", "size": size, "status": "已扫描"}, *archive_manifest]
        else:
            display_name = "文件夹上传"
            presentation_sources, folder_manifest, size = await _collect_weekly_folder(uploads, job_directory, display_name)
            manifest = [{"path": display_name, "kind": "directory", "size": size, "status": "已扫描"}, *folder_manifest]
    except (HTTPException, zipfile.BadZipFile, ValueError) as exc:
        shutil.rmtree(job_directory, ignore_errors=True)
        if isinstance(exc, HTTPException):
            raise
        detail = "压缩包损坏或格式不受支持" if isinstance(exc, zipfile.BadZipFile) else str(exc)
        raise HTTPException(status_code=400, detail=detail) from exc
    finally:
        for upload in uploads:
            await upload.close()
    if not presentation_sources:
        shutil.rmtree(job_directory, ignore_errors=True)
        raise HTTPException(status_code=400, detail="上传内容中没有找到可处理的 PPTX 文件")
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


@app.get("/api/weekly-report/jobs/{job_id}/export/{file_kind}")
async def weekly_report_file_export(job_id: str, file_kind: str):
    """下载周报任务生成的单个文件。"""
    media_types = {
        "ppt": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }
    if file_kind not in media_types:
        raise HTTPException(status_code=404, detail="不支持的周报导出文件类型")
    _cleanup_weekly_jobs()
    with WEEKLY_JOBS_LOCK:
        job = WEEKLY_JOBS.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="周报任务不存在或已过期")
        if job["status"] != "done" or not job.get("result"):
            raise HTTPException(status_code=409, detail="周报数据尚未处理完成")
        target = Path(job.get("export_paths", {}).get(file_kind, ""))
    if not target.is_file():
        raise HTTPException(status_code=410, detail="周报结果文件不存在或已清理，请重新上传")
    return FileResponse(target, media_type=media_types[file_kind], filename=target.name)


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
