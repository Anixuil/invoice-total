#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
server.py — 发票金额提取 Web 服务（FastAPI）

启动: uvicorn server:app --host 0.0.0.0 --port 8000
API:
  GET  /             前端页面
  POST /api/upload   多文件上传 PDF，返回逐份报告 + 总金额
"""

import tempfile
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from invoice_total import Extractor

app = FastAPI(title="发票金额提取", version="1.0.0")

BASE = Path(__file__).resolve().parent
STATIC = BASE / "static"
MAX_SIZE = 20 * 1024 * 1024  # 20MB

app.mount("/static", StaticFiles(directory=STATIC), name="static")


@app.get("/")
async def index():
    return FileResponse(STATIC / "index.html")


@app.post("/api/upload")
async def upload(files: list[UploadFile] = File(...)):
    if not files:
        raise HTTPException(status_code=400, detail="未收到文件")

    results = []
    with tempfile.TemporaryDirectory(prefix="inv_") as td:
        for f in files:
            name = f.filename or "unnamed.pdf"
            if not name.lower().endswith(".pdf"):
                results.append({"file": name, "ok": False, "error": "仅支持 PDF 文件"})
                continue
            data = await f.read()
            if len(data) > MAX_SIZE:
                results.append({"file": name, "ok": False, "error": "文件超过 20MB 限制"})
                continue
            path = Path(td) / name
            path.write_bytes(data)
            ex = Extractor(path)
            try:
                ex.load()
                r = ex.run()
                r["file"] = name
            except Exception as e:
                r = {"file": name, "ok": False, "error": f"解析失败: {e}"}
            results.append(r)

    ok = [r for r in results if r.get("ok") and r.get("total") is not None]
    grand = round(sum(r["total"] for r in ok), 2) if ok else None
    n_low = sum(1 for r in results if r.get("ok") and r.get("confidence") == "low")
    return {
        "grand_total": grand,
        "files": len(results),
        "ok_count": len(ok),
        "low_confidence_count": n_low,
        "results": results,
    }


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
