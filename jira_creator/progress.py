"""Small progress snapshots shared with the isolated OCR process."""

import json
import sqlite3
import time


def report_progress(database, job_id, detail, completed=0, total=0):
    if not database or not job_id:
        return
    progress = {"detail": detail, "completed": completed, "total": total,
                "percent": round(completed * 100 / total, 1) if total else None}
    db = sqlite3.connect(str(database), timeout=5)
    try:
        with db:
            db.execute("UPDATE jobs SET result=?,updated_at=? WHERE id=? AND status='running'",
                       (json.dumps({"progress": progress}, ensure_ascii=False), time.time(), job_id))
    finally:
        db.close()
