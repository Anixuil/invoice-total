#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Jira 导出数据处理：把 Word 中的手工表格步骤变成可追溯的规则流水线。"""

from __future__ import annotations

from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
import re
from typing import Any, Callable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


ProgressCallback = Callable[[dict[str, Any]], None] | None

STATUS_MAP = {
    "已解决": "已完成",
    "已关闭": "已完成",
    "处理中": "进行中",
}
KNOWN_STATUSES = set(STATUS_MAP) | {"开放", "已完成", "进行中"}

# 来源文档中的人员顺序。实际数据中没有任务的人员也保留在汇总中，方便周报直接复制。
DEFAULT_ROSTER = [
    "linyuping",
    "liuxin01",
    "wuyutian",
    "caoyuan",
    "zhengleyuan",
    "douhuanhuan",
    "xiaozhennan",
    "sunyanqiang",
    "linyu01",
    "lilixin",
    "fanshuyu",
    "yinwei",
    "wangruixia",
    "yuhaicheng",
    "caiyuanpeng",
    "ruishunzi",
    "guanhaojie",
    "lihongqi",
    "liwenbo",
    "mojunyou",
    "chenyu",
    "muzhengyi",
]

WEEKLY_ROSTER = [
    ("zhoujiayu", "周佳宇"), ("zhengleyuan", "郑乐园"), ("douhuanhuan", "窦欢欢"),
    ("sunyanqiang", "孙艳强"), ("lilixin", "李礼辛"), ("linyu01", "林榆"),
    ("fanshuyu", "范书毓"), ("xiaozhennan", "肖镇楠"), ("wangruixia", "王瑞霞"),
    ("yinwei", "尹维"), ("yuhaicheng", "余海城"), ("muzhengyi", "穆政逸"),
    ("ruishunzi", "芮顺子"), ("caiyuanpeng", "蔡远鹏"), ("guanhaojie", "官浩杰"),
    ("lihongqi", "李红旗"), ("liwenbo", "李文波"), ("chenyu", "陈愉"),
    ("mojunyou", "莫钧友"), ("linyuping", "林宇萍"), ("wuyutian", "吴宇天"),
    ("caoyuan", "曹原"), ("liuxin01", "刘昕"),
]
WEEKLY_NAME_MAP = dict(WEEKLY_ROSTER, **{"liuying": "刘颖"})
WEEKLY_FILE_KEYS = {
    "new_tasks": "本周新增任务", "completed_tasks": "本周已完成任务",
    "new_defects": "本周新增缺陷", "fixed_defects": "本周已修复缺陷",
    "pending_tasks": "待处理任务", "pending_defects": "总挂起缺陷数",
}

FIELD_ALIASES = {
    "issue_key": ["问题关键字", "问题键", "issuekey", "key"],
    "issue_id": ["问题id", "issueid"],
    "parent_id": ["父级id", "父问题id", "parentid"],
    "status": ["状态", "status"],
    "created_at": ["创建日期", "创建时间", "created", "createddate"],
    "summary": ["概要", "标题", "summary"],
    "assignee": ["经办人", "负责人", "assignee"],
    "reporter": ["报告人", "reporter"],
    "delivery_date": ["自定义字段交付日期", "交付日期", "deliverydate"],
    "resolution": ["解决结果", "resolution"],
    "issue_type": ["问题类型", "类型", "issuetype"],
    "developer": ["自定义字段开发人员", "开发人员", "developer"],
    "planned_completion": ["自定义字段计划完成日期", "计划完成日期", "plannedcompletion"],
}


def _notify(callback: ProgressCallback, stage: str, percent: int, detail: str) -> None:
    if callback:
        callback({"stage": stage, "percent": max(0, min(100, percent)), "detail": detail})


def _norm_header(value: Any) -> str:
    text = str(value or "").strip().lower()
    return re.sub(r"[\s_\-（）()：:]+", "", text)


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _date_text(value: Any) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M") if value.time() != datetime.min.time() else value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, (int, float)):
        try:
            converted = datetime(1899, 12, 30) + timedelta(days=float(value))
            return converted.strftime("%Y-%m-%d %H:%M") if converted.time() != datetime.min.time() else converted.strftime("%Y-%m-%d")
        except (OverflowError, ValueError):
            return _text(value)
    return _text(value)


def _excel_safe(value: Any) -> Any:
    """避免把用户输入中的 =、+、-、@ 当作 Excel 公式执行。"""
    if not isinstance(value, str):
        return value
    return "'" + value if value[:1] in {"=", "+", "-", "@"} else value


def _column_name(index: int) -> str:
    return get_column_letter(index + 1)


def _find_index(headers: list[str], aliases: list[str]) -> int | None:
    normalized = [_norm_header(header) for header in headers]
    candidates = {_norm_header(alias) for alias in aliases}
    for index, value in enumerate(normalized):
        if value in candidates:
            return index
    return None


def _module_indexes(headers: list[str]) -> list[int]:
    indexes = []
    for index, header in enumerate(headers):
        if _norm_header(header).startswith("模块"):
            indexes.append(index)
    return indexes


def _field_mapping(headers: list[str], indexes: dict[str, int | None], module_indexes: list[int]) -> dict[str, dict[str, Any]]:
    mapping: dict[str, dict[str, Any]] = {}
    for field, index in indexes.items():
        mapping[field] = {
            "header": headers[index] if index is not None and index < len(headers) else "",
            "column": _column_name(index) if index is not None else "",
            "found": index is not None,
        }
    mapping["modules"] = {
        "header": "、".join(headers[index] for index in module_indexes),
        "column": ", ".join(_column_name(index) for index in module_indexes),
        "found": bool(module_indexes),
    }
    return mapping


def process_jira_workbook(path: str | Path, progress_callback: ProgressCallback = None) -> dict[str, Any]:
    """读取 Jira xlsx，并返回页面预览与导出共用的 JSON 结果。"""
    _notify(progress_callback, "读取工作表", 8, "正在打开 Excel 工作簿")
    workbook = load_workbook(filename=path, read_only=True, data_only=True)
    try:
        worksheet = next((sheet for sheet in workbook.worksheets if sheet.max_row), None)
        if worksheet is None:
            raise ValueError("工作簿中没有可读取的数据")

        rows = worksheet.iter_rows()
        header_cells = next(rows, None)
        if not header_cells:
            raise ValueError("工作表为空")
        populated_header_indexes = [index for index, cell in enumerate(header_cells) if _text(cell.value)]
        if not populated_header_indexes:
            raise ValueError("首行没有可识别的字段名")
        header_cells = header_cells[: populated_header_indexes[-1] + 1]
        headers = [_text(cell.value) or f"列{index + 1}" for index, cell in enumerate(header_cells)]
        # 同名 Jira 字段保留全部列，避免第二个“模块”覆盖第一个。
        seen: Counter[str] = Counter()
        unique_headers = []
        for header in headers:
            seen[header] += 1
            unique_headers.append(header if seen[header] == 1 else f"{header}_{seen[header]}")

        indexes = {field: _find_index(headers, aliases) for field, aliases in FIELD_ALIASES.items()}
        module_indexes = _module_indexes(headers)
        missing = [field for field in ("status", "summary", "assignee") if indexes[field] is None]
        if missing:
            labels = {"status": "状态", "summary": "概要", "assignee": "经办人"}
            raise ValueError("缺少必需字段：" + "、".join(labels[field] for field in missing))

        mapping = _field_mapping(unique_headers, indexes, module_indexes)
        _notify(progress_callback, "识别字段", 18, f"已识别 {sum(item['found'] for item in mapping.values())} 个字段")

        records: list[dict[str, Any]] = []
        raw_rows: list[list[Any]] = []
        for row_number, cells in enumerate(rows, start=2):
            values = [cell.value for cell in cells]
            if not any(value not in (None, "") for value in values):
                continue
            raw_rows.append([_date_text(value) if isinstance(value, (date, datetime)) else _text(value) for value in values])
            get = lambda field: values[indexes[field]] if indexes[field] is not None and indexes[field] < len(values) else None
            issue_key = _text(get("issue_key")) or f"第{row_number}行"
            status_raw = _text(get("status"))
            summary = _text(get("summary"))
            assignee = _text(get("assignee"))
            developer = _text(get("developer"))
            reporter = _text(get("reporter"))
            if developer:
                summary_person = developer
                summary_source = "开发人员"
                summary_note = ""
            elif assignee:
                summary_person = assignee
                summary_source = "经办人"
                summary_note = "开发人员为空，按经办人汇总"
            elif reporter:
                summary_person = reporter
                summary_source = "报告人"
                summary_note = "开发人员和经办人均为空，按报告人汇总"
            else:
                summary_person = "未分配"
                summary_source = "未分配"
                summary_note = "开发人员、经办人和报告人均为空，归入未分配"
            modules = [_text(values[index]) for index in module_indexes if index < len(values) and _text(values[index])]
            consistency = "待补开发人员" if not developer else "一致" if assignee == developer else "人员不一致"
            display_status = STATUS_MAP.get(status_raw, status_raw or "未填写")
            # 汇总归属依次使用开发人员、经办人、报告人；均为空时归入未分配。
            # 人员字段不一致只做异常提示，不影响按上述优先级归入汇总。
            record = {
                "source_row": row_number,
                "issue_key": issue_key,
                "issue_id": _text(get("issue_id")),
                "parent_id": _text(get("parent_id")),
                "status_raw": status_raw,
                "status": display_status,
                "created_at": _date_text(get("created_at")),
                "summary": summary,
                "assignee": assignee or "未分配",
                "reporter": reporter,
                "developer": developer,
                "summary_person": summary_person,
                "summary_source": summary_source,
                "summary_note": summary_note,
                "modules": modules,
                "module": "、".join(dict.fromkeys(modules)),
                "delivery_date": _date_text(get("delivery_date")),
                "resolution": _text(get("resolution")),
                "issue_type": _text(get("issue_type")),
                "planned_completion": _date_text(get("planned_completion")),
                "consistency": consistency,
                "included": True,
                "summary_text": f"{summary}【{display_status}】",
            }
            records.append(record)

        _notify(progress_callback, "清洗数据", 42, f"已读取 {len(records)} 条任务")
        key_counts = Counter(record["issue_key"] for record in records)
        roster_order = {name: index + 1 for index, name in enumerate(DEFAULT_ROSTER)}
        for record in records:
            record["sort_order"] = roster_order.get(record["summary_person"], 9999)
            record["duplicate"] = key_counts[record["issue_key"]] > 1

        def sort_key(record: dict[str, Any]) -> tuple[Any, ...]:
            return (
                record["sort_order"],
                record["planned_completion"] or "9999-99-99",
                record["delivery_date"] or "9999-99-99",
                record["source_row"],
            )

        records.sort(key=sort_key)
        included = [record for record in records if record["included"]]
        grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for record in included:
            grouped[record["summary_person"]].append(record)

        summaries = []
        summary_names = DEFAULT_ROSTER + sorted(name for name in grouped if name not in roster_order)
        for name in summary_names:
            items = grouped.get(name, [])
            status_count = Counter(item["status"] for item in items)
            summaries.append({
                "name": name,
                "task_count": len(items),
                "completed_count": status_count.get("已完成", 0),
                "in_progress_count": status_count.get("进行中", 0),
                "open_count": status_count.get("开放", 0),
                "merged_text": ";\n".join(item["summary_text"] for item in items),
                "items": [item["issue_key"] for item in items],
            })

        anomalies = []
        for record in records:
            if record["consistency"] == "待补开发人员":
                anomalies.append({"row": record["source_row"], "issue_key": record["issue_key"], "type": "developer_missing", "label": "待补开发人员", "detail": record["summary_note"]})
            elif record["consistency"] == "人员不一致":
                anomalies.append({"row": record["source_row"], "issue_key": record["issue_key"], "type": "developer_mismatch", "label": "人员不一致", "detail": f"经办人：{record['assignee']}；开发人员：{record['developer']}；已按开发人员纳入人员汇总"})
            if record["status_raw"] not in KNOWN_STATUSES:
                anomalies.append({"row": record["source_row"], "issue_key": record["issue_key"], "type": "unknown_status", "label": "未知状态", "detail": record["status_raw"] or "状态为空"})
            if not record["planned_completion"]:
                anomalies.append({"row": record["source_row"], "issue_key": record["issue_key"], "type": "planned_date_missing", "label": "计划完成日期缺失", "detail": "原始字段为空"})
            if record["duplicate"]:
                anomalies.append({"row": record["source_row"], "issue_key": record["issue_key"], "type": "duplicate_issue", "label": "问题关键字重复", "detail": "建议回到 Jira 导出数据确认是否重复"})

        status_counts = Counter(record["status_raw"] for record in records)
        _notify(progress_callback, "核验并分组", 78, f"按人员归属汇总 {len(included)} 条，异常 {len(anomalies)} 条")
        result = {
            "ok": True,
            "source": Path(path).name,
            "sheet": worksheet.title,
            "field_mapping": mapping,
            "stats": {
                "total_rows": len(records),
                "included_rows": len(included),
                "pending_developer": sum(record["consistency"] == "待补开发人员" for record in records),
                "mismatched_developer": sum(record["consistency"] == "人员不一致" for record in records),
                "unique_assignees": len({record["summary_person"] for record in records}),
                "duplicate_keys": sum(count > 1 for count in key_counts.values()),
                "missing_planned_completion": sum(not record["planned_completion"] for record in records),
                "status_counts": dict(status_counts),
            },
            "summaries": summaries,
            "records": records,
            "anomalies": anomalies,
            "raw_headers": unique_headers,
            "raw_rows": raw_rows,
        }
        _notify(progress_callback, "生成结果", 100, "Jira 数据处理完成")
        return result
    finally:
        workbook.close()


def process_daily_jira_workbook(path: str | Path, progress_callback: ProgressCallback = None) -> dict[str, Any]:
    """按开发人员、经办人、报告人的优先级生成每日 Jira 汇总。"""
    _notify(progress_callback, "读取工作表", 8, "正在打开 Excel 工作簿")
    workbook = load_workbook(filename=path, read_only=True, data_only=True)
    try:
        worksheet = next((sheet for sheet in workbook.worksheets if sheet.max_row), None)
        if worksheet is None:
            raise ValueError("工作簿中没有可读取的数据")

        rows = worksheet.iter_rows()
        header_cells = next(rows, None)
        if not header_cells:
            raise ValueError("工作表为空")
        populated_header_indexes = [index for index, cell in enumerate(header_cells) if _text(cell.value)]
        if not populated_header_indexes:
            raise ValueError("首行没有可识别的字段名")
        header_cells = header_cells[: populated_header_indexes[-1] + 1]
        headers = [_text(cell.value) or f"列{index + 1}" for index, cell in enumerate(header_cells)]
        seen: Counter[str] = Counter()
        unique_headers = []
        for header in headers:
            seen[header] += 1
            unique_headers.append(header if seen[header] == 1 else f"{header}_{seen[header]}")

        indexes = {field: _find_index(headers, aliases) for field, aliases in FIELD_ALIASES.items()}
        missing = [field for field in ("status", "summary") if indexes[field] is None]
        if missing:
            labels = {"status": "状态", "summary": "概要"}
            raise ValueError("缺少每日处理必需字段：" + "、".join(labels[field] for field in missing))
        if all(indexes[field] is None for field in ("developer", "assignee", "reporter")):
            raise ValueError("缺少每日处理人员字段：开发人员、经办人或报告人")

        mapping = _field_mapping(unique_headers, indexes, _module_indexes(headers))
        _notify(progress_callback, "识别字段", 18, "已识别状态、概要和人员归属字段")

        records: list[dict[str, Any]] = []
        for row_number, cells in enumerate(rows, start=2):
            values = [cell.value for cell in cells]
            if not any(value not in (None, "") for value in values):
                continue
            get = lambda field: values[indexes[field]] if indexes[field] is not None and indexes[field] < len(values) else None
            assignee = _text(get("assignee"))
            developer = _text(get("developer"))
            reporter = _text(get("reporter"))
            status_raw = _text(get("status"))
            status = STATUS_MAP.get(status_raw, status_raw or "未填写")
            summary = _text(get("summary"))
            if developer:
                summary_person = developer
                summary_source = "开发人员"
                consistency = "相同" if not assignee or assignee == developer else "不同"
                summary_note = "" if consistency == "相同" else "经办人与开发人员不一致，按开发人员汇总"
            elif assignee:
                summary_person = assignee
                summary_source = "经办人"
                consistency = "缺少开发人员"
                summary_note = "开发人员为空，按经办人汇总"
            elif reporter:
                summary_person = reporter
                summary_source = "报告人"
                consistency = "缺少开发人员和经办人"
                summary_note = "开发人员和经办人均为空，按报告人汇总"
            else:
                summary_person = "未分配"
                summary_source = "未分配"
                consistency = "人员缺失"
                summary_note = "开发人员、经办人和报告人均为空，未纳入每日汇总"
            included = bool(developer or assignee or reporter)
            records.append({
                "source_row": row_number,
                "issue_key": _text(get("issue_key")) or f"第{row_number}行",
                "status_raw": status_raw,
                "status": status,
                "summary": summary,
                "assignee": assignee,
                "reporter": reporter,
                "developer": developer,
                "summary_person": summary_person,
                "summary_source": summary_source,
                "summary_note": summary_note,
                "module": "",
                "planned_completion": _date_text(get("planned_completion")),
                "consistency": consistency,
                "included": included,
                "summary_text": f"{summary}【{status}】",
            })

        _notify(progress_callback, "确定人员归属", 48, f"已读取 {len(records)} 条任务")
        roster_order = {name: index + 1 for index, name in enumerate(DEFAULT_ROSTER)}
        records.sort(key=lambda record: (
            0 if record["included"] else 1,
            roster_order.get(record["summary_person"], len(DEFAULT_ROSTER) + 1),
            record["source_row"],
        ))
        included_records = [record for record in records if record["included"]]
        grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for record in included_records:
            grouped[record["summary_person"]].append(record)

        summary_names = DEFAULT_ROSTER + sorted(name for name in grouped if name not in roster_order)
        summaries = []
        for name in summary_names:
            items = grouped.get(name, [])
            status_count = Counter(item["status"] for item in items)
            summaries.append({
                "name": name,
                "task_count": len(items),
                "completed_count": status_count.get("已完成", 0),
                "in_progress_count": status_count.get("进行中", 0),
                "open_count": status_count.get("开放", 0),
                "merged_text": ";\n".join(item["summary_text"] for item in items),
                "items": [item["issue_key"] for item in items],
            })

        anomalies = [{
            "row": record["source_row"],
            "issue_key": record["issue_key"],
            "type": "daily_person_fallback" if record["included"] else "daily_person_missing",
            "label": record["consistency"],
            "detail": record["summary_note"],
        } for record in records if record["summary_note"]]
        _notify(progress_callback, "排序并合并", 82, f"已按人员优先级纳入 {len(included_records)} 条记录，按 {len(summaries)} 人生成汇总")
        result = {
            "ok": True,
            "mode": "daily",
            "source": Path(path).name,
            "sheet": worksheet.title,
            "field_mapping": mapping,
            "stats": {
                "total_rows": len(records),
                "included_rows": len(included_records),
                "excluded_rows": len(records) - len(included_records),
                "pending_developer": sum(not record["developer"] for record in records),
                "mismatched_developer": sum(record["consistency"] == "不同" for record in records),
                "unique_assignees": len(grouped),
                "status_counts": dict(Counter(record["status_raw"] for record in records)),
            },
            "summaries": summaries,
            "records": records,
            "anomalies": anomalies,
        }
        _notify(progress_callback, "生成结果", 100, "每日 Jira 数据处理完成")
        return result
    finally:
        workbook.close()


def discover_weekly_files(paths: list[str | Path]) -> dict[str, Path]:
    """按 Jira 导出文件名识别周报六类来源，忽略 Excel 临时锁定文件。"""
    matches: dict[str, list[Path]] = {key: [] for key in WEEKLY_FILE_KEYS}
    for raw_path in paths:
        path = Path(raw_path)
        if path.name.startswith("~$") or path.suffix.lower() != ".xlsx":
            continue
        for key, marker in WEEKLY_FILE_KEYS.items():
            if marker in path.name:
                matches[key].append(path)
                break
    conflicts = {key: items for key, items in matches.items() if len(items) > 1}
    if conflicts:
        detail = "；".join(f"{key}: {', '.join(item.name for item in items)}" for key, items in conflicts.items())
        raise ValueError(f"周报文件存在多个候选文件：{detail}")
    missing = [key for key, items in matches.items() if not items]
    if missing:
        labels = "、".join(WEEKLY_FILE_KEYS[key] for key in missing)
        raise ValueError(f"缺少周报来源文件：{labels}")
    return {key: items[0] for key, items in matches.items()}


def _read_weekly_records(path: str | Path) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    workbook = load_workbook(filename=path, read_only=True, data_only=True)
    try:
        worksheet = next((sheet for sheet in workbook.worksheets if sheet.max_row), None)
        if worksheet is None:
            raise ValueError(f"工作簿为空：{Path(path).name}")
        rows = worksheet.iter_rows()
        header_cells = next(rows, None)
        if not header_cells:
            raise ValueError(f"工作表为空：{Path(path).name}")
        last = max((index for index, cell in enumerate(header_cells) if _text(cell.value)), default=-1)
        if last < 0:
            raise ValueError(f"首行没有字段名：{Path(path).name}")
        headers = [_text(cell.value) or f"列{index + 1}" for index, cell in enumerate(header_cells[:last + 1])]
        indexes = {field: _find_index(headers, aliases) for field, aliases in FIELD_ALIASES.items()}
        missing = [field for field in ("issue_key", "status", "summary") if indexes[field] is None]
        if missing:
            labels = {"issue_key": "问题关键字", "status": "状态", "summary": "概要"}
            raise ValueError(f"{Path(path).name} 缺少必需字段：{'、'.join(labels[field] for field in missing)}")
        module_indexes = _module_indexes(headers)
        records = []
        for row_number, cells in enumerate(rows, start=2):
            values = [cell.value for cell in cells]
            if not any(value not in (None, "") for value in values):
                continue
            get = lambda field: values[indexes[field]] if indexes[field] is not None and indexes[field] < len(values) else None
            status_raw = _text(get("status"))
            resolution = _text(get("resolution"))
            issue_type = _text(get("issue_type"))
            assignee = _text(get("assignee"))
            developer = _text(get("developer"))
            reporter = _text(get("reporter"))
            records.append({
                "source_row": row_number, "issue_key": _text(get("issue_key")) or f"第{row_number}行",
                "issue_id": _text(get("issue_id")), "status": status_raw, "summary": _text(get("summary")),
                "created_at": get("created_at"), "assignee": assignee, "reporter": reporter,
                "developer": developer, "module": "、".join(dict.fromkeys(_text(values[i]) for i in module_indexes if i < len(values) and _text(values[i]))),
                "delivery_date": get("delivery_date"), "resolution": resolution, "issue_type": issue_type,
                "planned_completion": get("planned_completion"),
            })
        return records, {"file": Path(path).name, "sheet": worksheet.title, "headers": headers}
    finally:
        workbook.close()


def _is_defect(record: dict[str, Any]) -> bool:
    return _text(record.get("issue_type")).lower() in {"故障", "缺陷", "bug", "defect"}


def _is_completed(record: dict[str, Any]) -> bool:
    status = _text(record.get("status"))
    resolution = _text(record.get("resolution"))
    return bool(resolution) and (status in {"已解决", "已关闭", "已完成"} or resolution in {"已解决", "已完成", "完成"})


def _date_value(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def _weekly_owners(record: dict[str, Any]) -> tuple[list[str], str]:
    assignee = _text(record.get("assignee"))
    developer = _text(record.get("developer"))
    reporter = _text(record.get("reporter"))
    if assignee and developer and assignee != developer:
        return list(dict.fromkeys([assignee, developer])), "经办人与开发人员不一致，按经办人基础统计并追加开发人员"
    if assignee:
        return [assignee], "开发人员为空，按经办人统计" if not developer else "经办人与开发人员一致"
    if developer:
        return [developer], "经办人为空，按开发人员统计"
    if reporter:
        return [reporter], "经办人和开发人员为空，按报告人兜底"
    return ["未分配"], "经办人、开发人员和报告人均为空"


def _weekly_display_person(value: Any) -> str:
    """将 Jira 用户名转换为周报明细中的中文姓名。"""
    text = _text(value)
    return WEEKLY_NAME_MAP.get(text, text)


def process_weekly_statistics(sources: dict[str, str | Path], progress_callback: ProgressCallback = None) -> dict[str, Any]:
    """处理当前周六类 Jira 导出，生成 23 人周报统计及明细。"""
    all_records: dict[str, list[dict[str, Any]]] = {}
    metadata = {}
    for index, (key, path) in enumerate(sources.items(), start=1):
        _notify(progress_callback, "读取周报文件", index * 10, f"正在读取 {Path(path).name}")
        all_records[key], metadata[key] = _read_weekly_records(path)
    # 已完成任务导出通常不带人员和问题类型，按问题关键字回填新增任务中的归属字段。
    new_task_lookup = {record["issue_key"]: record for record in all_records["new_tasks"]}
    for record in all_records["completed_tasks"]:
        reference = new_task_lookup.get(record["issue_key"])
        if reference:
            for field in ("assignee", "developer", "reporter", "issue_type", "delivery_date", "module"):
                if not record.get(field):
                    record[field] = reference.get(field)
    created_dates = [_date_value(item.get("created_at")) for rows in all_records.values() for item in rows]
    report_date = max((value for value in created_dates if value), default=date.today())
    roster = {username: {"序号": index, "姓名": name, "username": username} for index, (username, name) in enumerate(WEEKLY_ROSTER, start=1)}
    summary = [{**person, **{key: 0 for key in ("new_task_count", "completed_task_count", "new_defect_count", "fixed_defect_count", "delayed_defect_count", "pending_defect_count", "delayed_task_count", "pending_task_count")}} for person in roster.values()]
    by_username = {item["username"]: item for item in summary}
    anomalies: list[dict[str, Any]] = []
    sections = {"delayed_defects": [], "pending_defects": [], "delayed_tasks": [], "pending_tasks": []}

    def add_count(record: dict[str, Any], field: str, reason: str) -> None:
        owners, note = _weekly_owners(record)
        record = dict(record, owner_note=note, source_field=field)
        for owner in owners:
            target = by_username.get(owner)
            if target:
                target[field] += 1
            else:
                anomalies.append({"source": field, "issue_key": record["issue_key"], "label": "未知人员", "detail": owner})
        if note != "经办人与开发人员一致":
            anomalies.append({"source": field, "issue_key": record["issue_key"], "label": "人员归属备注", "detail": note})
        if not record.get("resolution"):
            anomalies.append({"source": field, "issue_key": record["issue_key"], "label": "解决结果为空", "detail": "未计入完成或修复统计"})
        return record

    for record in all_records["new_tasks"]:
        if not _is_defect(record):
            add_count(record, "new_task_count", "新增任务")
    for record in all_records["completed_tasks"]:
        if not _is_defect(record) and _is_completed(record):
            add_count(record, "completed_task_count", "完成任务")
    for record in all_records["new_defects"]:
        if _is_defect(record):
            add_count(record, "new_defect_count", "新增缺陷")
    for record in all_records["fixed_defects"]:
        if _is_defect(record) and _is_completed(record):
            add_count(record, "fixed_defect_count", "修复缺陷")

    def classify_pending(key: str, defect_field: str, task_field: str, section_key: str, require_overdue: bool = False) -> None:
        for record in all_records[key]:
            defect = _is_defect(record)
            if require_overdue:
                delivery = _date_value(record.get("delivery_date"))
                if not delivery:
                    anomalies.append({"source": key, "issue_key": record["issue_key"], "label": "交付日期为空", "detail": "无法判断逾期"})
                    continue
                if delivery >= report_date or _is_completed(record):
                    continue
            elif _is_completed(record):
                continue
            if (defect and defect_field) or (not defect and task_field):
                field = defect_field if defect else task_field
                detail = add_count(record, field, section_key)
                sections[section_key].append(detail)

    classify_pending("new_defects", "delayed_defect_count", "", "delayed_defects", True)
    classify_pending("pending_defects", "pending_defect_count", "", "pending_defects")
    classify_pending("pending_tasks", "", "delayed_task_count", "delayed_tasks", True)
    classify_pending("pending_tasks", "", "pending_task_count", "pending_tasks")
    return {"ok": True, "report_date": report_date.isoformat(), "sources": metadata, "summary": summary, "sections": sections, "anomalies": anomalies, "stats": {"total_rows": sum(len(rows) for rows in all_records.values()), "anomaly_count": len(anomalies)}}


def export_weekly_statistics_xlsx(result: dict[str, Any], target: str | Path) -> None:
    """导出当前周统计、延期挂起明细、异常和处理说明。"""
    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "周报统计"
    headers = ["序号", "姓名", "本周新增任务数", "本周完成任务数", "本周新增缺陷数", "本周已修复缺陷数", "本周延期缺陷数", "总挂起缺陷数", "本周延期任务数", "总挂起任务数"]
    summary_sheet.append(headers)
    summary_sheet.row_dimensions[1].height = 30
    header_fill = PatternFill("solid", fgColor="2F5597")
    header_font = Font(color="FFFFFF", bold=True)
    red_font = Font(color="FF0000")
    for cell in summary_sheet[1]:
        cell.fill, cell.font = header_fill, header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fields = ["new_task_count", "completed_task_count", "new_defect_count", "fixed_defect_count", "delayed_defect_count", "pending_defect_count", "delayed_task_count", "pending_task_count"]
    for item in result["summary"]:
        values = [item["序号"], item["姓名"], item["new_task_count"], item["completed_task_count"], item["new_defect_count"], item["fixed_defect_count"]]
        values.extend(item[field] or None for field in fields[4:])
        summary_sheet.append(values)
        summary_sheet.row_dimensions[summary_sheet.max_row].height = 22
    total_row = summary_sheet.max_row + 1
    summary_sheet.cell(total_row, 1, "合计")
    summary_sheet.row_dimensions[total_row].height = 24
    for column in range(3, 11):
        summary_sheet.cell(total_row, column, f"=SUM({get_column_letter(column)}2:{get_column_letter(column)}{total_row - 1})")
    thin = Side(style="thin", color="808080")
    for row in summary_sheet.iter_rows():
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in range(2, total_row):
        for column in (7, 8, 9, 10):
            if summary_sheet.cell(row, column).value:
                summary_sheet.cell(row, column).font = red_font
    for index, width in enumerate([8, 12, 16, 16, 16, 16, 16, 16, 16, 16], start=1):
        summary_sheet.column_dimensions[get_column_letter(index)].width = width
    summary_sheet.freeze_panes = "A2"
    summary_sheet.sheet_view.showGridLines = True
    summary_sheet.page_setup.orientation = "landscape"
    summary_sheet.page_setup.fitToWidth = 1
    summary_sheet.sheet_properties.pageSetUpPr.fitToPage = True

    detail_sheet = summary_sheet
    detail_headers = ["问题关键字", "状态", "创建日期", "概要", "经办人", "报告人", "模块", "交付日期", "解决结果", "问题类型", "开发人员", "计划完成日期"]
    detail_row = 1
    section_titles = [("delayed_defects", "本周延期缺陷数"), ("pending_defects", "总挂起缺陷数"), ("delayed_tasks", "本周延期任务数")]
    detail_row = total_row + 2
    for section_key, title in section_titles:
        detail_sheet.merge_cells(start_row=detail_row, start_column=1, end_row=detail_row, end_column=len(detail_headers))
        title_cell = detail_sheet.cell(detail_row, 1, title)
        title_cell.fill = PatternFill("solid", fgColor="D9E2F3")
        title_cell.font = Font(bold=True)
        title_cell.alignment = Alignment(horizontal="center")
        detail_sheet.row_dimensions[detail_row].height = 28
        detail_row += 1
        for column, value in enumerate(detail_headers, start=1):
            detail_sheet.cell(detail_row, column, value)
        for cell in detail_sheet[detail_row]:
            cell.fill = PatternFill(fill_type=None)
            cell.font = Font(color="000000", bold=False)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        detail_sheet.row_dimensions[detail_row].height = 32
        detail_row += 1
        for record in result["sections"].get(section_key, []):
            values = [record.get(field) for field in ["issue_key", "status", "created_at", "summary", "assignee", "reporter", "module", "delivery_date", "resolution", "issue_type", "developer", "planned_completion"]]
            values[4] = _weekly_display_person(values[4])
            values[5] = _weekly_display_person(values[5])
            values[10] = _weekly_display_person(values[10])
            for column, value in enumerate(values, start=1):
                detail_sheet.cell(detail_row, column, value)
            summary_lines = max(1, (len(str(record.get("summary") or "")) + 28) // 29)
            detail_sheet.row_dimensions[detail_row].height = min(22 * summary_lines, 82)
            detail_row += 1
        detail_row += 1
    for row in detail_sheet.iter_rows(min_row=total_row + 2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            if cell.row > total_row + 2 and cell.column in (3, 8, 12) and cell.value:
                cell.number_format = "yyyy/m/d h:mm"
    for section_offset in range(len(section_titles)):
        section_title_row = total_row + 2
        for previous_key, _ in section_titles[:section_offset]:
            section_title_row += 3 + len(result["sections"].get(previous_key, []))
        title_cell = detail_sheet.cell(section_title_row, 1)
        title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for index, width in enumerate([18, 12, 20, 42, 16, 16, 24, 20, 16, 14, 16, 20], start=1):
        detail_sheet.column_dimensions[get_column_letter(index)].width = width

    anomaly_sheet = workbook.create_sheet("异常清单")
    anomaly_sheet.append(["来源", "问题关键字", "异常类型", "说明"])
    for item in result["anomalies"]:
        anomaly_sheet.append([item.get("source", ""), item.get("issue_key", ""), item.get("label", ""), item.get("detail", "")])
    notes_sheet = workbook.create_sheet("处理说明")
    notes_sheet.append(["项目", "内容"])
    notes_sheet.append(["报告日期", result.get("report_date", "")])
    notes_sheet.append(["完成规则", "解决结果为空不计入完成或修复；状态与解决结果均明确完成时才计入。"])
    notes_sheet.append(["逾期规则", "交付日期早于报告日期且未完成，视为逾期。"])
    notes_sheet.append(["人员规则", "经办人与开发人员不一致时按经办人基础统计并追加开发人员；开发人员为空按经办人。"])
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        if sheet.title != "周报统计":
            sheet.auto_filter.ref = sheet.dimensions
    workbook.save(target)
    workbook.close()


def export_jira_xlsx(result: dict[str, Any], target: str | Path) -> None:
    """导出人员汇总、任务明细、异常清单和字段映射。"""
    workbook = Workbook()
    workbook.remove(workbook.active)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center")

    def add_sheet(title: str, headers: list[str], rows: list[list[Any]]) -> None:
        sheet = workbook.create_sheet(title)
        sheet.append([_excel_safe(value) for value in headers])
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
        for row in rows:
            sheet.append([_excel_safe(value) for value in row])
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column_cells in sheet.columns:
            width = min(max(max(len(str(cell.value or "")) for cell in column_cells) + 2, 10), 42)
            sheet.column_dimensions[column_cells[0].column_letter].width = width
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        if title == "人员汇总":
            sheet.column_dimensions["B"].width = 70
            for row_index, row in enumerate(sheet.iter_rows(min_row=2), start=2):
                row[1].alignment = Alignment(vertical="top", wrap_text=True)
                line_count = max(1, str(row[1].value or "").count("\n") + 1)
                sheet.row_dimensions[row_index].height = min(18 * line_count, 120)

    add_sheet(
        "人员汇总",
        ["人员", "合并任务内容"],
        [[item["name"], item["merged_text"]] for item in result["summaries"]],
    )
    add_sheet(
        "任务明细",
        ["来源行", "问题关键字", "问题ID", "父级ID", "原始状态", "展示状态", "创建日期", "概要", "经办人", "报告人", "开发人员", "汇总人员", "汇总依据", "备注", "模块", "交付日期", "解决结果", "问题类型", "计划完成日期", "一致性", "是否纳入汇总"],
        [[record[field] if field != "是否纳入汇总" else ("是" if record["included"] else "否") for field in ["source_row", "issue_key", "issue_id", "parent_id", "status_raw", "status", "created_at", "summary", "assignee", "reporter", "developer", "summary_person", "summary_source", "summary_note", "module", "delivery_date", "resolution", "issue_type", "planned_completion", "consistency", "是否纳入汇总"]] for record in result["records"]],
    )
    add_sheet(
        "异常清单",
        ["来源行", "问题关键字", "异常类型", "说明"],
        [[item["row"], item["issue_key"], item["label"], item["detail"]] for item in result["anomalies"]],
    )
    add_sheet(
        "字段映射",
        ["标准字段", "原始表头", "列", "是否识别"],
        [[field, details["header"], details["column"], "是" if details["found"] else "否"] for field, details in result["field_mapping"].items()],
    )
    workbook.save(target)
    workbook.close()


def export_daily_jira_xlsx(result: dict[str, Any], target: str | Path) -> None:
    """导出每日 Jira 最终人员汇总和人员归属明细。"""
    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "每日Jira汇总"
    summary_sheet.append(["汇总人员", "合并内容"])
    for item in result["summaries"]:
        summary_sheet.append([_excel_safe(item["name"]), _excel_safe(item["merged_text"])])

    detail_sheet = workbook.create_sheet("筛选明细")
    detail_sheet.append(["来源行", "问题关键字", "开发人员", "经办人", "报告人", "汇总人员", "汇总依据", "备注", "原始状态", "转换状态", "概要", "合并内容", "人员情况", "是否纳入"])
    for record in result["records"]:
        detail_sheet.append([_excel_safe(value) for value in [
            record["source_row"], record["issue_key"], record["developer"], record["assignee"],
            record["reporter"], record["summary_person"], record["summary_source"], record["summary_note"],
            record["status_raw"], record["status"], record["summary"], record["summary_text"],
            record["consistency"], "是" if record["included"] else "否",
        ]])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for sheet in (summary_sheet, detail_sheet):
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
    summary_sheet.column_dimensions["A"].width = 18
    summary_sheet.column_dimensions["B"].width = 80
    for row_index in range(2, summary_sheet.max_row + 1):
        line_count = max(1, str(summary_sheet.cell(row_index, 2).value or "").count("\n") + 1)
        summary_sheet.row_dimensions[row_index].height = min(18 * line_count, 180)
    detail_widths = [10, 18, 18, 18, 18, 18, 14, 40, 14, 14, 42, 52, 24, 12]
    for index, width in enumerate(detail_widths, start=1):
        detail_sheet.column_dimensions[get_column_letter(index)].width = width
    workbook.save(target)
    workbook.close()
