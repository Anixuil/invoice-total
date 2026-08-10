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
from openpyxl.styles import Alignment, Font, PatternFill
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
            modules = [_text(values[index]) for index in module_indexes if index < len(values) and _text(values[index])]
            consistency = "待补开发人员" if not developer else "一致" if assignee == developer else "人员不一致"
            display_status = STATUS_MAP.get(status_raw, status_raw or "未填写")
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
                "reporter": _text(get("reporter")),
                "developer": developer,
                "modules": modules,
                "module": "、".join(dict.fromkeys(modules)),
                "delivery_date": _date_text(get("delivery_date")),
                "resolution": _text(get("resolution")),
                "issue_type": _text(get("issue_type")),
                "planned_completion": _date_text(get("planned_completion")),
                "consistency": consistency,
                "included": consistency == "一致",
                "summary_text": f"{summary}【{display_status}】",
            }
            records.append(record)

        _notify(progress_callback, "清洗数据", 42, f"已读取 {len(records)} 条任务")
        key_counts = Counter(record["issue_key"] for record in records)
        roster_order = {name: index + 1 for index, name in enumerate(DEFAULT_ROSTER)}
        for record in records:
            record["sort_order"] = roster_order.get(record["assignee"], 9999)
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
            grouped[record["assignee"]].append(record)

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
                anomalies.append({"row": record["source_row"], "issue_key": record["issue_key"], "type": "developer_missing", "label": "待补开发人员", "detail": "开发人员为空，未纳入按人员一致性汇总"})
            elif record["consistency"] == "人员不一致":
                anomalies.append({"row": record["source_row"], "issue_key": record["issue_key"], "type": "developer_mismatch", "label": "人员不一致", "detail": f"经办人：{record['assignee']}；开发人员：{record['developer']}"})
            if record["status_raw"] not in KNOWN_STATUSES:
                anomalies.append({"row": record["source_row"], "issue_key": record["issue_key"], "type": "unknown_status", "label": "未知状态", "detail": record["status_raw"] or "状态为空"})
            if not record["planned_completion"]:
                anomalies.append({"row": record["source_row"], "issue_key": record["issue_key"], "type": "planned_date_missing", "label": "计划完成日期缺失", "detail": "原始字段为空"})
            if record["duplicate"]:
                anomalies.append({"row": record["source_row"], "issue_key": record["issue_key"], "type": "duplicate_issue", "label": "问题关键字重复", "detail": "建议回到 Jira 导出数据确认是否重复"})

        status_counts = Counter(record["status_raw"] for record in records)
        _notify(progress_callback, "核验并分组", 78, f"一致性汇总 {len(included)} 条，异常 {len(anomalies)} 条")
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
                "unique_assignees": len({record["assignee"] for record in records}),
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
            sheet.column_dimensions["F"].width = 70
            for row_index, row in enumerate(sheet.iter_rows(min_row=2), start=2):
                row[5].alignment = Alignment(vertical="top", wrap_text=True)
                line_count = max(1, str(row[5].value or "").count("\n") + 1)
                sheet.row_dimensions[row_index].height = min(18 * line_count, 120)

    add_sheet(
        "人员汇总",
        ["人员", "任务数", "已完成", "进行中", "开放", "合并任务内容"],
        [[item["name"], item["task_count"], item["completed_count"], item["in_progress_count"], item["open_count"], item["merged_text"]] for item in result["summaries"]],
    )
    add_sheet(
        "任务明细",
        ["来源行", "问题关键字", "问题ID", "父级ID", "原始状态", "展示状态", "创建日期", "概要", "经办人", "报告人", "开发人员", "模块", "交付日期", "解决结果", "问题类型", "计划完成日期", "一致性", "是否纳入汇总"],
        [[record[field] if field != "是否纳入汇总" else ("是" if record["included"] else "否") for field in ["source_row", "issue_key", "issue_id", "parent_id", "status_raw", "status", "created_at", "summary", "assignee", "reporter", "developer", "module", "delivery_date", "resolution", "issue_type", "planned_completion", "consistency", "是否纳入汇总"]] for record in result["records"]],
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
