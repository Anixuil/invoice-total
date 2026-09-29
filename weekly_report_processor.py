#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""部门项目周报的 PPT 审核、模板化合并和周例会文档生成。"""

from __future__ import annotations

from copy import deepcopy
from datetime import date, datetime, timedelta
from difflib import SequenceMatcher
from functools import lru_cache
from io import BytesIO
from pathlib import Path
import re
import unicodedata
from typing import Any, Callable

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Centipoints, Pt


ProgressCallback = Callable[[dict[str, Any]], None] | None
BASE = Path(__file__).resolve().parent
PPT_TEMPLATE = BASE / "templates" / "weekly_report_template.pptx"
DOCX_TEMPLATE = BASE / "templates" / "weekly_meeting_template.docx"

PPT_HEADING = re.compile(r"^(.+?)[（(]\s*汇报人\s*[：:]\s*(.+?)[）)]\s*$")
DATE_PATTERN = re.compile(r"(20\d{2})[年/\.\-](\d{1,2})[月/\.\-](\d{1,2})日?")
COVER_WEBSITE = "www.kingsware.cn"
PROJECT_HEADING = re.compile(
    r"^[一二三四五六七八九十百零〇0-9]+[、.．]\s*(.+?)[（(]\s*汇报人\s*[：:]\s*(.+?)[）)]\s*$"
)

SECTION_LABELS = {
    "current": ("本周工作完成情况", "本周完成情况", "本周工作进展", "本周工作进度", "本周进度", "本周进展", "本周完成", "本周工作概述", "本周关键进展", "本周工作总结"),
    "next": ("下周计划", "下周工作计划", "下周研发计划", "下周工作重点", "下周工作安排"),
    "issues": ("本周问题", "问题、风险", "问题与风险", "问题及风险", "风险与问题", "风险及问题", "风险与阻塞项", "本周风险", "风险问题", "问题", "风险", "阻塞项", "阻塞"),
}
SECTION_DISPLAY = {"current": "本周进展：", "next": "下周计划：", "issues": "本周问题："}
PPT_BODY_FONT = "思源黑体CN VF Light"
PPT_REPORTER_FONT = "思源黑体 CN Bold"
PPT_ISSUE_COLOR = RGBColor(255, 0, 0)
PPT_BODY_MARGIN_LEFT = Pt(8)
PPT_BODY_MARGIN_TOP = Pt(6)
CONTINUATION_MARKER = "\u200b"
PPT_FLOW_ORDER = ("current", "next", "issues")
PPT_FLOW_LABELS = {"current": "本周进展", "issues": "本周问题", "next": "下周计划"}
PPT_BODY_FONT_STEPS = (14, 13, 12, 11, 10.5, 10, 9.5, 9)
PPT_FLOW_FONT_STEPS = (10.5, 10, 9.5, 9)
PPT_FLOW_FONT_SIZE = PPT_FLOW_FONT_STEPS[0]
PPT_FLOW_LINE_HEIGHT = Pt(16.5)
PPT_FLOW_LABEL_FONT_SIZE = 14
PPT_FLOW_LABEL_LINE_HEIGHT = Pt(18)
PPT_FLOW_BOTTOM_SAFETY = Pt(4)
PPT_FLOW_GAP = Pt(6)
PPT_ISSUE_WIDTH_FRACTIONS = (0.35, 0.30, 0.25, 0.20, 0.15)
PPT_FLOW_LABEL_HEIGHT = PPT_FLOW_LABEL_LINE_HEIGHT * 4 + 2 * PPT_BODY_MARGIN_TOP
IGNORED_SLIDE_TITLES = (
    "在建项目整体情况",
    "项目整体情况",
    "本周jira情况",
    "jira情况",
)

TITLE_ALIASES = {
    "物业商机推荐Agent": "物业商机智能体",
    "物业商机推荐智能体": "物业商机智能体",
    "南京银行": "南京银行混沌测试",
    "南京银行项目周报": "南京银行混沌测试",
    "南京混沌测试": "南京银行混沌测试",
    "应急平台": "应急指挥调度平台",
    "应急演练指挥调度平台": "应急指挥调度平台",
    "中信证劵应急指挥调度平台二期": "中信证券应急指挥调度平台二期",
    "国信指标平台": "国信证券指标管理平台",
    "五矿证券": "五矿证券日志管理项目",
    "五矿日志项目": "五矿证券日志管理项目",
    "金智维K-Loghub日志平台": "五矿证券日志管理项目",
    "K-Loghub日志平台": "五矿证券日志管理项目",
    "墨巡miciusops智能运维平台": "墨巡MiciusOps智能运维平台",
    "市场工作成果与计划": "本周市场工作成果与计划",
}


def _notify(callback: ProgressCallback, stage: str, percent: int, detail: str) -> None:
    if callback:
        callback({"stage": stage, "percent": max(0, min(100, percent)), "detail": detail})


def _text(value: Any) -> str:
    return str(value or "").replace("\u3000", " ").strip()


def _normalize_title(value: str) -> str:
    text = PPT_HEADING.sub(r"\1", _text(value))
    text = re.sub(r"[\s:：,，、()（）\[\]【】\-—_/]+", "", text).lower()
    text = re.sub(r"(周报|项目报告|项目)$", "", text)
    return text


NORMALIZED_ALIASES = {
    _normalize_title(alias): _normalize_title(canonical)
    for alias, canonical in TITLE_ALIASES.items()
}


def _canonical_title(value: str) -> str:
    normalized = _normalize_title(value)
    return NORMALIZED_ALIASES.get(normalized, normalized)


def _title_score(left: str, right: str) -> tuple[float, str]:
    left_raw = _normalize_title(left)
    right_raw = _normalize_title(right)
    if not left_raw or not right_raw:
        return 0.0, "none"
    if left_raw == right_raw:
        return 1.0, "exact"
    left_key = _canonical_title(left)
    right_key = _canonical_title(right)
    if left_key == right_key:
        return 0.96, "alias"
    shorter, longer = sorted((left_key, right_key), key=len)
    if len(shorter) >= 4 and shorter in longer and len(shorter) / len(longer) >= 0.55:
        return 0.84, "similar"
    return SequenceMatcher(None, left_key, right_key).ratio(), "similar"


def _clean_line(value: str) -> str:
    # 兼容 PowerPoint/WPS 常见的字体映射项目符号（例如 U+F0B7 的“”），
    # 避免清理后又和统一的正文圆点叠加。
    return re.sub(r"^[\s·•▪◆◇■□▍]+", "", _text(value)).strip()


def _has_ordered_prefix(value: str) -> bool:
    """Return whether a paragraph already carries an explicit ordered marker."""
    return bool(re.match(
        r"^(?:(?:\d+|[A-Za-z])[\s]*[.\uFF0E、)]|"
        r"[\(（](?:\d+|[A-Za-z一-鿿]+)[\)）]|"
        r"[一-鿿]+[、.\uFF0E])\s*",
        _text(value),
    ))


def _clean_content(value: str) -> str:
    lines = []
    for raw in re.split(r"[\r\n\v]+", _text(value)):
        # 归一连续空格，不根据句尾汉字和百分比猜测姓名或插入换行。
        line = re.sub(r"[ \t]{2,}", " ", _clean_line(raw))
        if line and not re.fullmatch(r"[XxＸ]+(?:【.*?】)?", line):
            lines.append(line)
    return "\n".join(dict.fromkeys(lines))


def _clean_supplementary_content(value: str) -> str:
    """Keep paragraph breaks and repeated lines inside an actual source block."""
    return "\n".join(
        re.sub(r"[ \t]{2,}", " ", _clean_line(line))
        for line in re.split(r"\r\n|[\r\n\v]", _text(value))
    ).strip()


def _normalized_section_content(value: Any) -> str:
    """Normalize only whole empty placeholders, retaining actual paragraphs."""
    content = _clean_supplementary_content(_text(value))
    return "无" if not content or re.fullmatch(r"(?:无|暂无|未填写)[。．.!！]?", content) else content


def _is_empty_section_content(value: Any) -> bool:
    return _normalized_section_content(value) == "无"


def _merge_section_blocks(values, separator: str = "\n\n") -> str:
    """Remove duplicate whole blocks and discard empty blocks if content exists."""
    blocks = []
    for value in values:
        content = _normalized_section_content(value)
        if content != "无" and content not in blocks:
            blocks.append(content)
    return separator.join(blocks) if blocks else "无"


def _known_section_title_key(title: str) -> str | None:
    compact = re.sub(r"\s+", "", _text(title)).rstrip("：:")
    for key, labels in SECTION_LABELS.items():
        if compact in labels:
            return key
    return None


def _supplementary_key(title: str, kind: str = "normal") -> str:
    title = re.sub(r"\s+", "", _text(title)).rstrip("：:")
    known_key = _known_section_title_key(title)
    if kind == "issue" or known_key == "issues" or re.search(r"问题|风险|阻塞", title):
        return "issues"
    return known_key or f"supplementary:normal:{title}"


def _canonical_section_key(key: str) -> str:
    if key.startswith("supplementary:"):
        parts = key.split(":", 2)
        if len(parts) == 3:
            return _supplementary_key(parts[2], parts[1])
    return _known_section_title_key(key) or key


def _section_title(key: str) -> str:
    key = _canonical_section_key(key)
    return key.split(":", 2)[2] if key.startswith("supplementary:") else PPT_FLOW_LABELS[key]


def _section_is_issue(key: str | None) -> bool:
    return bool(key and _canonical_section_key(key) == "issues")


def _is_supplementary_key(key: str | None) -> bool:
    return bool(key and _canonical_section_key(key) not in {"current", "next"})


def _supplementary_sections(value: dict[str, Any]) -> list[dict[str, str]]:
    sections = value.get("supplementary_sections")
    if not isinstance(sections, list):
        sections = [{"title": "本周问题", "kind": "issue", "content": value.get("issues")}]
    grouped = {}
    for section in sections:
        if not isinstance(section, dict):
            continue
        title = _text(section.get("title")) or "补充内容"
        key = _supplementary_key(title, section.get("kind", "normal"))
        bucket = grouped.setdefault(key, {
            "title": _section_title(key), "kind": "issue" if _section_is_issue(key) else "normal", "blocks": [],
        })
        bucket["blocks"].append(section.get("content"))
    return [{"title": item["title"], "kind": item["kind"], "content": _merge_section_blocks(item["blocks"])}
            for item in grouped.values()]


def _output_sections(value: dict[str, Any]) -> dict[str, dict[str, str]]:
    sections = {
        key: {"title": PPT_FLOW_LABELS[key], "kind": "normal", "content": _normalized_section_content(value.get(key))}
        for key in ("current", "next")
    }
    for section in _supplementary_sections(value):
        key = _supplementary_key(section["title"], section["kind"])
        sections[key] = {**section, "content": _merge_section_blocks(
            [sections[key]["content"], section["content"]] if key in sections else [section["content"]]
        )}
    return sections


def _date_value(value: str) -> str:
    match = DATE_PATTERN.search(_text(value))
    if not match:
        return _text(value)
    return f"{int(match.group(1)):04d}-{int(match.group(2)):02d}-{int(match.group(3)):02d}"


def current_week_saturday() -> date:
    """按周一至周日计算本周周六，周日仍归入刚结束的这一周。"""
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    return monday + timedelta(days=5)


def _unique_cells(row) -> list[Any]:
    cells, seen = [], set()
    for cell in row.cells:
        marker = id(cell._tc)
        if marker not in seen:
            cells.append(cell)
            seen.add(marker)
    return cells


def _cell_lines(cell) -> list[str]:
    lines = []
    for paragraph in cell.paragraphs:
        lines.extend(part for part in re.split(r"[\r\n\v]+", paragraph.text) if part.strip())
    return lines


def _meeting_cell(document: Document):
    for table in document.tables:
        for row in table.rows:
            cells = _unique_cells(row)
            if cells and re.sub(r"\s+", "", cells[0].text) == "会议主要内容":
                return cells[-1]
    raise ValueError("部门周例会模板中没有找到“会议主要内容”区域")


def _meeting_date_cell(document: Document):
    for table in document.tables:
        for row in table.rows:
            cells = _unique_cells(row)
            if len(cells) >= 2 and re.sub(r"\s+", "", cells[0].text) == "会议时间":
                return cells[1]
    raise ValueError("部门周例会模板中没有找到“会议时间”字段")


@lru_cache(maxsize=1)
def _template_dimensions() -> tuple[int, int]:
    """缓存模板页面尺寸，避免每个源文件都重复打开模板。"""
    presentation = Presentation(PPT_TEMPLATE)
    return int(presentation.slide_width), int(presentation.slide_height)


def _project_slide_numbers(presentation: Presentation) -> list[int]:
    return [
        slide_number
        for slide_number, slide in enumerate(presentation.slides, start=1)
        if any(
            entry["text"] and PPT_HEADING.match(entry["text"])
            for entry in _flatten_shapes(slide.shapes)
        )
    ]


def _template_projects() -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """读取内置 PPT 和 Word 模板中的项目顺序，不把项目名单硬编码在处理器内。"""
    ppt = Presentation(PPT_TEMPLATE)
    deck_projects = []
    seen = set()
    for slide_number, slide in enumerate(ppt.slides, start=1):
        title = next((
            entry["text"] for entry in _flatten_shapes(slide.shapes)
            if entry["text"] and PPT_HEADING.match(entry["text"])
        ), "")
        match = PPT_HEADING.match(title)
        if not match:
            continue
        key = _canonical_title(match.group(1))
        if key in seen:
            continue
        seen.add(key)
        deck_projects.append({"id": f"deck_{len(deck_projects) + 1}", "key": key, "title": match.group(1), "reporter": match.group(2), "template_slide": slide_number})

    document = Document(DOCX_TEMPLATE)
    meeting_projects = []
    for paragraph in _meeting_cell(document).paragraphs:
        match = PROJECT_HEADING.match(_text(paragraph.text))
        if not match:
            continue
        key = _canonical_title(match.group(1))
        meeting_projects.append({"id": f"meeting_{len(meeting_projects) + 1}", "key": key, "title": match.group(1), "reporter": match.group(2)})
    return deck_projects, meeting_projects


def _flatten_shapes(shapes, parent_left: int = 0, parent_top: int = 0, prefix: str = "") -> list[dict[str, Any]]:
    entries = []

    def collect(items, offset_x, offset_y, scale_x, scale_y, path_prefix):
        for index, shape in enumerate(items):
            path = f"{path_prefix}{index + 1}"
            left = offset_x + int(shape.left) * scale_x
            top = offset_y + int(shape.top) * scale_y
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                transform = shape._element.grpSpPr.xfrm
                if (transform is not None and transform.chOff is not None
                        and transform.chExt is not None and transform.chExt.cx and transform.chExt.cy):
                    child_scale_x = scale_x * shape.width / transform.chExt.cx
                    child_scale_y = scale_y * shape.height / transform.chExt.cy
                    collect(shape.shapes, left - transform.chOff.x * child_scale_x,
                            top - transform.chOff.y * child_scale_y,
                            child_scale_x, child_scale_y, f"{path}.")
                else:
                    collect(shape.shapes, left, top, scale_x, scale_y, f"{path}.")
                continue
            entries.append({
                "path": path, "shape": shape, "text": _text(getattr(shape, "text", "")),
                "left": round(left), "top": round(top),
                "width": round(int(shape.width) * scale_x), "height": round(int(shape.height) * scale_y),
            })

    collect(shapes, parent_left, parent_top, 1.0, 1.0, prefix)
    _mark_section_entries(entries)
    return entries


def _section_label(text: str, key: str | None = None) -> str:
    canonical_key = _canonical_section_key(key) if key else None
    if canonical_key in SECTION_LABELS:
        labels = SECTION_LABELS[canonical_key]
        if key and key.startswith("supplementary:"):
            labels = (*labels, key.split(":", 2)[2])
    elif key and key.startswith("supplementary:"):
        labels = (key.split(":", 2)[2],)
    else:
        labels = (
            *tuple(label for values in SECTION_LABELS.values() for label in values), "项目概览",
        )
    value = _text(text)
    compact = re.sub(r"\s+", "", value)
    first_line = re.sub(r"\s+", "", re.split(r"[\r\n\v]+", value, maxsplit=1)[0])
    for label in sorted(labels, key=len, reverse=True):
        if compact == label or compact.startswith((f"{label}：", f"{label}:")):
            return label
        if first_line in {label, f"{label}：", f"{label}:"}:
            return label
    return ""


def _is_section_label(text: str, key: str | None = None) -> bool:
    return bool(_section_label(text, key))


def _mark_section_entries(entries: list[dict[str, Any]]) -> None:
    """Recognize renamed right-column furniture, never headings inside a body box."""
    for entry in entries:
        text = entry["text"]
        entry["section_key"] = None
        entry["section_label"] = ""
        name = entry["shape"].name
        if name.startswith("weekly-section:") and text:
            source_key = name[len("weekly-section:"):]
            entry["section_key"] = _canonical_section_key(source_key)
            entry["section_label"] = _section_label(text, source_key) or re.sub(r"\s+", "", text).rstrip("：:")
            continue
        if name.startswith("weekly-flow:") and name.endswith(":label"):
            key = name[len("weekly-flow:"):].rsplit(":", 2)[0]
            if key in SECTION_LABELS or key.startswith("supplementary:"):
                entry["section_key"] = _canonical_section_key(key)
                entry["section_label"] = _section_label(text, key) or re.sub(r"\s+", "", text).rstrip("：:")
                continue
        for key in ("current", "next", "issues"):
            label = _section_label(text, key)
            if label:
                entry["section_key"] = key
                entry["section_label"] = label
                break
        if entry["section_key"]:
            continue
        if _section_label(text, _supplementary_key("项目概览")):
            entry["section_key"] = _supplementary_key("项目概览")
            entry["section_label"] = "项目概览"
            continue
        title = re.sub(r"\s+", "", text).rstrip("：:")
        if (entry["top"] < 900000 or entry["left"] < 6000000
                or not 1 <= len(title) <= 24 or re.search(r"[：:。；;，,！？!?]", title)
                or entry["width"] > 2000000 or PPT_HEADING.match(text)):
            continue
        shape = entry["shape"]
        if not getattr(shape, "has_text_frame", False):
            continue
        vertical = shape.text_frame._txBody.bodyPr.get("vert", "horz") in {"eaVert", "vert", "vert270"}
        narrow = entry["width"] <= 1000000 and entry["height"] >= entry["width"] * 1.5
        colored = shape.fill.type not in {None, MSO_FILL_TYPE.BACKGROUND}
        backed = any(
            not other["text"] and other["width"] <= entry["width"] + 200000
            and other["left"] <= entry["left"] + 25000 and other["top"] <= entry["top"] + 25000
            and other["left"] + other["width"] >= entry["left"] + entry["width"] - 25000
            and other["top"] + other["height"] >= entry["top"] + entry["height"] - 25000
            and getattr(other["shape"], "fill", None) is not None
            and other["shape"].fill.type not in {None, MSO_FILL_TYPE.BACKGROUND}
            for other in entries
        )
        adjacent = any(
            other is not entry and other["left"] >= 6000000
            and other["width"] > entry["width"] * 1.5
            and getattr(other["shape"], "has_text_frame", False)
            and other["top"] < entry["top"] + entry["height"] + 100000
            and entry["top"] < other["top"] + other["height"] + 100000
            for other in entries
        )
        if adjacent and (vertical or narrow or colored or backed):
            entry["section_key"] = _supplementary_key(title)
            entry["section_label"] = title


def _entry_section_key(entry: dict[str, Any]) -> str | None:
    if "section_key" in entry:
        return _canonical_section_key(entry["section_key"]) if entry["section_key"] else None
    # Small synthetic/legacy callers may supply entries without shape metadata.
    for key in SECTION_LABELS:
        label = _section_label(entry["text"], key)
        if label:
            return _supplementary_key(label, "issue") if key == "issues" else key
    return _supplementary_key("项目概览") if _is_section_label(entry["text"]) else None


def _entry_role(entry: dict[str, Any]) -> str:
    return _entry_section_key(entry) or _text_role(entry["text"])


def _content_after_section_label(text: str, key: str, source_label: str = "") -> str:
    label = source_label or _section_label(text, key)
    if not label:
        return ""
    spaced_label = r"\s*".join(map(re.escape, label))
    remainder = re.sub(rf"^\s*{spaced_label}\s*[：:]?\s*", "", _text(text), count=1)
    return _clean_supplementary_content(remainder)


def _section_from_entries(entries: list[dict[str, Any]], key: str) -> str:
    key = _canonical_section_key(key)
    label_entries = [item for item in entries if _entry_section_key(item) == key]
    if not label_entries:
        return ""
    values = []
    for label in label_entries:
        inline_content = _content_after_section_label(label["text"], key, label.get("section_label", ""))
        if inline_content:
            values.append((label["top"], label["left"], inline_content))
    for item in entries:
        if _entry_section_key(item) or PPT_HEADING.match(item["text"]):
            continue
        if _body_section_key(entries, item) != key:
            continue
        content = _clean_supplementary_content(item["text"])
        if content:
            values.append((item["top"], item["left"], content))
    return _merge_section_blocks((value for _, _, value in sorted(values)), separator="\n")


def _section_lines(value: str) -> list[str]:
    """Return normalized, non-empty content lines while retaining display text."""
    return list(dict.fromkeys(
        line for raw_line in _text(value).splitlines()
        if (line := _clean_line(raw_line))
    ))


def _prefer_later_section_content(section_values: dict[str, str]) -> dict[str, str]:
    """Keep a duplicated item in the later, explicitly labeled weekly section.

    A frequent source-slide editing mistake is leaving a next-week bullet at the
    bottom of the current-week text box while also listing it under "下周计划".
    The explicit later section is the authoritative location for that item.
    """
    result = dict(section_values)
    next_lines = {
        _normalized_content(line)
        for line in _section_lines(result.get("next", ""))
    }

    current_lines = _clean_supplementary_content(result.get("current", "")).splitlines()
    result["current"] = _normalized_section_content("\n".join(
        line for line in current_lines
        if not line.strip() or _normalized_content(line) not in next_lines
    ))
    return result


def _slide_title(entries: list[dict[str, Any]]) -> tuple[str, str, str]:
    for item in entries:
        match = PPT_HEADING.match(item["text"])
        if match:
            return match.group(1).strip(), match.group(2).strip(), "explicit"
    candidates = [item for item in entries if item["top"] < 1200000 and item["text"] and not _entry_section_key(item)]
    if not candidates:
        return "", "", "missing"
    return min(candidates, key=lambda item: (item["top"], item["left"]))["text"], "", "inferred"


def _is_ignored_slide(texts: list[str], slide_number: int, slide_count: int) -> bool:
    """识别不参与项目核算的封面、结束页和固定汇总页。"""
    if slide_number in {1, slide_count}:
        return True
    compact_texts = [re.sub(r"\s+", "", _text(text)).lower() for text in texts]
    if any(_is_outro_text(text) for text in compact_texts):
        return True
    return any(
        title in text
        for text in compact_texts
        for title in IGNORED_SLIDE_TITLES
    )


def _is_outro_text(text: str) -> bool:
    compact = re.sub(r"[^\u4e00-\u9fffA-Za-z0-9]", "", _text(text)).lower()
    return "期待与您携手共赢" in compact


def _is_outro_slide(slide) -> bool:
    texts = [entry["text"] for entry in _flatten_shapes(slide.shapes) if entry["text"]]
    # 文案可能被拆成多个文本框（例如“期待与您携手”+“共赢”）。
    joined = re.sub(r"[^\u4e00-\u9fffA-Za-z0-9]", "", "".join(texts)).lower()
    return "期待与您携手共赢" in joined or any(_is_outro_text(text) for text in texts)


def _audit_shape_bounds(entries: list[dict[str, Any]], width: int, height: int, file: str, slide: int, project: str) -> list[dict[str, Any]]:
    issues = []
    for entry in entries:
        if entry["left"] < 0 or entry["top"] < 0 or entry["left"] + entry["width"] > width or entry["top"] + entry["height"] > height:
            issues.append({
                "severity": "warning", "code": "shape_overflow", "label": "对象超出画布",
                "file": file, "slide": slide, "location": f"对象 {entry['path']}", "project": project,
                "detail": "对象边界超出 PPT 页面，可能导致导出后裁切。",
                "suggestion": "回到原 PPT 调整对象大小或位置。",
            })
    return issues


def parse_presentation_source(path: str | Path, display_name: str, expected: list[dict[str, Any]]) -> dict[str, Any]:
    presentation = Presentation(path)
    width, height = int(presentation.slide_width), int(presentation.slide_height)
    template_width, template_height = _template_dimensions()
    issues = []
    if (width, height) != (template_width, template_height):
        issues.append({"severity": "error", "code": "slide_size", "label": "页面尺寸不一致", "file": display_name, "slide": 0, "location": "整份 PPT", "project": "", "detail": f"当前尺寸 {width}x{height}，模板尺寸 {template_width}x{template_height}。", "suggestion": "使用部门项目周报模板的宽高。"})
    slides = []
    slide_count = len(presentation.slides)
    for slide_number, slide in enumerate(presentation.slides, start=1):
        entries = _flatten_shapes(slide.shapes)
        texts = [entry["text"] for entry in entries if entry["text"]]
        has_sections = any(_entry_section_key(entry) for entry in entries)
        # 结束页可能包含模板残留文本，不能仅依赖“无内容字段”判断，
        # 否则它会被当作项目页复制到汇总 PPT，造成重复结尾页。
        is_outro = _is_outro_slide(slide)
        if is_outro or (not has_sections and _is_ignored_slide(texts, slide_number, slide_count)):
            continue
        if not texts:
            issues.append({"severity": "warning", "code": "blank_slide", "label": "空白页", "file": display_name, "slide": slide_number, "location": "整页", "project": "", "detail": "页面没有可识别文字内容。", "suggestion": "确认是否是误上传的空白页。"})
            continue
        if not has_sections:
            issues.append({"severity": "info", "code": "non_project_slide", "label": "非项目页", "file": display_name, "slide": slide_number, "location": "整页", "project": "", "detail": "封面、结束页或说明页未纳入项目合并。", "suggestion": "如需纳入，请在模板中增加对应项目页。"})
            continue
        title, reporter, title_mode = _slide_title(entries)
        if not title:
            issues.append({"severity": "error", "code": "title_missing", "label": "项目标题缺失", "file": display_name, "slide": slide_number, "location": "标题区域", "project": "", "detail": "项目页存在内容字段，但没有可识别的项目标题。", "suggestion": "补充项目名称并放在页面顶部。"})
            continue
        scores = sorted((_title_score(title, item["title"]) + (item,) for item in expected), key=lambda value: value[0], reverse=True)
        best_score, best_mode, best_project = scores[0] if scores else (0, "none", {"key": "", "title": ""})
        if best_score < 0.58:
            best_project = {"key": "", "title": ""}
        project_title = best_project.get("title", "")
        if (best_mode in {"alias", "similar"} or title_mode == "inferred") and project_title:
            issues.append({"severity": "warning", "code": "title_alias", "label": "标题需确认", "file": display_name, "slide": slide_number, "location": f"对象 {next((e['path'] for e in entries if e['text'] == title), '标题区域')}", "project": project_title, "detail": f"源标题“{title}”按别名或相似规则对应“{project_title}”。", "suggestion": "确认该页确实属于对应项目。"})
        embedded_packages = _embedded_package_relationships(slide)
        if embedded_packages:
            issues.append({
                "severity": "warning", "code": "embedded_object_dropped", "label": "嵌入对象已忽略",
                "file": display_name, "slide": slide_number, "location": "嵌入对象", "project": project_title,
                "detail": f"该页包含 {len(embedded_packages)} 个嵌入式对象（{', '.join(embedded_packages)}），整合时将跳过对象本身。",
                "suggestion": "如需保留显示效果，请在源 PPT 中将对象转为图片或普通 PPT 内容后重新上传。",
            })
        placeholders = [entry for entry in entries if re.search(r"X{2,}|Ｘ{2,}", entry["text"])]
        for entry in placeholders:
            issues.append({"severity": "error", "code": "placeholder", "label": "模板占位符未替换", "file": display_name, "slide": slide_number, "location": f"对象 {entry['path']}", "project": project_title, "detail": f"发现未替换内容：{entry['text'][:80]}", "suggestion": "补充真实周报内容后再合并。"})
        supplementary_keys = list(dict.fromkeys(
            key for entry in sorted(entries, key=lambda item: (item["top"], item["left"]))
            if _is_supplementary_key(key := _entry_section_key(entry))
        ))
        supplementary = [
            {"title": _section_title(key), "kind": "issue" if _section_is_issue(key) else "normal",
             "content": _section_from_entries(entries, key)}
            for key in supplementary_keys
        ]
        unresolved = [entry for entry in entries
                      if entry["text"] and entry["left"] > 7200000 and entry["top"] >= 900000
                      and not _entry_section_key(entry) and not PPT_HEADING.match(entry["text"])
                      and _body_section_key(entries, entry) is None]
        supplementary_labels = [entry for entry in entries
                                if _is_supplementary_key(_entry_section_key(entry))]
        ambiguous = any(
            left["left"] >= 6000000 and right["left"] >= 6000000
            and abs(left["top"] - right["top"]) <= 100000
            and _entry_section_key(left) != _entry_section_key(right)
            for index, left in enumerate(supplementary_labels) for right in supplementary_labels[index + 1:]
        )
        if unresolved:
            supplementary.append({"title": "待确认内容", "kind": "normal", "content": "\n\n".join(
                _clean_supplementary_content(entry["text"])
                for entry in sorted(unresolved, key=lambda item: (item["top"], item["left"]))
            )})
        if unresolved or ambiguous:
            issues.append({"severity": "warning", "code": "supplementary_ambiguous", "label": "补充栏目需确认",
                           "file": display_name, "slide": slide_number, "location": "右侧补充栏目",
                           "project": project_title, "detail": "右侧正文缺少明确栏目标签，或多个栏目标签的位置重叠。",
                           "suggestion": "确认栏目名称及正文归属；未归属的文字已保留为“待确认内容”。"})
        supplementary = _supplementary_sections({"supplementary_sections": supplementary})
        section_values = _prefer_later_section_content({
            "current": _normalized_section_content(_section_from_entries(entries, "current")),
            "next": _normalized_section_content(_section_from_entries(entries, "next")),
            "issues": _merge_section_blocks(item["content"] for item in supplementary if item["kind"] == "issue"),
        })
        issues.extend(_audit_shape_bounds(entries, width, height, display_name, slide_number, project_title))
        slides.append({
            "id": f"{display_name}#{slide_number}", "file": display_name, "slide": slide_number,
            "title": title, "reporter": reporter, "title_mode": title_mode,
            "project_key": best_project.get("key", "") if project_title else "",
            "project_title": project_title, "score": round(best_score, 3),
            "current": section_values["current"], "next": section_values["next"], "issues": section_values["issues"],
            "supplementary_sections": supplementary,
            "section_presence": {
                section_key: any(
                    _section_is_issue(_entry_section_key(entry)) if section_key == "issues"
                    else _entry_section_key(entry) == section_key for entry in entries
                )
                for section_key in SECTION_DISPLAY
            },
        })
    return {"file": display_name, "slide_count": len(presentation.slides), "slides": slides, "issues": issues, "size": {"width": width, "height": height}}


def parse_document_source(path: str | Path, display_name: str, expected: list[dict[str, Any]]) -> dict[str, Any]:
    """按正文顺序读取单项目 Word 周报，表格逐行保留字段关系。"""
    document = Document(path)
    sections = {key: [] for key in SECTION_DISPLAY}
    presence = {key: False for key in SECTION_DISPLAY}
    labels = {
        key: names for key, names in SECTION_LABELS.items()
    }
    active = None
    active_supplement = None
    supplementary_blocks = []
    headers = []
    preamble = []

    def heading_text(value: str) -> str:
        return re.sub(r"^(?:第?[一二三四五六七八九十0-9]+[、.．、)）\s]+)", "", _text(value))

    def known_heading(value: str):
        heading = heading_text(value)
        for key, names in labels.items():
            for name in sorted(names, key=len, reverse=True):
                match = re.match(rf"^({re.escape(name)}(?:\s*[（(][^）)]*[）)])?)\s*(?:[：:]\s*(.*))?$", heading)
                if match:
                    return key, match.group(1).strip(), match.group(2) or ""
        match = re.match(r"^(项目概览)\s*(?:[：:]\s*(.*))?$", heading)
        return ("normal", match.group(1), match.group(2) or "") if match else None

    def heading_style(paragraph):
        style = paragraph.style
        candidates = [paragraph._p.pPr]
        current_style = style
        while current_style is not None:
            candidates.append(current_style.element.pPr)
            current_style = current_style.base_style
        level = next((
            node.find(qn("w:outlineLvl")).get(qn("w:val"))
            for node in candidates if node is not None and node.find(qn("w:outlineLvl")) is not None
        ), None)
        if level == "9":
            level = None
        style_id = style.style_id if style is not None else ""
        if style_id.lower() in {"normal", "bodytext", "listparagraph", "正文", "普通"}:
            style_id = ""
        if style is not None and style.element.get(qn("w:default")) == "1":
            style_id = ""
        return level, style_id

    heading_levels = set()
    heading_styles = set()
    for element in document.element.body.iter(qn("w:p")):
        paragraph = Paragraph(element, document)
        recognized = known_heading(paragraph.text)
        if recognized and not recognized[2]:
            level, style_id = heading_style(paragraph)
            if level is not None:
                heading_levels.add(level)
            if style_id:
                heading_styles.add(style_id)

    def styled_heading(paragraph) -> bool:
        level, style_id = heading_style(paragraph)
        return (level is not None and level in heading_levels) or bool(style_id and style_id in heading_styles)

    def custom_title(value: str) -> bool:
        title = heading_text(value).rstrip("：:").strip()
        return bool(
            1 < len(title) <= 40 and re.search(r"[\u4e00-\u9fffA-Za-z]", title)
            and not re.search(r"[\n：:。！？；;，,]", title)
            and not DATE_PATTERN.search(title)
            and not re.match(r"^(?:汇报人|报告人|作者|项目名称|报告日期|汇报日期)$", title)
        )

    def consume(value: str, *, is_heading: bool = False, field_label: bool = False) -> None:
        nonlocal active, active_supplement, headers
        value = _text(value)
        recognized = known_heading(value)
        if recognized is None and active is not None and (is_heading or field_label) and custom_title(value):
            recognized = "normal", heading_text(value).rstrip("：:").strip(), ""
        if recognized:
            key, title, inline_content = recognized
            if key in {"issues", "normal"}:
                normalized_key = _supplementary_key(title, "issue" if key == "issues" else "normal")
                key = normalized_key if normalized_key in SECTION_LABELS else "normal"
            active = key
            headers = []
            if key in presence:
                presence[key] = True
            if key in {"issues", "normal"}:
                active_supplement = {"title": title, "kind": "issue" if key == "issues" else "normal", "lines": []}
                supplementary_blocks.append(active_supplement)
                if inline_content:
                    active_supplement["lines"].append(inline_content)
            else:
                active_supplement = None
                sections[key].append([])
                if inline_content:
                    sections[key][-1].append(inline_content)
            return
        if active_supplement is not None:
            active_supplement["lines"].append(value)
        elif active:
            sections[active][-1].append(value)
        elif value:
            preamble.append(value)

    for block in document.element.body:
        if block.tag == qn("w:p"):
            paragraph = Paragraph(block, document)
            for index, line in enumerate(paragraph.text.splitlines() or [""]):
                consume(line, is_heading=index == 0 and styled_heading(paragraph))
        elif block.tag == qn("w:tbl"):
            table = Table(block, document)
            headers = []
            seen_cells = set()
            field_table = any(
                len(row.cells) == 2 and known_heading(row.cells[0].text)
                for row in table.rows
            )
            for row in table.rows:
                cells = []
                for cell in _unique_cells(row):
                    if cell._tc in seen_cells:
                        continue
                    seen_cells.add(cell._tc)
                    cells.append(cell)
                values = [_text(cell.text).replace("\n", "；") for cell in cells]
                if not any(values):
                    if active_supplement is not None:
                        active_supplement["lines"].append("")
                    continue
                left_known = known_heading(values[0]) if values else None
                # 风险表的标题行和数据行不得被当成任意两列字段；切换栏目时清空其表头。
                left_styled = bool(cells and cells[0].paragraphs and styled_heading(cells[0].paragraphs[0]))
                is_field = len(cells) == 2 and (left_known or (
                    custom_title(values[0]) and (field_table or left_styled or values[0].endswith(("：", ":")))
                    and (not headers or left_styled or values[0].endswith(("：", ":")))
                    and not any(value in {"等级", "影响", "应对措施"} for value in values)
                ))
                if is_field:
                    consume(cells[0].text, field_label=True)
                    for paragraph in cells[1].paragraphs:
                        for line in paragraph.text.splitlines() or [""]:
                            consume(line)
                    continue
                if left_known:
                    consume(values[0])
                    for value in values[1:]:
                        consume(value)
                    continue
                if active == "issues" and not headers and any(value in {"等级", "影响", "应对措施"} for value in values):
                    headers = values
                    continue
                if headers:
                    consume("；".join(f"{headers[index]}：{value}" if index < len(headers) else value
                                      for index, value in enumerate(values) if value))
                elif active_supplement is not None:
                    for cell in cells:
                        for paragraph in cell.paragraphs:
                            for line in paragraph.text.splitlines() or [""]:
                                consume(line)
                else:
                    # 两列表格也可用左侧字段名标注对应周报区域。
                    for value in values:
                        consume(value)

    supplementary = _supplementary_sections({"supplementary_sections": [
        {"title": block["title"], "kind": block["kind"], "content": "\n".join(block["lines"])}
        for block in supplementary_blocks
    ]})
    section_values = {key: _merge_section_blocks("\n".join(block) for block in blocks)
                      for key, blocks in sections.items()}
    section_values["issues"] = _merge_section_blocks(item["content"] for item in supplementary if item["kind"] == "issue")

    fallback_title = re.sub(r"(?:工作周报|项目周报|周报).*?$", "", Path(display_name.replace("\\", "/")).stem)
    title = next((value for value in preamble if re.search(r"[\u4e00-\u9fff]", value)
                  and not DATE_PATTERN.search(value) and value not in {"工作周报", "周报"}
                  and not re.match(r"^(汇报人|报告人|作者)[：:]", value)), fallback_title)
    title = re.sub(r"(?:工作周报|项目周报|周报).*$", "", title).strip() or fallback_title
    heading_match = PPT_HEADING.match(title)
    reporter_match = re.search(r"(?:汇报人|报告人|作者)\s*[：:]\s*([^\n（）()]+)", "\n".join(preamble))
    reporter = _text(reporter_match.group(1)) if reporter_match else ""
    if heading_match:
        title, reporter = heading_match.groups()
    scores = sorted((_title_score(title, item["title"]) + (item,) for item in expected), key=lambda item: item[0], reverse=True)
    score, mode, project = scores[0] if scores else (0, "none", {})
    if score < 0.58:
        project = {}
    issues = []
    if project and mode != "exact":
        issues.append({"severity": "warning", "code": "title_alias", "label": "标题需确认", "file": display_name,
                       "slide": 0, "location": "Word 标题", "project": project["title"],
                       "detail": f"源标题“{title}”对应“{project['title']}”。", "suggestion": "确认该文档所属项目。"})
    slides = []
    if any(presence.values()) or supplementary:
        slides.append({"id": f"{display_name}#1", "file": display_name, "slide": 1, "source_format": "docx",
                       "title": title, "reporter": reporter, "title_mode": "inferred",
                       "project_key": project.get("key", ""), "project_title": project.get("title", ""),
                       "score": round(score, 3), "section_presence": presence,
                       "supplementary_sections": supplementary, **section_values})
    return {"file": display_name, "slide_count": len(slides), "slides": slides, "issues": issues, "source_format": "docx"}


def _audit_project_pages(project: dict[str, Any], slides: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """按项目汇总审核续页，避免要求每一页都重复包含完整周报字段。"""
    if not slides:
        return []
    issues = []
    files = sorted({item["file"] for item in slides})
    file_label = "、".join(files)
    page_count = len(slides)
    reporters = list(dict.fromkeys(item["reporter"] for item in slides if item["reporter"]))
    expected_reporter = project.get("reporter", "")
    if not reporters:
        issues.append({
            "severity": "warning", "code": "reporter_missing", "label": "汇报人缺失",
            "file": file_label, "slide": 0, "location": "项目标题区域", "project": project["title"],
            "detail": f"项目共 {page_count} 页，所有页面的标题都没有包含“汇报人”信息。",
            "suggestion": "按模板补充汇报人。",
        })
    else:
        mismatched = [
            reporter for reporter in reporters
            if _normalize_title(reporter) != _normalize_title(expected_reporter)
        ]
        if mismatched:
            issues.append({
                "severity": "warning", "code": "reporter_mismatch", "label": "汇报人不一致",
                "file": file_label, "slide": 0, "location": "项目标题区域", "project": project["title"],
                "detail": f"项目共 {page_count} 页，源文件汇报人为“{'、'.join(mismatched)}”，模板为“{expected_reporter}”。",
                "suggestion": "以模板中的汇报人为准，并在审核结果确认。",
            })

    for section_key, section_label in SECTION_DISPLAY.items():
        if section_key == "issues" and any(item.get("supplementary_sections") for item in slides):
            continue
        section_exists = any(item["section_presence"].get(section_key) for item in slides)
        if not section_exists:
            issues.append({
                "severity": "warning", "code": "section_missing", "label": "内容字段缺失",
                "file": file_label, "slide": 0, "location": f"{section_label}项目汇总区域", "project": project["title"],
                "detail": f"项目共 {page_count} 页，所有页面均未识别到“{section_label.rstrip('：')}”字段。",
                "suggestion": "按项目周报模板补充该字段；没有内容时默认填写“无”。",
            })
    return issues


def _is_generated_artifact(display_name: str) -> bool:
    stem = Path(display_name.replace("\\", "/")).stem.strip()
    return bool(re.match(r"^(项目周报|部门周例会)\s*\d", stem))


def _is_meeting_only_slide(slide: dict[str, Any]) -> bool:
    """Reserve Zhang Keke and Guangxi team reports for the meeting document."""
    values = (
        _text(slide.get("reporter", "")),
        _text(slide.get("title", "")),
        _text(slide.get("project_title", "")),
        _text(slide.get("file", "")),
    )
    return any("张珂珂" in value or "广西团队" in value for value in values)


def _is_meeting_only_project(project: dict[str, Any]) -> bool:
    values = (
        _text(project.get("reporter", "")),
        _text(project.get("title", "")),
        _text(project.get("key", "")),
        *(_text(item.get("reporter", "")) for item in project.get("slides", [])),
        *(_text(item.get("title", "")) for item in project.get("slides", [])),
        *(_text(item.get("file", "")) for item in project.get("slides", [])),
    )
    return any("张珂珂" in value or "广西团队" in value for value in values)


def _is_guangxi_project(project: dict[str, Any]) -> bool:
    values = (
        _text(project.get("title", "")),
        _text(project.get("key", "")),
        *(_text(item.get("title", "")) for item in project.get("slides", [])),
        *(_text(item.get("file", "")) for item in project.get("slides", [])),
    )
    return any("广西团队" in value for value in values)


def _project_result(project: dict[str, Any], slides: list[dict[str, Any]], issues: list[dict[str, Any]], source_kind: str) -> dict[str, Any]:
    current = _merge_section_blocks(item.get("current") for item in slides)
    next_plan = _merge_section_blocks(item.get("next") for item in slides)
    supplementary = _supplementary_sections({"supplementary_sections": [
        section for slide in slides for section in _supplementary_sections(slide)
    ]})
    if not supplementary:
        supplementary = [{"title": "本周问题", "kind": "issue", "content": "无"}]
    problems = _merge_section_blocks(item["content"] for item in supplementary if item["kind"] == "issue")
    reporter = next((item["reporter"] for item in slides if item["reporter"]), project["reporter"])
    project_issues = [item for item in issues if item.get("project") == project["title"] or item.get("project") == project.get("key")]
    has_error = any(item["severity"] == "error" for item in project_issues)
    has_warning = any(item["severity"] == "warning" for item in project_issues)
    status = "需处理" if has_error else "待确认" if has_warning else "通过"
    if not slides:
        current = next_plan = problems = "无"
    return {
        "id": project["id"], "key": project["key"], "title": project["title"], "reporter": reporter,
        "template_slide": project.get("template_slide"),
        "status": status, "source_kind": source_kind, "slides": slides, "source_files": sorted({item["file"] for item in slides}),
        "current": current, "next": next_plan, "issues": problems, "checks": project_issues,
        "supplementary_sections": supplementary,
    }


def process_weekly_report(
    presentation_sources: list[tuple[Path, str]],
    source_manifest: list[dict[str, Any]],
    progress_callback: ProgressCallback = None,
) -> dict[str, Any]:
    """以 ZIP 内 PPTX / DOCX 为内容源，对照内置 PPT/Word 模板完成审核和合并计划。"""
    deck_projects, meeting_projects = _template_projects()
    expected = []
    for item in deck_projects + meeting_projects:
        if not any(existing["key"] == item["key"] for existing in expected):
            expected.append(item)
    _notify(progress_callback, "读取模板", 10, f"已载入 {len(deck_projects)} 个项目 PPT 页模板和 {len(meeting_projects)} 个会议项目")

    parsed_files = []
    all_issues = []
    for index, (path, display_name) in enumerate(presentation_sources, start=1):
        if _is_generated_artifact(display_name):
            artifact_issue = {"severity": "info", "code": "generated_artifact", "label": "历史成品已忽略", "file": display_name, "slide": 0, "location": "文件", "project": "", "detail": "该文件看起来是历史汇总周报，不作为本次项目源。", "suggestion": "保留在目录中用于追溯，但不参与合并。"}
            parsed_files.append({"file": display_name, "slide_count": 0, "slides": [], "issues": [artifact_issue]})
            all_issues.append(artifact_issue)
            continue
        parser = parse_document_source if Path(path).suffix.lower() == ".docx" else parse_presentation_source
        parsed = parser(path, display_name, expected)
        parsed_files.append(parsed)
        all_issues.extend(parsed["issues"])
        _notify(progress_callback, "审核项目周报", 12 + round(index / max(len(presentation_sources), 1) * 52), f"已审核 {index} / {len(presentation_sources)} 个文件")

    all_slides = [slide for parsed in parsed_files for slide in parsed["slides"]]
    meeting_only_keys = {
        slide["project_key"]
        for slide in all_slides
        if slide.get("project_key") and _is_meeting_only_slide(slide)
    }
    by_key: dict[str, list[dict[str, Any]]] = {item["key"]: [] for item in expected}
    added_projects: dict[str, list[dict[str, Any]]] = {}
    for slide in all_slides:
        if slide["project_key"]:
            by_key.setdefault(slide["project_key"], []).append(slide)
        else:
            title = _text(slide["title"])
            key = _canonical_title(title)
            if not key:
                all_issues.append({"severity": "warning", "code": "unmatched_slide", "label": "项目页未对应", "file": slide["file"], "slide": slide["slide"], "location": "标题区域", "project": "", "detail": "源项目页没有可识别的标题，无法自动添加。", "suggestion": "在页面顶部补充项目名称和汇报人。"})
                continue
            if _is_meeting_only_slide(slide):
                meeting_only_keys.add(key)
            added_projects.setdefault(key, []).append(slide)

    for project in expected:
        if project["key"] in meeting_only_keys:
            continue
        all_issues.extend(_audit_project_pages(project, by_key.get(project["key"], [])))

    for project in deck_projects:
        if project["key"] in meeting_only_keys:
            continue
        if by_key.get(project["key"]):
            continue
        all_issues.append({
            "severity": "warning", "code": "missing_project_source", "label": "项目源文件缺失",
            "file": "", "slide": 0, "location": project["title"], "project": project["title"],
            "detail": f"模板项目“{project['title']}”没有匹配到任何源 PPT / Word 内容，生成页仅保留模板结构。",
            "suggestion": "将该项目的源 PPT 或 Word 周报放入上传 ZIP 后重新处理。",
        })

    deck_results = [_project_result(item, by_key.get(item["key"], []), all_issues, "项目周报") for item in deck_projects]
    deck_results = [item for item in deck_results if item["key"] not in meeting_only_keys]
    deck_keys = {item["key"] for item in deck_projects}
    for item in meeting_projects:
        slides = by_key.get(item["key"], [])
        if item["key"] in deck_keys or not slides:
            continue
        project = {**item, "template_slide": None}
        all_issues.append({
            "severity": "info", "code": "meeting_project_added_to_deck", "label": "会议项目已加入总周报",
            "file": "、".join(sorted({slide["file"] for slide in slides})), "slide": 0,
            "location": "总周报", "project": item["title"],
            "detail": f"项目“{item['title']}”仅配置在周例会模板中，已按上传 PPT 原页追加到总周报。",
            "suggestion": "如需为空项目生成固定占位页，可再将该项目加入 PPT 模板。",
        })
        deck_results.append(_project_result(project, slides, all_issues, "周例会模板项目"))
    added_results = []
    for key, slides in added_projects.items():
        if key in meeting_only_keys:
            continue
        title = next((item["title"] for item in slides if item["title"]), key)
        reporter = next((item["reporter"] for item in slides if item["reporter"]), "")
        project = {
            "id": f"added_{len(deck_results) + 1}",
            "key": key,
            "title": title,
            "reporter": reporter,
            "template_slide": None,
        }
        all_issues.append({"severity": "info", "code": "project_auto_added", "label": "新增项目已加入", "file": "、".join(sorted({item["file"] for item in slides})), "slide": 0, "location": "总周报", "project": title, "detail": f"未配置在内置模板中的项目“{title}”已自动追加到总周报 PPT。", "suggestion": "如需将该项目同步到周例会 DOCX，请在周例会模板中增加对应项目条目。"})
        added_results.append(_project_result(project, slides, all_issues, "自动添加项目"))
    deck_results.extend(added_results)
    deck_results = [project for project in deck_results if not _is_meeting_only_project(project)]
    meeting_results = [_project_result(item, by_key.get(item["key"], []), all_issues, "周例会") for item in meeting_projects]
    meeting_results.extend({**item, "source_kind": "自动添加项目"} for item in added_results)
    meeting_result_keys = {item["key"] for item in meeting_results}
    for key in sorted(meeting_only_keys - meeting_result_keys):
        slides = by_key.get(key, [])
        if not slides:
            continue
        project = {
            "id": f"meeting_only_{len(meeting_results) + 1}",
            "key": key,
            "title": next((slide["title"] for slide in slides if slide.get("title")), key),
            "reporter": next((slide["reporter"] for slide in slides if slide.get("reporter")), ""),
            "template_slide": None,
        }
        meeting_results.append(_project_result(project, slides, all_issues, "周例会"))
    for project in meeting_results:
        if _is_guangxi_project(project):
            project["reporter"] = "郑乐园"
    meeting_result_keys = {item["key"] for item in meeting_results}
    for project in list(parsed_files):
        for slide in project.get("slides", []):
            if _is_meeting_only_slide(slide):
                meeting_only_keys.add(slide.get("project_key", ""))
    for project in deck_results + meeting_results:
        if len(project["source_files"]) > 1:
            all_issues.append({"severity": "warning", "code": "multiple_sources", "label": "多个源文件合并", "file": "、".join(project["source_files"]), "slide": 0, "location": project["title"], "project": project["title"], "detail": "同一项目来自多个周报文件，系统按源文件和页码顺序合并。", "suggestion": "确认这些文件是否是同一项目的不同内容页。"})

    for parsed in parsed_files:
        if parsed["file"] and not parsed["slides"] and not _is_generated_artifact(parsed["file"]):
            all_issues.append({"severity": "warning", "code": "no_project_slides", "label": "未识别项目页", "file": parsed["file"], "slide": 0, "location": "源文件", "project": "", "detail": "文件中没有识别到本周进度、下周计划或问题等周报字段。", "suggestion": "检查是否为封面、结束页或不符合模板的版式。"})

    week_end = current_week_saturday()
    error_count = sum(item["severity"] == "error" for item in all_issues)
    warning_count = sum(item["severity"] == "warning" for item in all_issues)
    project_count = len(deck_results)
    result = {
        "ok": True,
        "week_end": week_end.isoformat(),
        "output_stem": f"项目周报{week_end:%m%d}",
        "stats": {
            "project_count": project_count,
            "matched_projects": sum(bool(item["slides"]) for item in deck_results),
            "passed": sum(item["status"] == "通过" for item in deck_results),
            "pending": sum(item["status"] == "待确认" for item in deck_results),
            "issues": sum(item["status"] == "需处理" for item in deck_results),
            "error_count": error_count,
            "warning_count": warning_count,
            "source_files": len(presentation_sources),
            "source_slides": len(all_slides),
        },
        "projects": deck_results,
        "meeting_projects": meeting_results,
        "issues": all_issues,
        "files": parsed_files,
        "sources": source_manifest,
        "assembly": [slide for project in deck_results for slide in project["slides"]],
    }
    _notify(progress_callback, "完成审核", 88, f"项目页审核完成，发现 {error_count} 个错误、{warning_count} 个待确认项")
    _notify(progress_callback, "生成结果", 100, "可下载项目周报 PPT、周例会 DOCX 和审核报告")
    return result


def _remove_element(element) -> None:
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def _remove_slide(presentation: Presentation, slide) -> None:
    slide_index = next(index for index, candidate in enumerate(presentation.slides) if candidate.part is slide.part)
    slide_id = presentation.slides._sldIdLst[slide_index]
    presentation.part.drop_rel(slide_id.rId)
    presentation.slides._sldIdLst.remove(slide_id)


def _retain_single_outro_slide(presentation: Presentation) -> None:
    """Keep only the template's ending slide in the generated deck."""
    outro_slides = [slide for slide in presentation.slides if _is_outro_slide(slide)]
    if len(outro_slides) <= 1:
        return
    # New slides are inserted before the template outro, so the final match is
    # the canonical ending page. Remove any earlier duplicates in reverse order.
    for slide in reversed(outro_slides[:-1]):
        _remove_slide(presentation, slide)


def _clear_paragraph_runs(paragraph, value: str) -> None:
    runs = paragraph.runs
    if runs:
        runs[0].text = value
        for run in runs[1:]:
            run.text = ""
    else:
        paragraph.text = value


def _copy_relationships(source_slide, target_slide, element) -> None:
    """复制图片/外链关系，避免跨 PPTX 的 rId 悬空。"""
    for node in element.iter():
        for attribute, source_rid in list(node.attrib.items()):
            if not attribute.startswith("{" + qn("r:id").split("}")[0].strip("{") + "}"):
                continue
            try:
                source_rel = source_slide.part.rels[source_rid]
            except KeyError:
                continue
            if source_rel.reltype == RT.IMAGE:
                _, target_rid = target_slide.part.get_or_add_image_part(BytesIO(source_rel.target_part.blob))
            elif source_rel.is_external:
                target_rid = target_slide.part.relate_to(source_rel.target_ref, source_rel.reltype, is_external=True)
            elif source_rel.reltype.endswith("/tags"):
                # 标签是 PowerPoint 的可选元数据，跨文件复制没有视觉影响。
                parent = node.getparent()
                if parent is not None:
                    parent.remove(node)
                continue
            else:
                raise ValueError(f"无法复制 {source_rel.reltype} 关系，请先在源 PPT 中展开对象")
            node.set(attribute, target_rid)


def _embedded_package_relationships(slide, shapes=None) -> list[str]:
    """返回幻灯片中需要丢弃的嵌入式 package 关系类型。"""
    package_reltype = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
    relationship_types = []
    shape_list = shapes if shapes is not None else slide.shapes
    for shape in shape_list:
        for node in shape._element.iter():
            source_rid = node.attrib.get(qn("r:id"))
            if not source_rid:
                continue
            source_rel = slide.part.rels.get(source_rid)
            if source_rel is None or source_rel.reltype != package_reltype:
                continue
            if source_rel.reltype not in relationship_types:
                relationship_types.append(source_rel.reltype)
    return relationship_types


def _matching_layout(presentation: Presentation, source_slide):
    source_name = getattr(source_slide.slide_layout, "name", "")
    fallback = presentation.slide_layouts[-1]
    return next((layout for layout in presentation.slide_layouts if layout.name == source_name), fallback)


def _clone_source_slide(presentation: Presentation, source_slide):
    target_slide = presentation.slides.add_slide(_matching_layout(presentation, source_slide))
    for shape in list(target_slide.shapes):
        _remove_element(shape._element)
    for shape in source_slide.shapes:
        if any(
            reltype == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
            for reltype in _embedded_package_relationships(source_slide, [shape])
        ):
            # python-pptx 无法安全复制 OLE/package 关系；丢弃整个对象，避免留下悬空 r:id。
            continue
        element = deepcopy(shape._element)
        _copy_relationships(source_slide, target_slide, element)
        target_slide.shapes._spTree.insert_element_before(element, "p:extLst")
    return target_slide


def _text_role(text: str) -> str:
    if PPT_HEADING.match(_text(text)):
        return "title"
    for key in SECTION_LABELS:
        if _is_section_label(text, key):
            return key
    if _section_label(text, _supplementary_key("项目概览")):
        return _supplementary_key("项目概览")
    return "body"


def _copy_xml_contents(target, source) -> None:
    target.attrib.clear()
    target.attrib.update(source.attrib)
    for child in list(target):
        target.remove(child)
    for child in source:
        target.append(deepcopy(child))


def _copy_text_style(source_shape, template_shape) -> None:
    """复制模板的段落与字体样式，保留源文本框几何和自动适配设置。"""
    source_frame = source_shape.text_frame
    template_frame = template_shape.text_frame
    source_paragraphs = source_frame.paragraphs
    template_paragraphs = template_frame.paragraphs
    if not template_paragraphs:
        return
    for paragraph_index, source_paragraph in enumerate(source_paragraphs):
        template_paragraph = template_paragraphs[min(paragraph_index, len(template_paragraphs) - 1)]
        template_ppr = template_paragraph._p.find(qn("a:pPr"))
        if template_ppr is not None:
            source_ppr = source_paragraph._p.get_or_add_pPr()
            _copy_xml_contents(source_ppr, template_ppr)
        else:
            source_ppr = source_paragraph._p.find(qn("a:pPr"))
            if source_ppr is not None:
                source_paragraph._p.remove(source_ppr)
        template_run = next(iter(template_paragraph.runs), None)
        if template_run is None:
            continue
        template_rpr = template_run._r.find(qn("a:rPr"))
        for source_run in source_paragraph.runs:
            if template_rpr is not None:
                source_rpr = source_run._r.get_or_add_rPr()
                _copy_xml_contents(source_rpr, template_rpr)
            else:
                source_rpr = source_run._r.find(qn("a:rPr"))
                if source_rpr is not None:
                    source_run._r.remove(source_rpr)


def _apply_template_text_format(source_slide, template_slide) -> None:
    source_entries = [entry for entry in _flatten_shapes(source_slide.shapes) if getattr(entry["shape"], "has_text_frame", False) and entry["text"]]
    template_entries = [entry for entry in _flatten_shapes(template_slide.shapes) if getattr(entry["shape"], "has_text_frame", False)]
    if not source_entries or not template_entries:
        return
    for source_entry in source_entries:
        template_entry = _matching_template_entry(source_entry, template_entries)
        if template_entry is None:
            continue
        _copy_text_style(source_entry["shape"], template_entry["shape"])


def _shape_font_size(shape, default: float = 12.0) -> float:
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if run.font.size is not None:
                return max(6.0, float(run.font.size.pt))
    return default


def _shape_capacity(shape) -> tuple[int, int]:
    """估算当前字号下文本框可容纳的字符数和行数。"""
    font_size = _shape_font_size(shape)
    emu_per_point = 12700
    chars_per_line = max(8, int(int(shape.width) / (font_size * emu_per_point * 1.08)))
    max_lines = max(1, int(int(shape.height) / (font_size * emu_per_point * 1.3)))
    return chars_per_line, max_lines


def _fit_body_font(shape, minimum: int = 9) -> None:
    text = _text(shape.text)
    if not text:
        return
    current_size = _shape_font_size(shape)
    start = max(minimum, int(round(current_size)))
    for size in range(start, minimum - 1, -1):
        chars_per_line, max_lines = _shape_capacity_for_font(shape, size)
        if len(text) <= chars_per_line * max_lines:
            if size < current_size:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(size)
            return
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(minimum)


def _has_paragraph_hierarchy(shape) -> bool:
    """从源段落格式识别层级，不根据标题文字猜测结构。"""
    paragraphs = [paragraph for paragraph in shape.text_frame.paragraphs if _text(paragraph.text)]
    if len(paragraphs) < 2:
        return False
    levels = {paragraph.level for paragraph in paragraphs}
    margins = set()
    bold_paragraphs = set()
    for paragraph in paragraphs:
        properties = paragraph._p.find(qn("a:pPr"))
        margins.add(int(properties.get("marL", "0")) if properties is not None else 0)
        runs = [run for run in paragraph.runs if _text(run.text)]
        bold_paragraphs.add(bool(runs) and all(
            run.font.bold if run.font.bold is not None else paragraph.font.bold
            for run in runs
        ))
    return len(levels) > 1 or len(margins) > 1 or len(bold_paragraphs) > 1


def _format_project_body(shape) -> None:
    """Apply the project-report body formatting."""
    if _has_paragraph_hierarchy(shape):
        # 源 PPT 的标题、子项和软换行保留在原段落中，不能重建成纯文本列表。
        return
    # Preserve tighter source-page margins when the original author already
    # fitted the content successfully; only reduce unusually large margins.
    if shape.text_frame.margin_left is None or shape.text_frame.margin_left > PPT_BODY_MARGIN_LEFT:
        shape.text_frame.margin_left = PPT_BODY_MARGIN_LEFT
    if shape.text_frame.margin_top is None or shape.text_frame.margin_top > PPT_BODY_MARGIN_TOP:
        shape.text_frame.margin_top = PPT_BODY_MARGIN_TOP
    lines = []
    for paragraph in shape.text_frame.paragraphs:
        # 段落内的垂直制表符是同一条内容的软换行；只在段落首部添加一次圆点，
        # 续行保持原样，避免每一行都被当成新的独立条目。
        raw_paragraph = _text(paragraph.text)
        continuation_only = raw_paragraph.startswith(CONTINUATION_MARKER)
        raw_paragraph = raw_paragraph.lstrip(CONTINUATION_MARKER)
        paragraph_lines = [
            _clean_line(part)
            for part in re.split(r"[\v]+", raw_paragraph)
        ]
        paragraph_lines = [part for part in paragraph_lines if part]
        if paragraph_lines:
            value = "\v".join(paragraph_lines)
            has_own_marker = _has_ordered_prefix(paragraph_lines[0])
            lines.append(value if continuation_only or has_own_marker else f"· {value}")
    if lines and "\n".join(lines) != shape.text:
        shape.text = "\n".join(lines)
    shape.text_frame.word_wrap = True
    for paragraph in shape.text_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.LEFT
        paragraph.level = 0
        for run in paragraph.runs:
            _set_ppt_run_font(run)


def _set_ppt_run_font(run, font_name: str = PPT_BODY_FONT, bold: bool | None = None) -> None:
    run.font.name = font_name
    if bold is not None:
        run.font.bold = bold
    rpr = run._r.get_or_add_rPr()
    for tag_name in ("latin", "ea", "cs"):
        element = rpr.find(qn(f"a:{tag_name}"))
        if element is None:
            element = rpr.makeelement(qn(f"a:{tag_name}"), {})
            rpr.append(element)
        element.set("typeface", font_name)


def _set_project_body_font(shape, font_size: float, color: RGBColor | None = None) -> None:
    effective_color = color or RGBColor(0, 0, 0)
    preserve_hierarchy = _has_paragraph_hierarchy(shape)
    for paragraph in shape.text_frame.paragraphs:
        if not preserve_hierarchy:
            paragraph.alignment = PP_ALIGN.LEFT
            paragraph.level = 0
            _clear_paragraph_bullets(paragraph)
        for run in paragraph.runs:
            _set_ppt_run_font(run)
            run.font.size = Pt(font_size)
            run.font.color.rgb = effective_color


def _set_shape_text_color(shape, color: RGBColor) -> None:
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.color.rgb = color


def _apply_issue_text_color(slide, issue_text: str) -> None:
    """Style actual supplementary regions, never matching shared body words."""
    entries = _flatten_shapes(slide.shapes)
    for entry in entries:
        shape = entry["shape"]
        if not getattr(shape, "has_text_frame", False):
            continue
        key = _entry_section_key(entry)
        if _is_supplementary_key(key):
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(255, 255, 0) if _section_is_issue(key) else RGBColor(20, 68, 216)
            _set_shape_text_color(shape, PPT_ISSUE_COLOR if _section_is_issue(key) else RGBColor(255, 255, 255))
            continue
        if _entry_role(entry) == "body":
            body_key = _body_section_key(entries, entry)
            _set_shape_text_color(shape, PPT_ISSUE_COLOR if _section_is_issue(body_key) else RGBColor(0, 0, 0))


def _project_body_font_size(shape, *, preserve_paragraphs: bool = False) -> float:
    if not preserve_paragraphs:
        _format_project_body(shape)
    # Avoid creating an almost-empty continuation slide when a source page
    # only exceeds the conservative estimate by one or two lines.
    # 以模板标准字号为起点，不能让源 PPT 的小字号把最终正文锁死在小号。
    for font_size in (14, 13, 12, 11, 10.5):
        _set_project_body_font(shape, font_size)
        if len(_text_chunks(shape)) <= 1:
            return font_size
    # 文本超出单页时交给分页逻辑处理，不把正文压缩到难以阅读的小字号。
    _set_project_body_font(shape, 10.5)
    return 10.5


def _clear_paragraph_bullets(paragraph) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    for child in list(ppr):
        if child.tag.rsplit("}", 1)[-1] in {"buNone", "buChar", "buAutoNum", "buBlip", "buFont"}:
            ppr.remove(child)
    ppr.set("marL", "0")
    ppr.set("indent", "0")


def _body_section_key(entries: list[dict[str, Any]], body_entry: dict[str, Any]) -> str | None:
    labels = sorted(
        (entry for entry in entries if _entry_section_key(entry)),
        key=lambda entry: (entry["top"], entry["left"]),
    )
    # A right-column label may be on either side of its body. Its name does
    # not determine whether this region is an issue, overview or other content.
    right_labels = [entry for entry in labels
                    if _is_supplementary_key(_entry_section_key(entry)) and entry["left"] >= 6000000]
    if body_entry["left"] >= 6000000 and right_labels:
        candidates = []
        for label in right_labels:
            next_tops = [other["top"] for other in right_labels if other["top"] > label["top"] + 100000]
            if label["top"] - 100000 <= body_entry["top"] < min(next_tops, default=10**10):
                candidates.append(label)
        if candidates:
            label = min(candidates, key=lambda item: (
                abs(item["top"] - body_entry["top"]), abs(item["left"] - body_entry["left"])
            ))
            return _entry_section_key(label)
        return None
    if body_entry["left"] > 7200000:
        return None
    candidates = []
    column_labels = [label for label in labels if label["left"] < 6000000]
    for label in column_labels:
        if body_entry["left"] + 100000 < label["left"]:
            continue
        next_top = min((other["top"] for other in column_labels
                        if other["top"] > label["top"] + 100000), default=10**10)
        if label["top"] - 100000 <= body_entry["top"] < next_top:
            candidates.append(label)
    if not candidates:
        return None
    label = min(candidates, key=lambda entry: abs(entry["top"] - body_entry["top"]))
    return _entry_section_key(label)


def _section_layout_parts(slide, entries: list[dict[str, Any]], key: str) -> dict[str, Any] | None:
    """Return the editable label, border and body used by one standard module."""
    labels = [entry for entry in entries if _entry_section_key(entry) == key]
    bodies = [
        entry for entry in entries
        if entry["text"]
        and _entry_role(entry) == "body"
        and _body_section_key(entries, entry) == key
    ]
    if not labels or not bodies:
        return None
    label = min(labels, key=lambda entry: (entry["top"], entry["left"]))
    body = max(bodies, key=lambda entry: len(entry["text"]))
    top_level_label = slide.shapes[int(label["path"].split(".", 1)[0]) - 1]
    borders = [
        entry for entry in entries
        if "." not in entry["path"]
        and not entry["text"]
        and entry["shape"].shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
        and abs(entry["top"] - body["top"]) <= 150000
        and abs(entry["height"] - body["height"]) <= 150000
        and entry["left"] <= body["left"]
        and entry["left"] + entry["width"] >= body["left"] + body["width"] - 250000
    ]
    border = min(borders, key=lambda entry: abs(entry["left"] - body["left"]))["shape"] if borders else None
    return {
        "label": top_level_label,
        "border": border,
        "body": body["shape"],
    }


def _flow_content_key(value: str) -> str:
    """Ignore display bullets/whitespace, without removing repeated content."""
    return "".join(
        re.sub(r"\s+", "", _clean_line(line.lstrip(CONTINUATION_MARKER)))
        for line in re.split(r"[\r\n\v]+", value)
    )


def _flow_section_shapes(slide) -> dict[str, dict[str, Any]]:
    """Include empty template body boxes: their geometry defines usable space."""
    entries = _flatten_shapes(slide.shapes)
    labels = [entry for entry in entries if _entry_section_key(entry)]
    sections = {key: {"label": None, "bodies": []} for key in ("current", "next")}
    for entry in labels:
        key = _entry_section_key(entry)
        sections.setdefault(key, {"label": None, "bodies": []})
        if sections[key]["label"] is None:
            sections[key]["label"] = entry["shape"]
    for entry in entries:
        if (entry["top"] < 900000
                or _entry_role(entry) != "body"
                or not getattr(entry["shape"], "has_text_frame", False)):
            continue
        key = _body_section_key(entries, entry)
        if key is None and labels and entry["left"] < 6000000:
            # Some submitted decks place horizontal labels partly over the body.
            label = min(labels, key=lambda item: (
                abs(item["top"] - entry["top"]), abs(item["left"] - entry["left"])
            ))
            key = _entry_section_key(label)
        if key:
            sections[key]["bodies"].append(entry["shape"])
    return sections


def _ppt_body_line_height(font_size: float) -> int:
    return Centipoints(round(2200 * font_size / 14))


def _trim_flow_paragraphs(paragraphs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Remove padding paragraphs/breaks without touching interior whitespace."""
    paragraphs = list(paragraphs)
    while paragraphs and not paragraphs[0]["text"].strip():
        paragraphs.pop(0)
    while paragraphs and not paragraphs[-1]["text"].strip():
        paragraphs.pop()
    if not paragraphs:
        return paragraphs
    for index in dict.fromkeys((0, len(paragraphs) - 1)):
        item = paragraphs[index]
        text = item["text"]
        leading = re.match(r"^(?:[ \t]*\v)+", text) if index == 0 else None
        trailing = re.search(r"(?:\v[ \t]*)+$", text) if index == len(paragraphs) - 1 else None
        start, end = leading.end() if leading else 0, trailing.start() if trailing else len(text)
        if start or end != len(text):
            paragraphs[index] = _flow_paragraph_slice(item, start, end, continuation=False, trim_break=False)
    return paragraphs


def _trim_body_whitespace(frame) -> None:
    paragraphs = _trim_flow_paragraphs([
        {"text": paragraph.text, "xml": paragraph._p} for paragraph in frame.paragraphs
    ])
    for element in list(frame._txBody.findall(qn("a:p"))):
        frame._txBody.remove(element)
    for item in paragraphs:
        frame._txBody.append(item["xml"])
    if not paragraphs:
        frame._txBody.append(OxmlElement("a:p"))


def _fit_compact_slide(slide) -> tuple[bool, dict[str, dict[str, Any]]]:
    """Fit each actual output box, keeping the largest readable candidate size."""
    fits = True
    fitted = {}
    for key, parts in _flow_section_shapes(slide).items():
        bodies = list(parts["bodies"])
        label = parts["label"]
        if not bodies and (label is None or not _content_after_section_label(label.text, key)):
            # A missing/empty inline region still needs a visible "无" box.
            fits = False
        if label is not None and _is_supplementary_key(key):
            vertical = label.text_frame._txBody.bodyPr.get("vert", "horz") != "horz" or label.width < Pt(60)
            if vertical and label.height < _flow_label_height(key):
                fits = False
        if label is not None and _content_after_section_label(label.text, key):
            fits = False
            bodies.append(label)
        for shape in bodies:
            _trim_body_whitespace(shape.text_frame)
            if not _text(shape.text) or str(shape.shape_id) in fitted:
                continue
            frame = shape.text_frame
            frame.word_wrap = True
            frame.auto_size = MSO_AUTO_SIZE.NONE
            frame.vertical_anchor = MSO_ANCHOR.TOP
            shape_fits = False
            for font_size in PPT_BODY_FONT_STEPS:
                line_height = _ppt_body_line_height(font_size)
                for paragraph in frame.paragraphs:
                    paragraph.alignment = PP_ALIGN.LEFT
                    _format_flow_paragraph(paragraph, line_height,
                                           PPT_ISSUE_COLOR if _section_is_issue(key) else RGBColor(0, 0, 0), font_size)
                format_ok, required = _text_shape_metrics(shape, font_size, line_height)
                if format_ok and required is not None and required <= shape.height:
                    shape_fits = True
                    break
            fits = fits and shape_fits
            fitted[str(shape.shape_id)] = {
                "font_size": font_size, "line_height": line_height, "section": key,
                "text": shape.text, "xml": deepcopy(shape._element),
            }
    return fits, fitted


def _flow_paragraphs(project: dict[str, Any], source_slides: list[Any]) -> dict[str, list[Any]]:
    """Match audited text to source paragraphs to retain run and list formatting."""
    output_sections = _output_sections(project)
    catalog = {key: [] for key in output_sections}
    for slide_index, slide in enumerate(source_slides):
        for key, parts in _flow_section_shapes(slide).items():
            if key not in catalog:
                continue
            for shape in parts["bodies"]:
                items = _trim_flow_paragraphs([
                    {"text": paragraph.text, "xml": deepcopy(paragraph._p), "source_slide": slide}
                    for paragraph in shape.text_frame.paragraphs
                ])
                blanks, previous = [], None
                for item_index, item in enumerate(items):
                    if not item["text"].strip():
                        blanks.append(item)
                        continue
                    cleaner = _clean_supplementary_content if _is_supplementary_key(key) else _clean_content
                    cleaned = cleaner(item["text"])
                    token = (slide_index, shape.shape_id, item_index)
                    if cleaned and _flow_content_key(cleaned) == _flow_content_key(item["text"]):
                        catalog[key].append({"lines": cleaned.split("\n"), "paragraph": item,
                                             "token": token, "previous": previous, "blanks": blanks})
                    previous, blanks = token, []
    sections = {}
    for key, section in output_sections.items():
        lines = section["content"].split("\n")
        paragraphs = []
        index = 0
        previous = None
        while index < len(lines):
            if not lines[index].strip() and not _is_supplementary_key(key):
                index += 1
                continue
            matches = [item for item in catalog[key]
                       if item["lines"] == lines[index:index + len(item["lines"])]]
            match = next((item for item in matches if item["previous"] == previous), matches[0] if matches else None)
            if match:
                if previous is not None and match["previous"] == previous and not _is_supplementary_key(key):
                    paragraphs.extend(match["blanks"])
                paragraphs.append(match["paragraph"])
                previous = match["token"]
                index += len(match["lines"])
            else:
                element = OxmlElement("a:p")
                run = OxmlElement("a:r")
                text = OxmlElement("a:t")
                text.text = lines[index]
                run.append(text)
                element.append(run)
                paragraphs.append({"text": lines[index], "xml": element})
                index += 1
                previous = None
        sections[key] = paragraphs
    return sections


def _flow_character_em(character: str) -> float:
    """Stable width estimates; do not depend on the machine's fallback font."""
    if unicodedata.category(character) in {"Mn", "Me", "Cf"}:
        return 0.0
    if unicodedata.east_asian_width(character) in {"W", "F"}:
        return 1.0
    if character.isspace():
        return 0.35
    if character in "ilI.,:;!'`|()[]{}":
        return 0.4
    if character in "MWmw@#%&":
        return 1.0
    if character.isascii():
        return 0.7 if character.isupper() else 0.62
    return 1.0


def _flow_line_ends(
    paragraph: dict[str, Any],
    width: int,
    margin_left: int = PPT_BODY_MARGIN_LEFT,
    margin_right: int = PPT_BODY_MARGIN_LEFT,
    font_size: float = PPT_FLOW_FONT_SIZE,
) -> list[int]:
    """Return character offsets at visual line ends, retaining soft line breaks."""
    properties = paragraph["xml"].find(qn("a:pPr"))
    level_indent = int(properties.get("lvl", "0")) * Pt(font_size) if properties is not None else 0
    left_indent = int(properties.get("marL", str(level_indent))) if properties is not None else 0
    right_indent = int(properties.get("marR", "0")) if properties is not None else 0
    first_indent = int(properties.get("indent", "0")) if properties is not None else 0
    usable = width - margin_left - margin_right - max(0, left_indent) - max(0, right_indent)
    text = paragraph["text"]
    em = Pt(font_size)
    advances = [int(em * _flow_character_em(character) * 1.08) for character in text]
    # Source tracking is in hundredths of a point. Keep it in the estimate as
    # well as the output, including inherited tracking on paragraph defaults.
    defaults = properties.find(qn("a:defRPr")) if properties is not None else None
    tracking = defaults.get("spc", "0") if defaults is not None else "0"
    offset = 0
    for child in paragraph["xml"]:
        if child.tag not in {qn("a:r"), qn("a:fld"), qn("a:br")}:
            continue
        value = child.find(qn("a:t"))
        length = 1 if child.tag == qn("a:br") else len(value.text or "") if value is not None else 0
        run_properties = child.find(qn("a:rPr"))
        spacing = Centipoints(int(run_properties.get("spc", tracking) if run_properties is not None else tracking))
        for index in range(offset, min(offset + length, len(advances))):
            if advances[index]:
                advances[index] = max(0, advances[index] + spacing)
        offset += length
    tab_list = properties.find(qn("a:tabLst")) if properties is not None else None
    tab_stops = sorted(int(tab.get("pos", "0")) for tab in tab_list) if tab_list is not None else []
    tab_size = max(1, int(paragraph.get("tab_size", Pt(36))))
    ends, offset, line_start, used = [], 0, 0, 0
    while offset < len(text):
        character = text[offset]
        if character in "\v\n\r":
            offset += 1
            ends.append(offset)
            line_start, used = offset, 0
            continue
        indent = max(0, first_indent) if not ends else 0
        line_width = max(em, usable - indent)
        # Keep an English word together if it fits on an empty line; long
        # identifiers still wrap by character. Offsets always refer to source XML.
        if (character.isascii() and (character.isalnum() or character == "_")
                and (offset == 0 or not (text[offset - 1].isascii()
                                        and (text[offset - 1].isalnum() or text[offset - 1] == "_")))):
            word = re.match(r"[A-Za-z0-9_]+", text[offset:])
            word_width = sum(advances[offset:offset + len(word.group())])
            if used and word_width <= max(em, usable) and used + word_width > line_width:
                ends.append(offset)
                line_start, used = offset, 0
                continue
        advance = advances[offset]
        if character == "\t":
            position = max(0, left_indent) + indent + used
            stop = next((stop for stop in tab_stops if stop > position), (position // tab_size + 1) * tab_size)
            advance = stop - position
        if offset > line_start and advance > 0 and used + advance > line_width:
            # Do not start a line with closing punctuation. Move its preceding
            # character too, but never rewind an entire line or split a word.
            cut = offset
            if character in "，。、；：！？）》】”’,.!?;:%)]}" and offset - line_start > 1:
                candidate = offset - 1
                while candidate > line_start and advances[candidate] == 0:
                    candidate -= 1
                if (candidate > line_start and not text[candidate].isascii()
                        and text[candidate] not in "（《【“‘"):
                    cut = candidate
            ends.append(cut)
            offset, line_start, used = cut, cut, 0
            continue
        used += advance
        offset += 1
    ends.append(len(text))
    return ends


def _flow_paragraph_slice(
    paragraph: dict[str, Any], start: int, end: int, *, continuation: bool = True, trim_break: bool = True,
) -> dict[str, Any]:
    # A consumed soft break is a page boundary, not an extra empty rendered line.
    if trim_break and end > start and paragraph["text"][end - 1] == "\v":
        end -= 1
    element = deepcopy(paragraph["xml"])
    cursor = 0
    for child in list(element):
        if child.tag in {qn("a:pPr"), qn("a:endParaRPr")}:
            continue
        text = child.find(qn("a:t"))
        value = "\v" if child.tag == qn("a:br") else (text.text or "") if text is not None else ""
        stop = cursor + len(value)
        left, right = max(start, cursor), min(end, stop)
        if left >= right:
            element.remove(child)
        elif text is not None:
            text.text = value[left - cursor:right - cursor]
        cursor = stop
    if start and continuation:
        properties = element.find(qn("a:pPr"))
        if properties is None:
            properties = OxmlElement("a:pPr")
            element.insert(0, properties)
        for child in list(properties):
            if child.tag.rsplit("}", 1)[-1] in {"buNone", "buChar", "buAutoNum", "buBlip"}:
                properties.remove(child)
        properties.append(OxmlElement("a:buNone"))
        properties.set("indent", "0")
    return {**paragraph, "text": paragraph["text"][start:end], "xml": element,
            "continued": paragraph.get("continued", False) or bool(start and continuation)}


def _flow_is_heading(paragraph: dict[str, Any]) -> bool:
    text = paragraph["text"].strip()
    if (paragraph.get("continued") or not text or len(text) > 40
            or "\v" in text or text.endswith(tuple("。！？；.!?;"))):
        return False
    if re.match(r"^[一二三四五六七八九十百]+[、．.]", text):
        return True
    properties = paragraph["xml"].find(qn("a:pPr"))
    if _has_ordered_prefix(text) or (properties is not None and (
        int(properties.get("lvl", "0")) > 0
        or any(properties.find(qn(f"a:{tag}")) is not None for tag in ("buChar", "buAutoNum", "buBlip"))
    )):
        return False
    defaults = properties.find(qn("a:defRPr")) if properties is not None else None
    default_bold = defaults.get("b", "0") if defaults is not None else "0"
    runs = [child for child in paragraph["xml"] if child.tag in {qn("a:r"), qn("a:fld")}
            and child.find(qn("a:t")) is not None and (child.find(qn("a:t")).text or "").strip()]
    return bool(runs) and all(
        (child.find(qn("a:rPr")).get("b", default_bold) if child.find(qn("a:rPr")) is not None
         else default_bold) in {"1", "true"} for child in runs
    )


def _flow_label_height(key: str) -> int:
    return max(4, len(_section_title(key))) * PPT_FLOW_LABEL_LINE_HEIGHT + 2 * PPT_BODY_MARGIN_TOP


def _flow_page_minimums(page: list[dict[str, Any]], width: int, font_size: float) -> tuple[list[int], list[int]]:
    line_counts = [sum(len(_flow_line_ends(item, width, font_size=font_size)) for item in fragment["paragraphs"])
                   for fragment in page]
    padding = 2 * PPT_BODY_MARGIN_TOP + PPT_FLOW_BOTTOM_SAFETY
    minimums = [max(_flow_label_height(fragment["section"]), lines * _ppt_body_line_height(font_size) + padding)
                for fragment, lines in zip(page, line_counts)]
    return minimums, line_counts


def _flow_append(page: list[dict[str, Any]], items: list[tuple[str, dict[str, Any]]]) -> list[dict[str, Any]]:
    """Copy the cheap page structure, never deep-copy source slide objects."""
    page = [{**fragment, "paragraphs": list(fragment["paragraphs"])} for fragment in page]
    for key, paragraph in items:
        if not page or page[-1]["section"] != key:
            page.append({"section": key, "paragraphs": []})
        page[-1]["paragraphs"].append(paragraph)
    return page


def _flow_required_height(key: str, paragraphs: list[dict[str, Any]], width: int, font_size: float) -> int:
    lines = sum(len(_flow_line_ends(item, width, font_size=font_size)) for item in paragraphs)
    return max(_flow_label_height(key),
               lines * _ppt_body_line_height(font_size) + 2 * PPT_BODY_MARGIN_TOP + PPT_FLOW_BOTTOM_SAFETY)


def _flow_fragment(key: str, paragraphs: list[dict[str, Any]], rectangle: tuple[int, int, int, int],
                   geometry: dict[str, int], font_size: float) -> dict[str, Any]:
    left, top, width, height = rectangle
    body_left = left + geometry["label_width"] + PPT_FLOW_GAP
    body_width = width - geometry["label_width"] - PPT_FLOW_GAP
    return {
        "section": key, "paragraphs": list(paragraphs), "left": left, "top": top, "height": height,
        "label_width": geometry["label_width"], "body_left": body_left, "body_width": body_width,
        "font_size": font_size, "line_height": _ppt_body_line_height(font_size),
        "line_count": sum(len(_flow_line_ends(item, body_width, font_size=font_size)) for item in paragraphs),
    }


def _standard_flow_page(sections: dict[str, list[Any]], geometry: dict[str, int]) -> list[dict[str, Any]] | None:
    """Choose a readable single page, preferring a full-height right issue column on ties."""
    keys = set(sections)
    if keys not in ({"current", "next", "issues"}, {"current", "issues"}, {"next", "issues"}):
        return None
    left, top = geometry["left"], geometry["top"]
    width = geometry["body_left"] + geometry["body_width"] - left
    height = geometry["bottom"] - top
    inset = geometry["label_width"] + PPT_FLOW_GAP
    for font_size in PPT_BODY_FONT_STEPS:
        if len(keys) == 2:
            primary = "current" if "current" in keys else "next"
            for fraction in PPT_ISSUE_WIDTH_FRACTIONS:
                right_width = round((width - PPT_FLOW_GAP) * fraction)
                left_width = width - PPT_FLOW_GAP - right_width
                if min(right_width, left_width) - inset < Pt(24):
                    continue
                if (_flow_required_height(primary, sections[primary], left_width - inset, font_size) <= height
                        and _flow_required_height("issues", sections["issues"], right_width - inset, font_size) <= height):
                    return [
                        _flow_fragment(primary, sections[primary], (left, top, left_width, height), geometry, font_size),
                        _flow_fragment("issues", sections["issues"],
                                       (left + left_width + PPT_FLOW_GAP, top, right_width, height), geometry, font_size),
                    ]
            continue
        candidates = []
        # Prefer a readable 35% issue region; narrow it only when the main
        # content needs more width at the same font size.
        for mode, fractions in (("side", PPT_ISSUE_WIDTH_FRACTIONS),
                                ("wide", (*PPT_ISSUE_WIDTH_FRACTIONS, 0.40))):
            for width_rank, fraction in enumerate(fractions):
                right_width = round((width - PPT_FLOW_GAP) * fraction)
                left_width = width - PPT_FLOW_GAP - right_width
                if min(right_width, left_width) - inset < Pt(24):
                    continue
                current_width = left_width if mode == "side" else width
                current_min = _flow_required_height("current", sections["current"], current_width - inset, font_size)
                next_min = _flow_required_height("next", sections["next"], left_width - inset, font_size)
                issue_min = _flow_required_height("issues", sections["issues"], right_width - inset, font_size)
                bottom_min = next_min if mode == "side" else max(next_min, issue_min)
                upper_limit = height - PPT_FLOW_GAP - bottom_min
                if current_min > upper_limit or (mode == "side" and issue_min > height):
                    continue
                current_height = min(max(round((height - PPT_FLOW_GAP) * 0.70), current_min), upper_limit)
                bottom_top = top + current_height + PPT_FLOW_GAP
                bottom_height = height - current_height - PPT_FLOW_GAP
                right_left = left + left_width + PPT_FLOW_GAP
                rectangles = {
                    "current": (left, top, current_width, current_height),
                    "next": (left, bottom_top, left_width, bottom_height),
                    "issues": (right_left, top if mode == "side" else bottom_top,
                               right_width, height if mode == "side" else bottom_height),
                }
                page = [_flow_fragment(key, sections[key], rectangles[key], geometry, font_size)
                        for key in ("current", "next", "issues")]
                candidates.append(((mode != "side", width_rank), page))
        if candidates:
            return min(candidates, key=lambda candidate: candidate[0])[1]
    return None


def _position_flow_pages(pages: list[list[dict[str, Any]]], geometry: dict[str, int]) -> list[list[dict[str, Any]]]:
    index = 0
    for page in pages:
        for fragment in page:
            for key in ("left", "label_width", "body_left", "body_width"):
                fragment.setdefault(key, geometry[key])
            fragment["index"] = index
            index += 1
    return pages


def _expand_flow_page(page: list[dict[str, Any]], geometry: dict[str, int]) -> None:
    """Return unused footer/side space to earlier pages without changing their text."""
    if not page:
        return
    minimums, counts = _flow_page_minimums(page, geometry["body_width"], page[0]["font_size"])
    extra = geometry["bottom"] - geometry["top"] - PPT_FLOW_GAP * (len(page) - 1) - sum(minimums)
    if extra < 0:
        raise ValueError("周报正文超出可用页面区域")
    total = max(1, sum(counts))
    top, distributed = geometry["top"], 0
    for index, (fragment, minimum, lines) in enumerate(zip(page, minimums, counts)):
        share = extra - distributed if index == len(page) - 1 else extra * lines // total
        fragment.update(top=top, height=minimum + share, line_count=lines)
        for key in ("left", "label_width", "body_left", "body_width"):
            fragment[key] = geometry[key]
        top += fragment["height"] + PPT_FLOW_GAP
        distributed += share


def _flow_pages_with_empty_sections(
    sections: dict[str, list[Any]], geometry: dict[str, int],
) -> list[list[dict[str, Any]]] | None:
    """Reserve placeholder space before pagination; no placeholder can create its own page."""
    empty = {key: items for key, items in sections.items()
             if _is_empty_section_content("\n".join(item["text"] for item in items))}
    content = {key: items for key, items in sections.items() if key not in empty}
    if not empty or not content:
        return None
    left, top, bottom = geometry["left"], geometry["top"], geometry["bottom"]
    width = geometry["body_left"] + geometry["body_width"] - left
    height = bottom - top
    inset = geometry["label_width"] + PPT_FLOW_GAP
    candidates = []
    for mode in ("side", "footer"):
        reduced = dict(geometry)
        rectangles = []
        if mode == "side":
            reserved = max(round(width * 0.15), inset + Pt(24))
            box_height = (height - PPT_FLOW_GAP * (len(empty) - 1)) // len(empty)
            if any(box_height < _flow_label_height(key) for key in empty):
                continue
            reduced["body_width"] -= reserved + PPT_FLOW_GAP
            for index, key in enumerate(empty):
                rectangles.append((left + width - reserved, top + index * (box_height + PPT_FLOW_GAP),
                                   reserved, box_height))
        else:
            reserved = max(_flow_label_height(key) for key in empty)
            reduced["bottom"] -= reserved + PPT_FLOW_GAP
            box_width = (width - PPT_FLOW_GAP * (len(empty) - 1)) // len(empty)
            if box_width - inset < Pt(24):
                continue
            for index, key in enumerate(empty):
                rectangles.append((left + index * (box_width + PPT_FLOW_GAP), bottom - reserved,
                                   box_width, reserved))
        if (reduced["body_width"] < Pt(100)
                or reduced["bottom"] - top < max(_flow_label_height(key) for key in content)):
            continue
        try:
            pages = _linear_flow_page_plan(content, reduced)
        except ValueError:
            # The alternate reserved region may still fit a heading and body
            # that cannot be split inside this narrower/shorter candidate.
            continue
        if not pages:
            continue
        _position_flow_pages(pages, reduced)
        for page in pages[:-1]:
            _expand_flow_page(page, geometry)
        pages[-1].extend(_flow_fragment(key, items, rectangle, geometry, pages[-1][0]["font_size"])
                         for (key, items), rectangle in zip(empty.items(), rectangles))
        tail_sections = {}
        for fragment in pages[-1]:
            tail_sections.setdefault(fragment["section"], []).extend(fragment["paragraphs"])
        repacked = _standard_flow_page(tail_sections, geometry)
        if repacked is not None:
            pages[-1] = repacked
        candidates.append(((len(pages), -min(item["font_size"] for page in pages for item in page), mode != "side"), pages))
    if not candidates:
        raise ValueError("周报模板没有足够空间将空栏目与实际正文放在同页")
    return min(candidates, key=lambda candidate: candidate[0])[1]


def _flow_page_plan(sections: dict[str, list[Any]], geometry: dict[str, int]) -> list[list[dict[str, Any]]]:
    page = _standard_flow_page(sections, geometry)
    if page is not None:
        return _position_flow_pages([page], geometry)
    pages = _flow_pages_with_empty_sections(sections, geometry)
    if pages is None:
        pages = _linear_flow_page_plan(sections, geometry)
    return _position_flow_pages(pages, geometry)


def _linear_flow_page_plan(sections: dict[str, list[Any]], geometry: dict[str, int]) -> list[list[dict[str, Any]]]:
    """Fit complete paragraphs at the largest size, then split at the floor."""
    top, bottom, width = geometry["top"], geometry["bottom"], geometry["body_width"]
    height = bottom - top
    if any(height < _flow_label_height(key) for key, items in sections.items() if items):
        raise ValueError("周报模板正文区域不足以容纳单列栏目")
    order = [key for key in ("current", "next") if key in sections]
    order.extend(key for key in sections if key not in {"current", "next"})
    pending = [(key, item) for key in order for item in _trim_flow_paragraphs(sections[key])]
    pages, page = [], []
    font_size, minimum_font = PPT_FLOW_FONT_SIZE, PPT_FLOW_FONT_STEPS[-1]

    def fits(candidate, size):
        minimums, _ = _flow_page_minimums(candidate, width, size)
        return sum(minimums) + PPT_FLOW_GAP * max(0, len(candidate) - 1) <= height

    def finish_page():
        nonlocal page, font_size
        if page:
            for fragment in page:
                fragment["font_size"] = font_size
                fragment["line_height"] = _ppt_body_line_height(font_size)
            pages.append(page)
        page, font_size = [], PPT_FLOW_FONT_SIZE

    while pending:
        key, paragraph = pending[0]
        unit_count, tail_pair = 1, False
        # A short plan followed only by "无" is one placement unit. Pulling just
        # the plan forward would create a whole page containing a single "无".
        if key == "next" and all(section == "next" or _section_is_issue(section) for section, _ in pending):
            plan = [item for section, item in pending if section == "next"]
            issues = [item for section, item in pending if _section_is_issue(section)]
            short_plan = sum(len(_flow_line_ends(item, width, font_size=minimum_font)) for item in plan) <= 4
            no_issues = bool(issues) and _is_empty_section_content("\n".join(item["text"] for item in issues))
            if short_plan and no_issues and fits(_flow_append([], pending), minimum_font):
                unit_count, tail_pair = len(pending), True
        if not tail_pair:
            # Keep a recognized heading (and any intervening blank paragraphs)
            # with the following text, including a chain of short headings.
            while (unit_count < len(pending) and pending[unit_count][0] == key
                   and (_flow_is_heading(pending[unit_count - 1][1])
                        or (unit_count > 1 and not pending[unit_count - 1][1]["text"].strip()))):
                unit_count += 1
        unit = pending[:unit_count]
        trial = _flow_append(page, unit)
        selected_font = next((size for size in PPT_FLOW_FONT_STEPS if size <= font_size and fits(trial, size)), None)
        if selected_font is not None:
            page, font_size = trial, selected_font
            del pending[:unit_count]
            continue

        # Recalculate the entire page at 9pt before attempting a continuation.
        # Splits must leave at least two visual lines on each side, and any
        # heading in this placement unit stays with the prefix.
        split_found = False
        if not tail_pair:
            body_key, body = unit[-1]
            ends = _flow_line_ends(body, width, font_size=minimum_font)
            max_prefix = min(len(ends) - 2, int(
                (height - 2 * PPT_BODY_MARGIN_TOP - PPT_FLOW_BOTTOM_SAFETY) // _ppt_body_line_height(minimum_font)
            ))
            for line_count in range(max_prefix, 1, -1):
                cut = ends[line_count - 1]
                prefix = _flow_paragraph_slice(body, 0, cut)
                suffix = _flow_paragraph_slice(body, cut, len(body["text"]), trim_break=False)
                if (len(_flow_line_ends(prefix, width, font_size=minimum_font)) < 2
                        or len(_flow_line_ends(suffix, width, font_size=minimum_font)) < 2):
                    continue
                candidate = _flow_append(page, unit[:-1] + [(body_key, prefix)])
                if fits(candidate, minimum_font):
                    page, font_size = candidate, minimum_font
                    del pending[:unit_count - 1]
                    pending[0] = (body_key, suffix)
                    split_found = True
                    break
        if not page:
            # A tiny/custom template may not have space for a heading and the
            # minimum two-line prefix. Fail explicitly rather than loop forever.
            raise ValueError("周报模板正文区域不足以容纳标题及正文续段")
        finish_page()
        if split_found:
            continue
        # Nothing was consumed: retry the intact unit on the next empty page.

    finish_page()
    # Only now stretch the selected content. Pagination and font selection must
    # never use the expanded boxes as their text's required height.
    fragment_index = 0
    for page in pages:
        minimums, line_counts = _flow_page_minimums(page, width, page[0]["font_size"])
        extra = height - PPT_FLOW_GAP * (len(page) - 1) - sum(minimums)
        if extra < 0:
            raise ValueError("周报单列分页高度超过模板正文区域")
        total_lines = sum(line_counts)
        y, distributed = top, 0
        for index, (item, minimum, lines) in enumerate(zip(page, minimums, line_counts)):
            share = extra - distributed if index == len(page) - 1 else extra * lines // total_lines
            item.update(index=fragment_index, top=y, height=minimum + share, line_count=lines)
            fragment_index += 1
            distributed += share
            y += item["height"] + PPT_FLOW_GAP
    return pages


def _prepare_flow_canvas(slide, presentation) -> tuple[dict[str, int], dict[str, Any]]:
    """Remove compact section furniture, retaining the header and branding."""
    sections = _flow_section_shapes(slide)
    shapes = [shape for parts in sections.values()
              for shape in ([parts["label"]] if parts["label"] is not None else []) + parts["bodies"]]
    # Several source labels may now share the canonical issue key. Clear all
    # of their furniture, even though only the first supplies its text style.
    for entry in _flatten_shapes(slide.shapes):
        if _entry_section_key(entry) and not any(entry["shape"]._element is shape._element for shape in shapes):
            shapes.append(entry["shape"])
    if not shapes:
        raise ValueError("周报项目页缺少可识别的正文区域")
    left = max(Pt(18), min(shape.left for shape in shapes))
    top = max(Pt(72), min(shape.top for shape in shapes))
    # Use matching page margins instead of inheriting a narrow source frame.
    right = presentation.slide_width - left
    bottom = min(presentation.slide_height - Pt(18), max(shape.top + shape.height for shape in shapes))
    labels = {key: deepcopy(parts["label"]._element) if parts["label"] is not None else None
              for key, parts in sections.items()}
    for key, parts in sections.items():
        label = parts["label"]
        if label is None:
            continue
        # Grouped templates often use a separate colored rectangle behind the
        # label textbox. Carry that fill over before removing the old furniture.
        backgrounds = [shape for shape in slide.shapes
                       if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
                       and not _text(getattr(shape, "text", ""))
                       and shape.left <= label.left + Pt(2) and shape.top <= label.top + Pt(2)
                       and shape.left + shape.width >= label.left + label.width - Pt(2)
                       and shape.top + shape.height >= label.top + label.height - Pt(2)
                       and shape.width <= label.width + Pt(16)]
        for background in sorted(backgrounds, key=lambda shape: shape.width * shape.height):
            fill = background._element.spPr.find(qn("a:solidFill"))
            if fill is not None:
                for child in list(labels[key].spPr):
                    if child.tag.rsplit("}", 1)[-1] in {"solidFill", "noFill", "gradFill", "blipFill", "pattFill", "grpFill"}:
                        labels[key].spPr.remove(child)
                labels[key].spPr.insert_element_before(deepcopy(fill), "a:ln", "a:effectLst", "a:effectDag", "a:scene3d", "a:sp3d")
                break
    label_width = min(Pt(40), max(Pt(28), min(
        (parts["label"].width for parts in sections.values() if parts["label"] is not None),
        default=Pt(32),
    )))
    for shape in list(slide.shapes):
        is_section_shape = any(shape._element is candidate._element for candidate in shapes)
        is_empty_frame = (
            shape.shape_type in {MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.TEXT_BOX}
            and not _text(getattr(shape, "text", ""))
            and shape.top >= top - Pt(12) and shape.left >= left - Pt(12)
            and shape.left + shape.width <= right + Pt(12)
            and shape.top + shape.height <= bottom + Pt(12)
        )
        if is_section_shape or is_empty_frame:
            _remove_element(shape._element)
    body_left = left + label_width + PPT_FLOW_GAP
    if right - body_left < Pt(100):
        raise ValueError("周报模板正文宽度不足以容纳单列内容")
    return {"left": left, "top": top, "bottom": bottom, "label_width": label_width,
            "body_left": body_left, "body_width": right - body_left}, labels


def _format_flow_paragraph(paragraph, line_height: int, color: RGBColor, font_size: float = PPT_FLOW_FONT_SIZE) -> None:
    """Write absolute spacing and explicit character metrics, preserving emphasis."""
    paragraph.line_spacing = line_height
    paragraph.space_before = paragraph.space_after = Pt(0)
    paragraph.font.size = Pt(font_size)
    properties = paragraph._p.get_or_add_pPr()
    if properties.get("marL") is None:
        properties.set("marL", str(int(properties.get("lvl", "0")) * Pt(font_size)))
    if properties.get("marR") is None:
        properties.set("marR", "0")
    if properties.get("indent") is None:
        properties.set("indent", "0")
    character_properties = [properties.find(qn("a:defRPr"))]
    end_properties = paragraph._p.find(qn("a:endParaRPr"))
    if end_properties is None:
        end_properties = OxmlElement("a:endParaRPr")
        paragraph._p.append(end_properties)
    character_properties.append(end_properties)
    for child in paragraph._p:
        if child.tag in {qn("a:r"), qn("a:fld"), qn("a:br")}:
            run_properties = child.find(qn("a:rPr"))
            if run_properties is None:
                run_properties = OxmlElement("a:rPr")
                child.insert(0, run_properties)
            character_properties.append(run_properties)
    for properties in character_properties:
        properties.set("sz", str(round(font_size * 100)))
        families = ("latin", "ea", "cs")
        for family_index, family in enumerate(families):
            font = properties.find(qn(f"a:{family}"))
            if font is None:
                font = OxmlElement(f"a:{family}")
                successors = tuple(f"a:{item}" for item in families[family_index + 1:])
                properties.insert_element_before(font, *successors, "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst")
            font.set("typeface", PPT_BODY_FONT)
        for child in list(properties):
            if child.tag.rsplit("}", 1)[-1] in {"solidFill", "noFill", "gradFill", "blipFill", "pattFill", "grpFill"}:
                properties.remove(child)
        fill = OxmlElement("a:solidFill")
        rgb = OxmlElement("a:srgbClr")
        rgb.set("val", str(color))
        fill.append(rgb)
        properties.insert_element_before(fill, "a:effectLst", "a:effectDag", "a:highlight", "a:uLnTx", "a:uLn",
                                         "a:uFillTx", "a:uFill", "a:latin", "a:ea", "a:cs", "a:sym",
                                         "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst")


def _flow_render_coordinates(fragment: dict[str, Any], geometry: dict[str, int]) -> dict[str, int]:
    coordinates = {name: fragment.get(name, geometry[name])
                   for name in ("left", "label_width", "body_left", "body_width")}
    region_left = coordinates["left"]
    region_right = coordinates["body_left"] + coordinates["body_width"]
    full_width_row = (
        region_left == geometry["left"]
        and region_right == geometry["body_left"] + geometry["body_width"]
    )
    if _section_is_issue(fragment["section"]) and not full_width_row:
        # Side columns use an outer-right title; a standalone full-width row
        # keeps the left title shared by the other vertically stacked rows.
        coordinates["left"] = region_right - coordinates["label_width"]
        coordinates["body_left"] = region_left
    return coordinates


def _render_flow_page(slide, fragments: list[dict[str, Any]], geometry: dict[str, int], labels: dict[str, Any]) -> None:
    for shape in list(slide.shapes):
        if shape.name.startswith("weekly-flow:"):
            _remove_element(shape._element)
    for fragment in fragments:
        key = fragment["section"]
        coordinates = _flow_render_coordinates(fragment, geometry)
        name = f"weekly-flow:{key}:{fragment['index']}"
        label_xml = labels.get(key)
        if label_xml is not None:
            element = deepcopy(label_xml)
            non_visual = element.find(qn("p:nvSpPr")).find(qn("p:cNvPr"))
            non_visual.set("id", str(slide.shapes._next_shape_id))
            slide.shapes._spTree.insert_element_before(element, "p:extLst")
            label = slide.shapes[-1]
        else:
            label = slide.shapes.add_textbox(0, 0, 1, 1)
        if _is_supplementary_key(key) or label.fill.type in {None, MSO_FILL_TYPE.BACKGROUND}:
            label.fill.solid()
            label.fill.fore_color.rgb = RGBColor(255, 255, 0) if _section_is_issue(key) else RGBColor(20, 68, 216)
        label.name = name + ":label"
        label.left, label.top = coordinates["left"], fragment["top"]
        label.width, label.height = coordinates["label_width"], fragment["height"]
        label.rotation = 0
        label.text = "\n".join(_section_title(key))
        label.text_frame._txBody.bodyPr.set("vert", "horz")
        label.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        label.text_frame.word_wrap = False
        label.text_frame.auto_size = MSO_AUTO_SIZE.NONE
        label.text_frame.margin_top = label.text_frame.margin_bottom = PPT_BODY_MARGIN_TOP
        label.text_frame.margin_left = label.text_frame.margin_right = 0
        for paragraph in label.text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            _clear_paragraph_bullets(paragraph)
            _format_flow_paragraph(paragraph, PPT_FLOW_LABEL_LINE_HEIGHT,
                                   PPT_ISSUE_COLOR if _section_is_issue(key) else RGBColor(255, 255, 255), PPT_FLOW_LABEL_FONT_SIZE)
        border = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,
                                        coordinates["body_left"], fragment["top"],
                                        coordinates["body_width"], fragment["height"])
        border.name = name + ":border"
        border.fill.background()
        border.line.color.rgb = RGBColor(49, 93, 250)
        border.line.width = Pt(1)
        body = slide.shapes.add_textbox(coordinates["body_left"], fragment["top"],
                                       coordinates["body_width"], fragment["height"])
        body.name = name + ":body"
        frame = body.text_frame
        frame.word_wrap = True
        frame.auto_size = MSO_AUTO_SIZE.NONE
        frame.vertical_anchor = MSO_ANCHOR.TOP
        frame.margin_left = frame.margin_right = PPT_BODY_MARGIN_LEFT
        frame.margin_top = frame.margin_bottom = PPT_BODY_MARGIN_TOP
        frame._txBody.bodyPr.set("vert", "horz")
        for element in list(frame._txBody.findall(qn("a:p"))):
            frame._txBody.remove(element)
        for item in fragment["paragraphs"]:
            element = deepcopy(item["xml"])
            if item.get("source_slide") is not None:
                _copy_relationships(item["source_slide"], slide, element)
            frame._txBody.append(element)
        for paragraph in frame.paragraphs:
            paragraph.alignment = PP_ALIGN.LEFT
            _format_flow_paragraph(paragraph, fragment["line_height"],
                                   PPT_ISSUE_COLOR if _section_is_issue(key) else RGBColor(0, 0, 0), fragment["font_size"])


def _text_shape_metrics(
    shape, font_size: float, line_height: int, bottom_safety: int = PPT_FLOW_BOTTOM_SAFETY,
) -> tuple[bool, int | None]:
    """Check written metrics; percentage spacing cannot prove a text box fits."""
    frame = shape.text_frame
    format_ok = frame.auto_size == MSO_AUTO_SIZE.NONE
    required = frame.margin_top + frame.margin_bottom + bottom_safety
    known_line_heights = True
    for paragraph in frame.paragraphs:
        spacing = paragraph.line_spacing
        format_ok = format_ok and (
            spacing == line_height
            and paragraph.font.size == Pt(font_size)
            and paragraph.space_before == paragraph.space_after == Pt(0)
        )
        properties = [paragraph._p.find(qn("a:endParaRPr"))]
        properties.extend(child.find(qn("a:rPr")) for child in paragraph._p
                          if child.tag in {qn("a:r"), qn("a:fld"), qn("a:br")})
        format_ok = format_ok and all(item is not None and item.get("sz") == str(round(font_size * 100))
                                      for item in properties)
        # python-pptx returns a Length (an integer) for fixed-point spacing,
        # and a float for multiples such as 1.3. Only the former is measurable.
        if not isinstance(spacing, int) or spacing <= 0:
            known_line_heights = False
            continue
        lines = len(_flow_line_ends({"text": paragraph.text, "xml": paragraph._p,
                                    "tab_size": frame._txBody.bodyPr.get("defTabSz", str(Pt(36)))}, shape.width,
                                    frame.margin_left, frame.margin_right, font_size))
        required += lines * spacing + (paragraph.space_before or 0) + (paragraph.space_after or 0)
    return format_ok, required if known_line_heights else None


def _flow_shape_metrics(
    shape, is_label: bool = False, font_size: float = PPT_FLOW_FONT_SIZE, line_height: int | None = None,
) -> tuple[bool, int | None]:
    frame = shape.text_frame
    side_margin = 0 if is_label else PPT_BODY_MARGIN_LEFT
    format_ok, required = _text_shape_metrics(
        shape, PPT_FLOW_LABEL_FONT_SIZE if is_label else font_size,
        PPT_FLOW_LABEL_LINE_HEIGHT if is_label else (line_height if line_height is not None else _ppt_body_line_height(font_size)),
        0 if is_label else PPT_FLOW_BOTTOM_SAFETY,
    )
    margins_ok = (frame.margin_left == frame.margin_right == side_margin
                  and frame.margin_top == frame.margin_bottom == PPT_BODY_MARGIN_TOP)
    return format_ok and margins_ok, required


def _flow_page_geometry_matches(shapes: dict[str, Any], fragments: list[dict[str, Any]], geometry: dict[str, int]) -> bool:
    rectangles = []
    right = geometry["body_left"] + geometry["body_width"]
    for fragment in fragments:
        coordinates = _flow_render_coordinates(fragment, geometry)
        prefix = f"weekly-flow:{fragment['section']}:{fragment['index']}"
        body, label, border = (shapes.get(prefix + suffix) for suffix in (":body", ":label", ":border"))
        if body is None or label is None or border is None:
            return False
        title_gap_matches = (
            label.left == body.left + body.width + PPT_FLOW_GAP
            if coordinates["left"] > coordinates["body_left"]
            else body.left == label.left + label.width + PPT_FLOW_GAP
        )
        region_left = min(label.left, body.left)
        region_right = max(label.left + label.width, body.left + body.width)
        if not (
            body.top == label.top == border.top == fragment["top"]
            and body.height == label.height == border.height == fragment["height"]
            and body.height >= _flow_label_height(fragment["section"])
            and body.left == border.left == coordinates["body_left"]
            and body.width == border.width == coordinates["body_width"]
            and label.left == coordinates["left"] and label.width == coordinates["label_width"]
            and title_gap_matches
            and region_left >= geometry["left"] and region_right <= right
            and body.top >= geometry["top"] and body.top + body.height <= geometry["bottom"]
        ):
            return False
        rectangle = (region_left, body.top, region_right, body.top + body.height)
        if any(rectangle[0] < other[2] and other[0] < rectangle[2]
               and rectangle[1] < other[3] and other[1] < rectangle[3] for other in rectangles):
            return False
        rectangles.append(rectangle)
    return bool(fragments)


def _estimated_chunk_count(shape, font_size: float = 11.0) -> int:
    """Estimate wrapping using the same conservative character model as pagination."""
    text = _text(shape.text)
    if not text:
        return 1
    chars_per_line, max_lines = _shape_capacity_for_font(shape, font_size)
    line_count = sum(
        max(1, (len(raw_line or " ") + chars_per_line - 1) // chars_per_line)
        for raw_line in re.split(r"\r?\n", text)
    )
    return max(1, (line_count + max_lines - 1) // max_lines)


def _adapt_overflowing_project_layout(slide) -> bool:
    """Let an overflowing current/next section own a full row before paginating."""
    entries = _flatten_shapes(slide.shapes)
    parts = {
        key: _section_layout_parts(slide, entries, key)
        for key in ("current", "next", "issues")
    }
    if not all(parts.values()):
        return False
    overflowing = [
        key for key, value in parts.items()
        if _estimated_chunk_count(value["body"]) > 1
    ]
    # The standard template already gives issues the complete right column.
    # Row expansion is useful when exactly one of the two left modules overflows.
    if len(overflowing) != 1 or overflowing[0] not in {"current", "next"}:
        return False

    dominant_key = overflowing[0]
    companion_key = "next" if dominant_key == "current" else "current"
    dominant = parts[dominant_key]
    companion = parts[companion_key]
    issues = parts["issues"]
    if dominant["border"] is None or companion["border"] is None:
        return False

    shapes = {
        dominant["border"], dominant["body"],
        issues["label"], issues["body"],
    }
    original_geometry = {
        shape: (shape.left, shape.top, shape.width, shape.height)
        for shape in shapes
    }
    issue_right = issues["label"].left + issues["label"].width
    width_delta = issue_right - (dominant["border"].left + dominant["border"].width)
    dominant["border"].width += width_delta
    dominant["body"].width += width_delta

    # Move the issue module into the non-dominant row. This keeps all three
    # sections on a normal page while giving the long section the full width.
    issues["label"].top = companion["border"].top
    issues["label"].height = companion["border"].height
    issues["body"].top = companion["body"].top
    issues["body"].height = companion["body"].height

    if all(_estimated_chunk_count(value["body"]) <= 1 for value in parts.values()):
        return True
    for shape, (left, top, width, height) in original_geometry.items():
        shape.left, shape.top, shape.width, shape.height = left, top, width, height
    return False


def _format_project_slide_bodies(slide) -> None:
    entries = _flatten_shapes(slide.shapes)
    for entry in entries:
        shape = entry["shape"]
        if entry["top"] < 900000 or not entry["text"] or not getattr(shape, "has_text_frame", False):
            continue
        if _entry_role(entry) != "body":
            continue
        key = _body_section_key(entries, entry)
        font_size = _project_body_font_size(
            shape, preserve_paragraphs=_is_supplementary_key(key) and not _section_is_issue(key),
        )
        color = PPT_ISSUE_COLOR if _section_is_issue(key) else None
        _set_project_body_font(shape, font_size, color)


def _shape_capacity_for_font(shape, font_size: float) -> tuple[int, int]:
    emu_per_point = 12700
    chars_per_line = max(8, int(int(shape.width) / (font_size * emu_per_point * 1.08)))
    max_lines = max(1, int(int(shape.height) / (font_size * emu_per_point * 1.3)))
    return chars_per_line, max_lines


def _text_chunks(shape) -> list[str]:
    text = _text(shape.text)
    if not text:
        return [""]
    chars_per_line, max_lines = _shape_capacity(shape)
    lines: list[tuple[str, bool]] = []
    # 仅把真正的段落换行（LF/CRLF）作为条目边界；不要使用 splitlines()，
    # 因为它会把 PowerPoint 的软换行（\v）也拆成新的段落。
    for raw_line in re.split(r"\r?\n", text) or [text]:
        line = raw_line or " "
        first_chunk = True
        while len(line) > chars_per_line:
            # 同一原始段落被切开的片段使用软换行，后续格式化时仍属于同一个
            # PowerPoint 段落，不会再次添加项目符号。
            lines.append((line[:chars_per_line], not first_chunk))
            line = line[chars_per_line:]
            first_chunk = False
        lines.append((line, not first_chunk))
    # 文本框高度已包含项目符号的实际布局；过度预留行数会把本可容纳的正文
    # 错误判定为溢出，导致字号被无谓缩小。
    chunk_size = max(1, max_lines)
    chunks = []
    for index in range(0, len(lines), chunk_size):
        chunk_lines = lines[index:index + chunk_size]
        rendered = chunk_lines[0][0] if chunk_lines else ""
        if chunk_lines and chunk_lines[0][1]:
            rendered = CONTINUATION_MARKER + rendered
        for value, continuation in chunk_lines[1:]:
            rendered += ("\v" if continuation else "\n") + value
        chunks.append(rendered)
    return chunks


def _overflow_chunks(slide) -> dict[str, list[str]]:
    chunks = {}
    entries = _flatten_shapes(slide.shapes)
    for entry in entries:
        if entry["top"] < 900000 or not entry["text"] or _entry_role(entry) != "body":
            continue
        key = _body_section_key(entries, entry)
        font_size = _project_body_font_size(
            entry["shape"], preserve_paragraphs=_is_supplementary_key(key) and not _section_is_issue(key),
        )
        color = PPT_ISSUE_COLOR if _section_is_issue(key) else None
        _set_project_body_font(entry["shape"], font_size, color)
        values = _text_chunks(entry["shape"])
        chunks[entry["path"]] = values
    return chunks


def _section_pagination(slide, chunks: dict[str, list[str]]) -> dict[str, Any]:
    """Build ordered section starts while allowing adjacent sections to share a boundary page."""
    entries = _flatten_shapes(slide.shapes)
    section_by_path = {}
    label_by_path = {}
    durations = {key: 0 for key in ("current", "next")}
    for entry in entries:
        key = _entry_section_key(entry)
        if key:
            label_by_path[entry["path"]] = key
            durations.setdefault(key, 0)
        if entry["path"] not in chunks:
            continue
        key = _body_section_key(entries, entry)
        if key:
            section_by_path[entry["path"]] = key
            durations[key] = max(durations.get(key, 0), len(chunks[entry["path"]]))

    starts = {}
    cursor = 0
    for key in durations:
        if not durations[key]:
            continue
        starts[key] = cursor
        cursor += durations[key] - 1
    page_count = max(
        [1]
        + [starts[key] + duration for key, duration in durations.items() if duration]
        + [len(values) for path, values in chunks.items() if path not in section_by_path]
    )
    return {
        "section_by_path": section_by_path,
        "label_by_path": label_by_path,
        "starts": starts,
        "page_count": page_count,
    }


def _body_texts(slide) -> dict[str, str]:
    """Capture complete body text before overflow pages replace individual chunks."""
    return {
        entry["path"]: entry["text"]
        for entry in _flatten_shapes(slide.shapes)
        if entry["top"] >= 900000 and entry["text"] and _entry_role(entry) == "body"
    }


def _set_page_chunks(
    slide,
    chunks: dict[str, list[str]],
    page_index: int,
    pagination: dict[str, Any] | None = None,
) -> None:
    entries = _flatten_shapes(slide.shapes)
    active_sections = set()
    section_by_path = pagination.get("section_by_path", {}) if pagination else {}
    starts = pagination.get("starts", {}) if pagination else {}
    for entry in entries:
        values = chunks.get(entry["path"])
        if values is None:
            continue
        section_key = section_by_path.get(entry["path"])
        local_index = page_index - starts.get(section_key, 0)
        value = values[local_index] if 0 <= local_index < len(values) else ""
        entry["shape"].text = value
        if value and section_key:
            active_sections.add(section_key)
    if pagination:
        for entry in entries:
            section_key = pagination.get("label_by_path", {}).get(entry["path"])
            if section_key and section_key not in active_sections:
                entry["shape"].text = ""
    _format_project_slide_bodies(slide)


def _normalized_content(value: str) -> str:
    return re.sub(r"\s+", "", _text(value))


def _xml_signature(element, ignored_attributes: set[str] | None = None):
    if element is None:
        return None
    ignored = ignored_attributes or set()
    attributes = tuple(sorted(
        (name, value) for name, value in element.attrib.items()
        if name.rsplit("}", 1)[-1] not in ignored
    ))
    return element.tag, attributes, tuple(_xml_signature(child, ignored) for child in element)


def _text_style_matches(shape, template_shape) -> bool:
    template_paragraphs = template_shape.text_frame.paragraphs
    if not template_paragraphs:
        return True
    for paragraph_index, paragraph in enumerate(shape.text_frame.paragraphs):
        template_paragraph = template_paragraphs[min(paragraph_index, len(template_paragraphs) - 1)]
        if _xml_signature(paragraph._p.find(qn("a:pPr"))) != _xml_signature(template_paragraph._p.find(qn("a:pPr"))):
            return False
        template_run = next(iter(template_paragraph.runs), None)
        template_rpr = template_run._r.find(qn("a:rPr")) if template_run is not None else None
        template_signature = _xml_signature(template_rpr, {"sz"})
        for run in paragraph.runs:
            if _xml_signature(run._r.find(qn("a:rPr")), {"sz"}) != template_signature:
                return False
    return True


def _matching_template_entry(source_entry: dict[str, Any], template_entries: list[dict[str, Any]]):
    role = _entry_role(source_entry)
    candidates = [entry for entry in template_entries if _entry_role(entry) == role]
    if not candidates and _is_supplementary_key(_entry_section_key(source_entry)):
        fallback = "issues" if _section_is_issue(role) else "current"
        candidates = [entry for entry in template_entries if _entry_section_key(entry) == fallback]
    if not candidates:
        candidates = [entry for entry in template_entries if _entry_role(entry) == "body"]
    if not candidates:
        return None
    return min(
        candidates,
        key=lambda entry: abs(entry["left"] - source_entry["left"]) + abs(entry["top"] - source_entry["top"]),
    )


def _generation_qa(
    presentation: Presentation,
    generated_groups: list[dict[str, Any]],
    max_rounds: int = 4,
    progress_callback: ProgressCallback = None,
) -> dict[str, Any]:
    """多轮核验内容、模板文字样式和版面稳定性，并修复可确定的问题。"""
    history = []
    stable_rounds = 0
    final_issues = []
    thresholds = {"content": 0.99, "style": 0.97, "layout": 0.97}
    for round_number in range(1, max_rounds + 1):
        _notify(progress_callback, "生成质量核验", 91 + round_number, f"正在执行第 {round_number} / {max_rounds} 轮内容、样式与版面核验")
        issues = []
        repairs = 0
        content_total = content_ok = 0
        style_total = style_ok = 0
        layout_total = layout_ok = 0
        for group in generated_groups:
            for slide in group["slides"]:
                content_total += 1
                title, reporter, _ = _slide_title(_flatten_shapes(slide.shapes))
                if group.get("preserve_source_title"):
                    expected_title, expected_reporter, _ = _slide_title(_flatten_shapes(group["source_slide"].shapes))
                    title_ok = _normalize_title(title) == _normalize_title(expected_title)
                    reporter_ok = _normalize_title(reporter) == _normalize_title(expected_reporter)
                else:
                    title_ok = _canonical_title(title) == group["project"]["key"]
                    reporter_ok = _normalize_title(reporter) == _normalize_title(group["project"]["reporter"])
                if title_ok and reporter_ok:
                    content_ok += 1
                else:
                    repairs += 1
                if group.get("preserve_source_title"):
                    _replace_title(slide, expected_title, expected_reporter, group.get("title_reference"))
                else:
                    _replace_title(
                        slide,
                        group["project"]["title"],
                        group["project"]["reporter"],
                        group.get("title_reference"),
                    )
            if "compact_fit" in group:
                for slide in group["slides"]:
                    entries = _flatten_shapes(slide.shapes)
                    shapes = {str(entry["shape"].shape_id): entry["shape"] for entry in entries}
                    for section in _supplementary_sections(group["source"]):
                        key = _supplementary_key(section["title"], section["kind"])
                        if section["title"] == "待确认内容":
                            continue
                        content_total += 1
                        actual = _section_from_entries(entries, key)
                        if _flow_content_key(_normalized_section_content(actual)) == _flow_content_key(section["content"]):
                            content_ok += 1
                        else:
                            issues.append({"severity": "error", "code": "generated_content_mismatch",
                                           "label": "生成内容不一致", "file": group["source"]["file"],
                                           "slide": group["source"]["slide"], "location": section["title"],
                                           "project": group["project"]["title"],
                                           "detail": "补充栏目标题或正文与源文件识别结果不一致。",
                                           "suggestion": "请确认该栏目的实际标题、正文和源页归属。"})
                    for shape_id, fitted in group["compact_fit"].items():
                        shape = shapes.get(shape_id)
                        content_total += 1
                        content_matches = shape is not None and _flow_content_key(shape.text) == _flow_content_key(fitted["text"])
                        content_ok += int(content_matches)
                        style_total += 1
                        style_matches = shape is not None and _xml_signature(shape._element) == _xml_signature(fitted["xml"])
                        style_ok += int(style_matches)
                        layout_total += 1
                        format_ok, required = (_text_shape_metrics(shape, fitted["font_size"], fitted["line_height"])
                                               if shape is not None else (False, None))
                        fits = shape is not None and format_ok and required is not None and required <= shape.height
                        layout_ok += int(fits)
                        if not content_matches or not style_matches or not fits:
                            issues.append({
                                "severity": "error" if not content_matches else "warning",
                                "code": "generated_content_mismatch" if not content_matches else "generated_text_metrics_mismatch",
                                "label": "生成内容不一致" if not content_matches else "生成文字尺寸异常",
                                "file": group["source"]["file"], "slide": group["source"]["slide"],
                                "location": f"正文对象 {shape_id}", "project": group["project"]["title"],
                                "detail": "原布局正文与缩字试排结果不一致，已恢复选定字号、行距和内容。",
                                "suggestion": "请确认最终生成核验结果。",
                            })
                            if shape is None:
                                slide.shapes._spTree.insert_element_before(deepcopy(fitted["xml"]), "p:extLst")
                            else:
                                _copy_xml_contents(shape._element, fitted["xml"])
                            repairs += 1
                    for entry in entries:
                        layout_total += 1
                        if (entry["left"] >= 0 and entry["top"] >= 0
                                and entry["left"] + entry["width"] <= presentation.slide_width
                                and entry["top"] + entry["height"] <= presentation.slide_height):
                            layout_ok += 1
                # The selected size is final: legacy formatting/QA must not run
                # another font-selection pass or replace it with template text.
                continue
            if group.get("flow_pages"):
                page_shapes = [
                    {shape.name: shape for shape in slide.shapes if shape.name.startswith("weekly-flow:")}
                    for slide in group["slides"]
                ]
                repair_pages = set()
                for key in group["flow_sections"]:
                    content_total += 1
                    actual = "".join(
                        _flow_content_key(shape.text)
                        for shapes in page_shapes for name, shape in shapes.items()
                        if name.startswith(f"weekly-flow:{key}:") and name.endswith(":body")
                    )
                    if actual == _flow_content_key(group["flow_sections"][key]):
                        content_ok += 1
                    else:
                        repair_pages.update(range(len(group["slides"])))
                        issues.append({
                            "severity": "error", "code": "generated_content_mismatch", "label": "生成内容不一致",
                            "file": group["source"]["file"], "slide": group["source"]["slide"],
                            "location": _section_title(key), "project": group["project"]["title"],
                            "detail": "单列分页正文与项目汇总内容不一致，可能存在重复或缺失。",
                            "suggestion": "系统已按栏目分页结果重新写入；请确认最终审核结果。",
                        })
                for page_index, (shapes, references) in enumerate(zip(page_shapes, group["flow_references"])):
                    fragments = {
                        f"weekly-flow:{item['section']}:{item['index']}": item
                        for item in group["flow_pages"][page_index]
                    }
                    content_total += 1
                    if list(shapes) == list(references):
                        content_ok += 1
                    else:
                        repair_pages.add(page_index)
                    for name, reference in references.items():
                        style_total += 1
                        shape = shapes.get(name)
                        if shape is not None and _xml_signature(shape._element, {"id"}) == _xml_signature(reference, {"id"}):
                            style_ok += 1
                        else:
                            repair_pages.add(page_index)
                    layout_total += 1
                    if _flow_page_geometry_matches(shapes, group["flow_pages"][page_index], group["flow_geometry"]):
                        layout_ok += 1
                    else:
                        repair_pages.add(page_index)
                        issues.append({
                            "severity": "warning", "code": "generated_layout_mismatch", "label": "生成栏目布局不一致",
                            "file": group["source"]["file"], "slide": group["source"]["slide"],
                            "location": f"单列第 {page_index + 1} 页", "project": group["project"]["title"],
                            "detail": "正文框、标签或边框未对齐，或栏目间距、正文区底部位置不符。",
                            "suggestion": "系统已按分页结果恢复栏目尺寸，请确认最终审核结果。",
                        })
                    for entry in _flatten_shapes(group["slides"][page_index].shapes):
                        layout_total += 1
                        if (entry["left"] >= 0 and entry["top"] >= 0
                                and entry["left"] + entry["width"] <= presentation.slide_width
                                and entry["top"] + entry["height"] <= presentation.slide_height):
                            layout_ok += 1
                        if (entry["shape"].name.startswith("weekly-flow:")
                                and entry["shape"].name.endswith((":body", ":label"))):
                            layout_total += 1
                            fragment = fragments.get(entry["shape"].name.rsplit(":", 1)[0])
                            format_ok, required = (_flow_shape_metrics(
                                entry["shape"], entry["shape"].name.endswith(":label"),
                                fragment["font_size"], fragment["line_height"],
                            ) if fragment is not None else (False, None))
                            if format_ok and required is not None and required <= entry["height"]:
                                layout_ok += 1
                            else:
                                repair_pages.add(page_index)
                                issues.append({
                                    "severity": "warning",
                                    "code": "generated_text_overflow" if format_ok else "generated_text_metrics_mismatch",
                                    "label": "生成文本仍可能溢出" if format_ok else "生成文字尺寸异常",
                                    "file": group["source"]["file"], "slide": group["source"]["slide"],
                                    "location": entry["shape"].name, "project": group["project"]["title"],
                                    "detail": ("文字超出包含内边距和底部安全余量的可用高度。" if format_ok
                                               else "实际字号、固定行距或内边距与单列排版设置不一致。"),
                                    "suggestion": "系统已按固定行距重新写入，请确认最终审核结果。",
                                })
                for page_index in repair_pages:
                    _render_flow_page(group["slides"][page_index], group["flow_pages"][page_index],
                                      group["flow_geometry"], group["flow_labels"])
                    repairs += 1
                continue
            source_entries = group.get("source_body_texts") or _body_texts(group["source_slide"])
            page_maps = [
                {entry["path"]: entry for entry in _flatten_shapes(slide.shapes)}
                for slide in group["slides"]
            ]
            for path, source_text in source_entries.items():
                content_total += 1
                expected = _normalized_content(source_text)
                actual = _normalized_content("".join(
                    page[path]["text"] for page in page_maps
                    if path in page and page[path]["text"]
                ))
                if actual == expected:
                    content_ok += 1
                else:
                    issues.append({
                        "severity": "error", "code": "generated_content_mismatch", "label": "生成内容不一致",
                        "file": group["source"]["file"], "slide": group["source"]["slide"],
                        "location": f"对象 {path}", "project": group["project"]["title"],
                        "detail": "生成页中的正文与源 PPT 不一致，可能存在重复或缺失。",
                        "suggestion": "系统已重新写入分页内容；请在最终审核结果中确认。",
                    })
                    for page_index, slide in enumerate(group["slides"]):
                        _set_page_chunks(slide, group["chunks"], page_index, group.get("pagination"))
                    repairs += 1

            style_reference = group.get("style_reference") or group.get("template_slide")
            reference_entries = [
                entry for entry in _flatten_shapes(style_reference.shapes)
                if getattr(entry["shape"], "has_text_frame", False)
            ] if style_reference is not None else []
            for slide in group["slides"]:
                for entry in _flatten_shapes(slide.shapes):
                    shape = entry["shape"]
                    if entry["text"] and getattr(shape, "has_text_frame", False):
                        layout_total += 1
                        if len(_text_chunks(shape)) <= 1:
                            layout_ok += 1
                        else:
                            issues.append({
                                "severity": "warning", "code": "generated_text_overflow", "label": "生成文本仍可能溢出",
                                "file": group["source"]["file"], "slide": group["source"]["slide"],
                                "location": f"对象 {entry['path']}", "project": group["project"]["title"],
                                "detail": "自动缩放和分页后，文本量仍超过当前文本框估算容量。",
                                "suggestion": "建议人工检查该页文本框实际显示效果。",
                            })
                        reference_entry = _matching_template_entry(entry, reference_entries) if reference_entries else None
                        if reference_entry is not None:
                            style_total += 1
                            if _text_style_matches(shape, reference_entry["shape"]):
                                style_ok += 1
                            else:
                                _copy_text_style(shape, reference_entry["shape"])
                                repairs += 1
                        _format_project_slide_bodies(slide)
                    layout_total += 1
                    if entry["left"] >= 0 and entry["top"] >= 0 and entry["left"] + entry["width"] <= presentation.slide_width and entry["top"] + entry["height"] <= presentation.slide_height:
                        layout_ok += 1

        scores = {
            "content": content_ok / content_total if content_total else 1.0,
            "style": style_ok / style_total if style_total else 1.0,
            "layout": layout_ok / layout_total if layout_total else 1.0,
        }
        passed = all(scores[key] >= value for key, value in thresholds.items())
        stable_rounds = stable_rounds + 1 if passed and repairs == 0 else 0
        history.append({
            "round": round_number,
            "content_score": round(scores["content"] * 100, 1),
            "style_score": round(scores["style"] * 100, 1),
            "layout_score": round(scores["layout"] * 100, 1),
            "repairs": repairs,
            "issues": len(issues),
            "stable": stable_rounds >= 2,
        })
        final_issues = issues
        if stable_rounds >= 2:
            break
    last = history[-1]
    overall_score = round(min(last["content_score"], last["style_score"], last["layout_score"]), 1)
    return {
        "status": "稳定" if stable_rounds >= 2 else "需人工复核",
        "stable": stable_rounds >= 2,
        "score": overall_score,
        "thresholds": {key: round(value * 100) for key, value in thresholds.items()},
        "rounds": history,
        "issues": final_issues,
    }


def _iter_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_shapes(shape.shapes)


def _project_title_shape(slide):
    candidates = [
        shape for shape in _iter_shapes(slide.shapes)
        if getattr(shape, "has_text_frame", False) and _text(shape.text)
    ]
    return next(
        (shape for shape in candidates if PPT_HEADING.match(_text(shape.text))),
        candidates[0] if candidates else None,
    )


def _replace_title(slide, title: str, reporter: str, reference_shape=None) -> None:
    candidates = [shape for shape in _iter_shapes(slide.shapes) if getattr(shape, "has_text_frame", False) and _text(shape.text)]
    title_shape = next((shape for shape in candidates if PPT_HEADING.match(_text(shape.text))), candidates[0] if candidates else None)
    if not title_shape:
        return
    if reference_shape is not None:
        title_shape.left = reference_shape.left
        title_shape.top = reference_shape.top
        title_shape.width = reference_shape.width
        title_shape.height = reference_shape.height
        _copy_text_style(title_shape, reference_shape)
    paragraph = title_shape.text_frame.paragraphs[0]
    _clear_paragraph_runs(paragraph, _clean_line(title))
    title_run = paragraph.runs[0]
    title_run.font.size = Pt(24)
    _set_ppt_run_font(title_run, PPT_REPORTER_FONT, bold=True)
    paragraph.alignment = PP_ALIGN.LEFT
    _clear_paragraph_bullets(paragraph)
    reporter_run = paragraph.add_run()
    reporter_run.text = f"（汇报人：{reporter or '待补充'}）"
    reporter_run.font.size = Pt(18)
    _set_ppt_run_font(reporter_run, PPT_REPORTER_FONT, bold=False)
    for extra in title_shape.text_frame.paragraphs[1:]:
        _clear_paragraph_runs(extra, "")


def _finalize_generated_titles(generated_groups: list[dict[str, Any]]) -> None:
    """Restore reporter styling after QA has copied template text styles."""
    for group in generated_groups:
        if group.get("preserve_source_title"):
            title, reporter, _ = _slide_title(_flatten_shapes(group["source_slide"].shapes))
        else:
            title = group["project"]["title"]
            reporter = group["project"]["reporter"]
        for slide in group["slides"]:
            _replace_title(slide, title, reporter, group.get("title_reference"))


def _fill_missing_project_sections(slide) -> None:
    """在缺少源文件时保留项目模板页，并在各内容区域填入“无”。"""
    entries = _flatten_shapes(slide.shapes)
    for section_key in SECTION_DISPLAY:
        labels = [entry for entry in entries if _is_section_label(entry["text"], section_key)]
        if not labels:
            continue
        label = min(labels, key=lambda entry: (entry["top"], entry["left"]))
        next_tops = [
            entry["top"] for entry in entries
            if entry is not label and entry["top"] > label["top"] + 1000 and _is_section_label(entry["text"])
        ]
        bottom = min(next_tops) if next_tops else 10**10
        bodies = [
            entry for entry in entries
            if entry is not label
            and getattr(entry["shape"], "has_text_frame", False)
            and entry["left"] >= label["left"] + label["width"]
            and entry["top"] >= label["top"] - 100000
            and entry["top"] < bottom
            and not _is_section_label(entry["text"])
            and not PPT_HEADING.match(entry["text"])
        ]
        if not bodies:
            continue
        body = min(bodies, key=lambda entry: (abs(entry["top"] - label["top"]), entry["left"]))["shape"]
        body.text_frame.margin_left = 0
        body.text_frame.vertical_anchor = MSO_ANCHOR.TOP
        paragraph = body.text_frame.paragraphs[0]
        _clear_paragraph_runs(paragraph, "无")
        paragraph.alignment = PP_ALIGN.LEFT
        _clear_paragraph_bullets(paragraph)
        for extra in body.text_frame.paragraphs[1:]:
            _clear_paragraph_runs(extra, "")
        _format_project_body(body)


def _section_body_entries(entries: list[dict[str, Any]], key: str) -> list[dict[str, Any]]:
    """Find editable body boxes that belong to one explicitly labeled section."""
    labels = [entry for entry in entries if _entry_section_key(entry) == key]
    if not labels:
        return []
    label = min(labels, key=lambda entry: (entry["top"], entry["left"]))
    next_tops = [
        entry["top"] for entry in entries
        if entry is not label
        and entry["top"] > label["top"] + 1000
        and _entry_section_key(entry)
        and (entry["left"] >= 6000000) == (label["left"] >= 6000000)
    ]
    bottom = min(next_tops) if next_tops else 10**10
    bodies = []
    for entry in entries:
        shape = entry["shape"]
        if (
            entry is label
            or not getattr(shape, "has_text_frame", False)
            or _entry_role(entry) != "body"
            or not (label["top"] - 100000 <= entry["top"] < bottom)
        ):
            continue
        if _is_supplementary_key(key):
            is_in_section = _body_section_key(entries, entry) == key
        else:
            is_in_section = entry["left"] + 100000 >= label["left"] and entry["left"] <= 7200000
        if is_in_section:
            bodies.append(entry)
    # Some source pages contain an empty template placeholder alongside the
    # actual body text box. Prefer the populated box so its intended layout is
    # retained; use the largest empty box only when there is no source content.
    return sorted(
        bodies,
        key=lambda entry: (
            not bool(_text(entry["text"])),
            -(entry["width"] * entry["height"]),
            entry["top"],
            entry["left"],
        ),
    )


def _ensure_section_borders(slide) -> None:
    """Copy the standard blue content border when a source section omitted it."""
    for key in ("current", "next"):
        entries = _flatten_shapes(slide.shapes)
        labels = [entry for entry in entries if _is_section_label(entry["text"], key)]
        bodies = _section_body_entries(entries, key)
        text_body = next((entry for entry in bodies if entry["text"]), None)
        label = min(labels, key=lambda entry: (entry["top"], entry["left"])) if labels else None
        has_border = any(
            not entry["text"]
            and "." not in entry["path"]
            and entry["shape"].shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
            and label is not None
            and abs(entry["top"] - label["top"]) <= 150000
            and abs(entry["left"] - text_body["left"]) <= 150000
            and abs(entry["height"] - label["height"]) <= 150000
            for entry in bodies
        ) if text_body is not None else False
        if not labels or text_body is None or has_border:
            continue
        border_templates = [
            entry for entry in entries
            if not entry["text"]
            and "." not in entry["path"]
            and entry["shape"].shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
            and entry["width"] >= 2000000
        ]
        if not border_templates:
            continue
        template = max(border_templates, key=lambda entry: entry["width"] * entry["height"])["shape"]
        element = deepcopy(template._element)
        xfrm = element.spPr.xfrm
        xfrm.off.set("x", str(text_body["left"]))
        xfrm.off.set("y", str(label["top"]))
        xfrm.ext.set("cx", str(text_body["width"]))
        xfrm.ext.set("cy", str(label["height"]))
        non_visual = element.find(qn("p:nvSpPr")).find(qn("p:cNvPr"))
        non_visual.set("id", str(slide.shapes._next_shape_id))
        non_visual.set("name", f"{SECTION_DISPLAY[key].rstrip('：')}边框")
        text_body["shape"]._element.addprevious(element)


def _normalize_slide_section_labels(slide) -> None:
    """Rewrite label furniture only, retaining source body headings and label styling."""
    for entry in _flatten_shapes(slide.shapes):
        key = _entry_section_key(entry)
        if not key:
            continue
        shape = entry["shape"]
        inline = _content_after_section_label(entry["text"], key, entry.get("section_label"))
        if inline:
            # Inline label/body text is rebuilt by the flow renderer, where
            # the title and its paragraphs have independent styles and boxes.
            continue
        value = _section_title(key)
        paragraphs = list(shape.text_frame.paragraphs)
        _clear_paragraph_runs(paragraphs[0], value)
        for paragraph in paragraphs[1:]:
            _remove_element(paragraph._p)
        shape.name = "weekly-section:" + key


def _write_project_sections(slide, section_values: dict[str, Any], *, format_bodies: bool = True) -> None:
    """Write the audited section values back into a cloned source project slide."""
    entries = _flatten_shapes(slide.shapes)
    # The issue column in the supplied template is a grouped shape whose body
    # overlaps the main column geometrically. Its content is preserved here;
    # only the independently laid-out current/next-week sections are rebuilt.
    for key in ("current", "next"):
        bodies = _section_body_entries(entries, key)
        if not bodies:
            continue
        populated = [entry for entry in bodies if _text(entry["text"])]
        if (len(populated) == 1
                and _clean_content(populated[0]["text"]) == _clean_content(section_values.get(key, ""))
                and _flow_content_key(populated[0]["text"]) == _flow_content_key(section_values.get(key, ""))):
            # Keep matching rich paragraphs, including soft breaks in a single
            # paragraph. The second comparison also catches removed duplicates.
            continue
        bodies[0]["shape"].text = _normalized_section_content(section_values.get(key, ""))
        for entry in bodies[1:]:
            entry["shape"].text = ""
    if format_bodies:
        _format_project_slide_bodies(slide)
    for section in _supplementary_sections(section_values):
        if not _is_empty_section_content(section["content"]):
            continue
        bodies = _section_body_entries(entries, _supplementary_key(section["title"], section["kind"]))
        if bodies and all(_is_empty_section_content(entry["text"]) for entry in bodies):
            bodies[0]["shape"].text = "无"
            for entry in bodies[1:]:
                entry["shape"].text = ""
    _normalize_slide_section_labels(slide)
    _apply_issue_text_color(slide, section_values.get("issues", ""))


def _normalize_source_project_layout(slide) -> None:
    """统一源项目页正文的左上对齐和项目符号，避免沿用错误的居中属性。"""
    for entry in _flatten_shapes(slide.shapes):
        shape = entry["shape"]
        if not entry["text"] or not getattr(shape, "has_text_frame", False):
            continue
        if _has_paragraph_hierarchy(shape):
            continue
        if (
            entry["top"] < 900000
            or PPT_HEADING.match(entry["text"])
            or _entry_section_key(entry)
        ):
            continue
        shape.text_frame.margin_left = 0
        shape.text_frame.vertical_anchor = MSO_ANCHOR.TOP
        for paragraph in shape.text_frame.paragraphs:
            if not paragraph.text.strip():
                continue
            paragraph.alignment = PP_ALIGN.LEFT
            _clear_paragraph_bullets(paragraph)
        _format_project_body(shape)


def _expand_template_groups(slide, sections_only: bool = False) -> None:
    """将模板组合展开为页面坐标，使文本识别和容量计算使用真实尺寸。"""

    def is_content_group(shape) -> bool:
        entries = _flatten_shapes(shape.shapes)
        if any(_entry_section_key(entry) for entry in entries):
            return True
        # Some template pages group only the blank body frames, with the label
        # outside the group. They still need slide-space coordinates for sizing.
        return shape.top >= 900000 and any(
            getattr(entry["shape"], "has_text_frame", False)
            and _entry_role(entry) == "body"
            for entry in entries
        )

    while True:
        group = next((shape for shape in slide.shapes
                      if shape.shape_type == MSO_SHAPE_TYPE.GROUP
                      and (not sections_only or is_content_group(shape))), None)
        if group is None:
            return
        transform = group._element.grpSpPr.xfrm
        if transform is None or transform.chOff is None or transform.chExt is None:
            raise ValueError("周报模板组合形状缺少坐标变换信息")
        if not transform.chExt.cx or not transform.chExt.cy:
            raise ValueError("周报模板组合形状的内部尺寸为零")
        # 旋转或翻转的组合不能只通过平移、缩放展开，避免静默改变版式。
        if (transform.get("rot", "0") != "0"
                or transform.get("flipH", "0") in {"1", "true"}
                or transform.get("flipV", "0") in {"1", "true"}):
            raise ValueError("Word 周报模板的组合形状不支持旋转或翻转，请先取消组合")
        scale_x = group.width / transform.chExt.cx
        scale_y = group.height / transform.chExt.cy
        origin_x, origin_y = group.left, group.top
        child_x, child_y = transform.chOff.x, transform.chOff.y
        for shape in list(group.shapes):
            left = round(origin_x + (shape.left - child_x) * scale_x)
            top = round(origin_y + (shape.top - child_y) * scale_y)
            width, height = round(shape.width * scale_x), round(shape.height * scale_y)
            shape.left, shape.top = left, top
            shape.width, shape.height = width, height
            # 在原组合的位置依次移入，保留对象前后顺序；嵌套组合在下一轮展开。
            shape._element.getparent().remove(shape._element)
            group._element.addprevious(shape._element)
            non_visual = shape._element.xpath("./*/p:cNvPr")
            if non_visual:
                non_visual[0].set("id", str(slide.shapes._next_shape_id))
        _remove_element(group._element)


def build_weekly_presentation(
    result: dict[str, Any],
    source_lookup: dict[str, Path],
    target: str | Path,
    template: str | Path = PPT_TEMPLATE,
    progress_callback: ProgressCallback = None,
) -> None:
    """保留模板封面和结束页，短项目沿用原页，长项目按栏目单列续页。"""
    presentation = Presentation(template)
    week_end = datetime.strptime(result["week_end"], "%Y-%m-%d").date()
    cover_text = week_end.strftime("%Y年%m月%d日")
    for shape in presentation.slides[0].shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        compact_text = re.sub(r"\s+", "", _text(shape.text)).lower()
        if compact_text == COVER_WEBSITE:
            for paragraph in shape.text_frame.paragraphs:
                _clear_paragraph_runs(paragraph, "")
            continue
        if DATE_PATTERN.search(_text(shape.text)):
            paragraph = shape.text_frame.paragraphs[0]
            cover_value = DATE_PATTERN.sub(cover_text, _text(shape.text))
            cover_value = re.sub(
                rf"(?:\s*{re.escape(COVER_WEBSITE)}){{2,}}",
                f" {COVER_WEBSITE}",
                cover_value,
            )
            _clear_paragraph_runs(paragraph, cover_value.strip())

    placeholder_numbers = sorted(_project_slide_numbers(presentation), reverse=True)
    placeholder_slides = [presentation.slides[slide_number - 1] for slide_number in placeholder_numbers]
    outro_index = len(presentation.slides) - 1
    outro_slides = [slide for slide in presentation.slides if _is_outro_slide(slide)]
    # Keep the template's final outro page as the insertion anchor. Remove only
    # earlier duplicate outro pages so the generated deck always ends with one.
    for slide in reversed(outro_slides[:-1]):
        _remove_slide(presentation, slide)
    source_cache: dict[str, Presentation] = {}
    style_presentation = Presentation(template)
    style_slides = {
        project["key"]: style_presentation.slides[project["template_slide"] - 1]
        for project in result["projects"]
        if project.get("template_slide")
    }
    default_template_number = next(iter(_project_slide_numbers(style_presentation)), None)
    default_project_template = (
        style_presentation.slides[default_template_number - 1] if default_template_number else None
    )
    default_title_reference = next(
        (_project_title_shape(slide) for slide in style_slides.values() if _project_title_shape(slide)),
        None,
    )
    generated_groups = []

    def place_before_outro(slide) -> None:
        slide_ids = presentation.slides._sldIdLst
        # 新克隆页会先追加到末尾；将它移动到当前唯一结尾页之前。
        slide_id = next(item for item in slide_ids if item.id == slide.slide_id)
        slide_ids.remove(slide_id)
        outro_slide = next(
            (candidate for candidate in presentation.slides if _is_outro_slide(candidate)),
            None,
        )
        if outro_slide is None:
            slide_ids.append(slide_id)
            return
        outro_id = next(item for item in slide_ids if item.id == outro_slide.slide_id)
        slide_ids.insert(slide_ids.index(outro_id), slide_id)

    def place_after(reference_slide, slide) -> None:
        """Keep an overflow continuation adjacent to its source page."""
        slide_ids = presentation.slides._sldIdLst
        slide_id = next(item for item in slide_ids if item.id == slide.slide_id)
        slide_ids.remove(slide_id)
        reference_id = next(item for item in slide_ids if item.id == reference_slide.slide_id)
        slide_ids.insert(slide_ids.index(reference_id) + 1, slide_id)

    def append_flow_project(project, source_slides, flow_template, title_reference) -> None:
        canvas = _clone_source_slide(presentation, flow_template)
        _expand_template_groups(canvas, sections_only=True)
        sections = _flow_paragraphs(project, source_slides)
        geometry, labels = _prepare_flow_canvas(canvas, presentation)
        pages = _flow_page_plan(sections, geometry)
        slides = [canvas] + [_clone_source_slide(presentation, canvas) for _ in pages[1:]]
        references = []
        for slide, fragments in zip(slides, pages):
            _replace_title(slide, project["title"], project["reporter"], title_reference)
            _render_flow_page(slide, fragments, geometry, labels)
            references.append({shape.name: deepcopy(shape._element) for shape in slide.shapes
                               if shape.name.startswith("weekly-flow:")})
            place_before_outro(slide)
        generated_groups.append({
            "project": project, "source": project["slides"][0] if project["slides"] else {"file": "", "slide": 0},
            "source_slide": canvas, "title_reference": title_reference, "slides": slides,
            "flow_pages": pages, "flow_geometry": geometry, "flow_labels": labels,
            "flow_references": references,
            "flow_sections": {key: section["content"] for key, section in _output_sections(project).items()},
        })

    for project in result["projects"]:
        if _is_meeting_only_project(project):
            continue
        template_slide = style_slides.get(project["key"])
        title_reference = _project_title_shape(template_slide) if template_slide is not None else default_title_reference
        if not project["slides"]:
            if template_slide is None:
                continue
            append_flow_project(project, [], template_slide, title_reference)
            continue
        source_slides = []
        prepared_groups = []
        needs_section_reflow = False
        project_has_content = any(
            source.get("current") or source.get("next") or source.get("issues") or source.get("supplementary_sections")
            for source in project["slides"]
        )
        for source in project["slides"]:
            if project_has_content and not (
                source.get("current") or source.get("next") or source.get("issues") or source.get("supplementary_sections")
            ):
                continue
            display_name = source["file"]
            if source.get("source_format") == "docx":
                # Word 没有幻灯片布局，先填入项目模板，再与 PPT 源页一样逐框试排。
                word_template = template_slide if template_slide is not None else default_project_template
                if word_template is None:
                    raise ValueError(f"没有可用于 Word 周报的项目页模板：{display_name}")
                cloned = _clone_source_slide(presentation, word_template)
                _expand_template_groups(cloned)
                _replace_title(cloned, project["title"], project["reporter"], title_reference)
                section_shapes = _flow_section_shapes(cloned)
                for key in ("current", "next"):
                    bodies = section_shapes[key]["bodies"]
                    if not bodies:
                        raise ValueError(f"项目“{project['title']}”模板缺少 {SECTION_DISPLAY[key]}正文区域")
                    bodies.sort(key=lambda shape: shape.width * shape.height, reverse=True)
                    bodies[0].text = _normalized_section_content(source.get(key))
                    for shape in bodies[1:]:
                        shape.text = ""
                supplementary = _supplementary_sections(source)
                slots = [parts for key, parts in section_shapes.items() if _is_supplementary_key(key)]
                if supplementary and not slots:
                    raise ValueError(f"项目“{project['title']}”模板缺少补充栏目区域")
                for index, parts in enumerate(slots):
                    section = supplementary[index] if index < len(supplementary) else None
                    label = parts["label"]
                    if label is not None:
                        label.text = section["title"] if section else ""
                        if section:
                            label.name = "weekly-section:" + _supplementary_key(section["title"], section["kind"])
                    bodies = sorted(parts["bodies"], key=lambda shape: shape.width * shape.height, reverse=True)
                    if section and not bodies:
                        raise ValueError(f"项目“{project['title']}”模板缺少补充栏目正文区域")
                    for body_index, shape in enumerate(bodies):
                        shape.text = _normalized_section_content(section["content"]) if section and body_index == 0 else ""
                # Multiple real sections cannot share one template slot. Use
                # the same dynamic flow renderer as overflowing PPT sources.
                needs_section_reflow = needs_section_reflow or len(supplementary) > len(slots)
                _normalize_slide_section_labels(cloned)
                _format_project_slide_bodies(cloned)
                _apply_issue_text_color(cloned, source.get("issues", ""))
                fits, fitted = _fit_compact_slide(cloned)
                prepared_groups.append({
                    "project": project, "source": source, "source_slide": cloned,
                    "template_slide": word_template, "style_reference": cloned,
                    "title_reference": title_reference, "compact_fit": fitted,
                    "fits": fits, "slides": [cloned],
                })
                continue
            if display_name not in source_cache:
                source_cache[display_name] = Presentation(source_lookup[display_name])
            source_slide = source_cache[display_name].slides[source["slide"] - 1]
            if _is_outro_slide(source_slide):
                continue
            source_slides.append(source_slide)
            cloned = _clone_source_slide(presentation, source_slide)
            _expand_template_groups(cloned, sections_only=True)
            _ensure_section_borders(cloned)
            _write_project_sections(cloned, source, format_bodies=False)
            _apply_issue_text_color(cloned, source.get("issues", ""))
            fits, fitted = _fit_compact_slide(cloned)
            prepared_groups.append({
                "project": project,
                "source": source,
                "source_slide": cloned,
                "template_slide": template_slide,
                "style_reference": cloned,
                "title_reference": title_reference,
                "preserve_source_title": True,
                "compact_fit": fitted,
                "fits": fits,
                "slides": [cloned],
            })
        standard_columns = set(_output_sections(project)) == {"current", "next", "issues"}
        if not standard_columns and not needs_section_reflow and all(group["fits"] for group in prepared_groups):
            for group in prepared_groups:
                place_before_outro(group["source_slide"])
                generated_groups.append(group)
            continue

        # Standard projects share the adaptive column policy even when their
        # source boxes fit; distinct custom columns retain fitting source pages.
        for group in prepared_groups:
            _remove_slide(presentation, group["source_slide"])
        flow_template = template_slide if template_slide is not None else (
            source_slides[0] if source_slides else default_project_template
        )
        if flow_template is None:
            raise ValueError(f"没有可用于项目“{project['title']}”的周报模板")
        append_flow_project(project, source_slides, flow_template, title_reference)
    for group in generated_groups:
        if group.get("flow_pages") or "compact_fit" in group:
            continue
        chunks = group["chunks"]
        pagination = group.get("pagination") or _section_pagination(group["source_slide"], chunks)
        page_count = pagination["page_count"]
        if page_count == 1:
            continue
        # Clone the complete source slide before replacing the first page with
        # its chunk, so every continuation starts from the same page layout.
        for page_index in range(1, page_count):
            continuation = _clone_source_slide(presentation, group["source_slide"])
            _set_page_chunks(continuation, chunks, page_index, pagination)
            _apply_issue_text_color(continuation, group.get("issue_text", ""))
            place_after(group["slides"][-1], continuation)
            group["slides"].append(continuation)
        _set_page_chunks(group["slides"][0], chunks, 0, pagination)
        _apply_issue_text_color(group["slides"][0], group.get("issue_text", ""))
    for placeholder_slide in placeholder_slides:
        _remove_slide(presentation, placeholder_slide)
    _retain_single_outro_slide(presentation)
    qa = _generation_qa(presentation, generated_groups, progress_callback=progress_callback)
    _finalize_generated_titles(generated_groups)
    result["qa"] = qa
    result["issues"].extend(qa["issues"])
    result["stats"]["qa_score"] = qa["score"]
    result["stats"]["qa_status"] = qa["status"]
    result["stats"]["error_count"] = sum(item["severity"] == "error" for item in result["issues"])
    result["stats"]["warning_count"] = sum(item["severity"] == "warning" for item in result["issues"])
    presentation.save(target)


def _set_paragraph_xml_text(paragraph_element, value: str) -> None:
    text_nodes = paragraph_element.xpath(".//w:t")
    if text_nodes:
        text_nodes[0].text = value
        for node in text_nodes[1:]:
            node.text = ""
    else:
        run = paragraph_element.find(qn("w:r"))
        if run is None:
            run = paragraph_element.makeelement(qn("w:r"), {})
            paragraph_element.append(run)
        text_node = run.find(qn("w:t"))
        if text_node is None:
            text_node = run.makeelement(qn("w:t"), {})
            run.append(text_node)
        text_node.text = value


def _set_word_paragraph_color(paragraph_element, color: str) -> None:
    for run in paragraph_element.xpath(".//w:r"):
        properties = run.find(qn("w:rPr"))
        if properties is None:
            properties = run.makeelement(qn("w:rPr"), {})
            run.insert(0, properties)
        color_element = properties.find(qn("w:color"))
        if color_element is None:
            color_element = properties.makeelement(qn("w:color"), {})
            properties.append(color_element)
        color_element.set(qn("w:val"), color)


def _paragraph_copy(
    template_paragraph,
    value: str,
    color: str | None = None,
    *,
    bold: bool = False,
    first_line_chars: int | None = None,
    keep_with_next: bool | None = None,
):
    paragraph = deepcopy(template_paragraph._p)
    _set_paragraph_xml_text(paragraph, value)
    if color:
        _set_word_paragraph_color(paragraph, color)
    if bold:
        for run in paragraph.xpath(".//w:r"):
            run.get_or_add_rPr().get_or_add_b().val = True
    if first_line_chars is not None:
        indent = paragraph.get_or_add_pPr().get_or_add_ind()
        for attribute in ("firstLine", "hanging", "hangingChars"):
            indent.attrib.pop(qn(f"w:{attribute}"), None)
        # Word 以百分之一字符记录缩进；只移动首行，不给续行增加缩进。
        indent.set(qn("w:firstLineChars"), str(first_line_chars * 100))
    if keep_with_next is not None:
        paragraph.get_or_add_pPr().keepNext_val = keep_with_next
    return paragraph


WORD_FONT = "宋体"
MEETING_BODY_FONT_SIZE = 12  # 小四


def _set_document_font(document: Document, font_name: str = WORD_FONT) -> None:
    for style in document.styles:
        if not getattr(style, "font", None):
            continue
        style.font.name = font_name
        rfonts = style._element.rPr.rFonts
        for key in ("ascii", "hAnsi", "eastAsia", "cs"):
            rfonts.set(qn(f"w:{key}"), font_name)
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            rpr = run._r.get_or_add_rPr()
            rfonts = rpr.rFonts
            for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                rfonts.set(qn(f"w:{key}"), font_name)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = font_name
                        rpr = run._r.get_or_add_rPr()
                        rfonts = rpr.rFonts
                        for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                            rfonts.set(qn(f"w:{key}"), font_name)


def _set_meeting_body_size(cell) -> None:
    for paragraph in cell.paragraphs:
        if PROJECT_HEADING.match(_text(paragraph.text)):
            continue
        for run in paragraph.runs:
            run.font.size = Pt(MEETING_BODY_FONT_SIZE)


def _meeting_bullet_value(value: str) -> str:
    clean = re.sub(r"^[\s\u00b7\u2022]+", "", _text(value))
    return f"\u00b7 {clean}" if clean else "\u00b7 "


def build_weekly_meeting_document(result: dict[str, Any], target: str | Path, template: str | Path = DOCX_TEMPLATE) -> None:
    """在部门周例会模板中填充项目进度、计划及源文件实际补充栏目。"""
    document = Document(template)
    date_cell = _meeting_date_cell(document)
    _clear_paragraph_runs(date_cell.paragraphs[0], datetime.strptime(result["week_end"], "%Y-%m-%d").strftime("%Y/%m/%d"))
    cell = _meeting_cell(document)
    source_paragraphs = list(cell.paragraphs)
    headings = [paragraph for paragraph in source_paragraphs if PROJECT_HEADING.match(_text(paragraph.text))]
    heading_by_key = {_canonical_title(PROJECT_HEADING.match(_text(paragraph.text)).group(1)): paragraph for paragraph in headings}
    label_templates = {}
    # 0919 参考格式的正文是“· 内容”，不是只有圆点的占位段。
    # 回退到项目标题会把标题行距和 keepNext/keepLines 带入所有正文，造成大片分页留白。
    bullet_template = next((
        paragraph for paragraph in source_paragraphs
        if _text(paragraph.text).startswith(("·", "•"))
    ), None)
    if bullet_template is None:
        raise ValueError("部门周例会模板中没有找到以“·”或“•”开头的正文段落")
    blank_template = next((paragraph for paragraph in source_paragraphs if not paragraph.text.strip()), bullet_template)
    for paragraph in source_paragraphs:
        compact = _text(paragraph.text).replace(" ", "")
        for key, labels in SECTION_LABELS.items():
            if any(compact.startswith(label) for label in labels):
                label_templates.setdefault(key, paragraph)

    def normal_paragraph(template_paragraph, value: str, *, heading: bool = False):
        paragraph = _paragraph_copy(
            template_paragraph, value, "0070C0" if heading else "000000",
            bold=heading, first_line_chars=None if heading else 2,
            keep_with_next=heading,
        )
        # 正文模板可能来自问题列表；显式取消样式继承的编号和红色主题。
        properties = paragraph.get_or_add_pPr()
        numbering = properties.get_or_add_numPr()
        numbering.get_or_add_numId().val = 0
        for run in paragraph.xpath(".//w:r"):
            run_properties = run.get_or_add_rPr()
            run_properties.get_or_add_b().val = heading
            for color_element in run_properties.findall(qn("w:color")):
                for attribute in ("themeColor", "themeTint", "themeShade"):
                    color_element.attrib.pop(qn(f"w:{attribute}"), None)
        return paragraph

    tc = cell._tc
    for paragraph in list(cell.paragraphs):
        _remove_element(paragraph._p)
    for project in result["meeting_projects"]:
        heading = heading_by_key.get(project["key"])
        heading_template = heading or headings[0]
        heading_text = _text(heading.text) if heading else f"{project['title']}（汇报人：{project['reporter'] or '待补充'}）"
        tc.append(_paragraph_copy(heading_template, heading_text))
        for key in ("current", "next"):
            label_template = label_templates.get(key) or bullet_template
            # 小节名只跟随后面的首段；正文沿用参考文档的自然跨页设置。
            tc.append(_paragraph_copy(label_template, SECTION_DISPLAY[key], bold=True, keep_with_next=True))
            values = [line for line in _normalized_section_content(project.get(key)).splitlines() if line.strip()]
            for value in values:
                tc.append(_paragraph_copy(bullet_template, _meeting_bullet_value(value),
                                          first_line_chars=2, keep_with_next=False))
            tc.append(_paragraph_copy(blank_template, "", keep_with_next=False))
        for section in _supplementary_sections(project):
            title = section["title"].rstrip("：:") + "："
            content = _normalized_section_content(section["content"])
            if section["kind"] == "issue":
                label_template = label_templates.get("issues") or bullet_template
                tc.append(_paragraph_copy(label_template, title, "FF0000", bold=True, keep_with_next=True))
                values = [line for line in content.splitlines() if line.strip()]
                for value in values:
                    tc.append(_paragraph_copy(bullet_template, _meeting_bullet_value(value), "FF0000",
                                              first_line_chars=2, keep_with_next=False))
            else:
                label_template = label_templates.get("current") or bullet_template
                tc.append(normal_paragraph(label_template, title, heading=True))
                values = content.splitlines()
                for value in values:
                    tc.append(normal_paragraph(bullet_template, value))
            tc.append(_paragraph_copy(blank_template, "", keep_with_next=False))
    _set_meeting_body_size(cell)
    _set_document_font(document)
    document.save(target)


def _excel_safe(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    return "'" + value if value[:1] in {"=", "+", "-", "@"} else value


def export_weekly_report_xlsx(result: dict[str, Any], target: str | Path) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    header_fill = PatternFill("solid", fgColor="181816")
    header_font = Font(color="FFFFFF", bold=True)

    def supplementary_text(project: dict[str, Any]) -> str:
        return "\n\n".join(
            section["title"].rstrip("：:") + "：\n" + _normalized_section_content(section["content"])
            for section in _supplementary_sections(project)
        ) or "无"

    def add_sheet(title: str, headers: list[str], rows: list[list[Any]]) -> None:
        sheet = workbook.create_sheet(title)
        sheet.append(headers)
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in rows:
            sheet.append([_excel_safe(value) for value in row])
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column in sheet.columns:
            width = min(max(max(len(str(cell.value or "")) for cell in column) + 2, 10), 54)
            sheet.column_dimensions[column[0].column_letter].width = width
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

    add_sheet(
        "审核结果",
        ["严重程度", "文件", "页码", "位置", "项目", "检查项", "描述", "建议"],
        [[item["severity"], item["file"], item["slide"], item["location"], item["project"], item["label"], item["detail"], item["suggestion"]] for item in result["issues"]],
    )
    add_sheet(
        "项目整合",
        ["项目", "汇报人", "状态", "源文件", "源页码", "本周进展", "下周计划", "补充内容"],
        [[item["title"], item["reporter"], item["status"], "、".join(item["source_files"]), "、".join(str(slide["slide"]) for slide in item["slides"]), _normalized_section_content(item.get("current")), _normalized_section_content(item.get("next")), supplementary_text(item)] for item in result["projects"]],
    )
    add_sheet(
        "周例会内容",
        ["项目", "汇报人", "本周进展", "下周计划", "补充内容"],
        [[item["title"], item["reporter"], _normalized_section_content(item.get("current")), _normalized_section_content(item.get("next")), supplementary_text(item)] for item in result["meeting_projects"]],
    )
    add_sheet(
        "文件目录",
        ["路径", "类型", "大小", "状态"],
        [[item.get("path", ""), item.get("kind", ""), item.get("size", ""), item.get("status", "")] for item in result["sources"]],
    )
    qa = result.get("qa") or {}
    add_sheet(
        "生成质量核验",
        ["轮次", "内容准确度", "模板样式准确度", "版面稳定度", "自动修复数", "问题数", "连续稳定"],
        [[
            item.get("round", ""),
            item.get("content_score", ""),
            item.get("style_score", ""),
            item.get("layout_score", ""),
            item.get("repairs", ""),
            item.get("issues", ""),
            "是" if item.get("stable") else "否",
        ] for item in qa.get("rounds", [])],
    )
    workbook.save(target)
    workbook.close()
