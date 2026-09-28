from pathlib import Path
import re

import pytest
from docx import Document
from docx.oxml.ns import qn
from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.ns import qn as pptx_qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Centipoints, Inches, Pt

from weekly_report_processor import (
    PPT_TEMPLATE,
    _flatten_shapes,
    _flow_line_ends,
    _flow_page_plan,
    _flow_shape_metrics,
    _prefer_later_section_content,
    _section_from_entries,
    _template_projects,
    _trim_body_whitespace,
    build_weekly_meeting_document,
    build_weekly_presentation,
    parse_document_source,
    parse_presentation_source,
    process_weekly_report,
)


def test_next_week_duplicate_is_not_kept_in_current_week() -> None:
    values = {
        "current": "已完成联调\n电子文档中心系统【并行测试】\n用户权限中心系统【并行测试】",
        "next": "电子文档中心系统【并行测试】\n用户权限中心系统【并行测试】",
        "issues": "无",
    }

    result = _prefer_later_section_content(values)

    assert result["current"] == "已完成联调"
    assert result["next"] == values["next"]
    assert result["issues"] == "无"


def test_issue_section_uses_the_right_column_boundary() -> None:
    entries = [
        {"text": "本周工作完成情况", "top": 994410, "left": 472440},
        {"text": "本周问题", "top": 990888, "left": 7780146},
        {"text": "下周计划", "top": 4932680, "left": 469900},
        {"text": "数据库连接异常", "top": 994410, "left": 7768590},
    ]

    assert _section_from_entries(entries, "issues") == "数据库连接异常"


def test_untemplated_project_ppt_is_added_to_weekly_presentation(tmp_path: Path) -> None:
    template = Presentation(PPT_TEMPLATE)
    presentation = Presentation()
    presentation.slide_width = template.slide_width
    presentation.slide_height = template.slide_height
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])

    for text, left, top, width, height in (
        ("新增自动识别项目（汇报人：张三）", 1, 0.3, 8, 0.4),
        ("本周工作完成情况", 0.4, 1.0, 1.5, 0.4),
        ("已完成联调", 1.2, 1.0, 5, 1),
        ("下周计划", 0.4, 2.2, 1.5, 0.4),
        ("安排上线", 1.2, 2.2, 5, 1),
        ("本周问题", 0.4, 3.4, 1.5, 0.4),
        ("风险项", 1.2, 3.4, 5, 1),
    ):
        slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height)).text = text

    source = tmp_path / "新增项目.pptx"
    presentation.save(source)
    result = process_weekly_report([(source, source.name)], [])

    added = next(item for item in result["projects"] if item["title"] == "新增自动识别项目")
    assert added["source_kind"] == "自动添加项目"
    assert added["current"] == "已完成联调"

    presentation_path = tmp_path / "项目周报.pptx"
    build_weekly_presentation(result, {source.name: source}, presentation_path)
    output = Presentation(presentation_path)
    slide = next(
        item for item in output.slides
        if "新增自动识别项目" in "".join(shape.text for shape in item.shapes if getattr(shape, "has_text_frame", False))
    )
    issue_body = next(entry for entry in _flatten_shapes(slide.shapes) if "风险项" in entry["text"])
    assert all(str(run.font.color.rgb) == "FF0000" for paragraph in issue_body["shape"].text_frame.paragraphs for run in paragraph.runs)

    meeting = next(item for item in result["meeting_projects"] if item["title"] == "新增自动识别项目")
    assert meeting["next"] == "安排上线"

    document_path = tmp_path / "部门周例会.docx"
    build_weekly_meeting_document(result, document_path)
    document_text = "\n".join(paragraph.text for table in Document(document_path).tables for row in table.rows for cell in row.cells for paragraph in cell.paragraphs)
    assert "新增自动识别项目（汇报人：张三）" in document_text
    assert "已完成联调" in document_text
    assert "安排上线" in document_text
    issue_paragraphs = [
        paragraph for table in Document(document_path).tables for row in table.rows for cell in row.cells
        for paragraph in cell.paragraphs
        if paragraph.text.startswith("本周问题") or paragraph.text.endswith("风险项")
    ]
    assert len(issue_paragraphs) >= 2
    assert all(
        run._r.find(qn("w:rPr")).find(qn("w:color")).get(qn("w:val")) == "FF0000"
        for paragraph in issue_paragraphs for run in paragraph.runs
    )


def test_word_template_only_project_is_added_to_weekly_presentation(tmp_path: Path) -> None:
    template = Presentation(PPT_TEMPLATE)
    presentation = Presentation()
    presentation.slide_width = template.slide_width
    presentation.slide_height = template.slide_height
    presentation.slides.add_slide(presentation.slide_layouts[6]).shapes.add_textbox(
        Inches(1), Inches(1), Inches(4), Inches(1)
    ).text = "项目周报"
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])

    for text, left, top, width, height in (
        ("demo开发、客户交流、其他（汇报人：郑乐园）", 0.5, 0.3, 8, 0.4),
        ("本周进度", 0.4, 1.0, 1.5, 0.4),
        ("已完成客户交流", 1.2, 1.0, 5, 1),
        ("下周计划", 0.4, 2.2, 1.5, 0.4),
        ("继续跟进商机", 1.2, 2.2, 5, 1),
        ("问题", 0.4, 3.4, 1.5, 0.4),
        ("暂无", 1.2, 3.4, 5, 1),
    ):
        slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height)).text = text
    presentation.slides.add_slide(presentation.slide_layouts[6]).shapes.add_textbox(
        Inches(1), Inches(1), Inches(4), Inches(1)
    ).text = "期待与您携手共赢"

    source = tmp_path / "demo开发.pptx"
    presentation.save(source)
    result = process_weekly_report([(source, source.name)], [])

    project = next(item for item in result["projects"] if item["title"] == "demo开发、客户交流、其他")
    assert project["source_kind"] == "周例会模板项目"
    assert project["slides"][0]["file"] == source.name

    presentation_path = tmp_path / "项目周报.pptx"
    build_weekly_presentation(result, {source.name: source}, presentation_path)
    output = Presentation(presentation_path)
    assert any(
        "demo开发、客户交流、其他" in "".join(
            shape.text for shape in output_slide.shapes if getattr(shape, "has_text_frame", False)
        )
        for output_slide in output.slides
    )


def _build_flow_fixture(tmp_path: Path, source_format: str, pages: list[dict[str, str]], style_pptx_body=None):
    """Build a real source file and exercise the public presentation builder."""
    project = {**_template_projects()[0][0], "reporter": "分页回归"}
    source = tmp_path / f"分页回归来源.{source_format}"
    if source_format == "docx":
        document = Document()
        document.add_paragraph(f"{project['title']}（汇报人：{project['reporter']}）")
        for page in pages:
            for key, label in (("current", "本周进度"), ("issues", "本周问题"), ("next", "下周计划")):
                document.add_paragraph(label)
                for line in page.get(key, "").split("\n"):
                    document.add_paragraph(line)
        document.save(source)
        parsed = parse_document_source(source, source.name, [project])
    else:
        template = Presentation(PPT_TEMPLATE)
        presentation = Presentation()
        presentation.slide_width = template.slide_width
        presentation.slide_height = template.slide_height
        width, height = presentation.slide_width, presentation.slide_height
        right_left = max(7400000, int(width * 0.65))
        for page in pages:
            slide = presentation.slides.add_slide(presentation.slide_layouts[6])
            title = slide.shapes.add_textbox(int(width * 0.08), int(height * 0.04), int(width * 0.8), int(height * 0.09))
            title.text = f"{project['title']}（汇报人：{project['reporter']}）"
            for key, label, left, top, box_width, box_height in (
                ("current", "本周进度", int(width * 0.08), int(height * 0.19), right_left - int(width * 0.10), int(height * 0.44)),
                ("next", "下周计划", int(width * 0.08), int(height * 0.65), right_left - int(width * 0.10), int(height * 0.28)),
                ("issues", "本周问题", right_left, int(height * 0.19), int(width * 0.93) - right_left, int(height * 0.74)),
            ):
                label_left = int(width * (0.94 if key == "issues" else 0.03))
                label_shape = slide.shapes.add_textbox(label_left, top, int(width * 0.04), box_height)
                label_shape.name = f"legacy:{key}:label"
                label_shape.text = label
                body = slide.shapes.add_textbox(left, top, box_width, box_height)
                body.name = f"legacy:{key}:body"
                body.text = page.get(key) or "无"
                for paragraph in body.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(10.5)
                if style_pptx_body is not None:
                    style_pptx_body(key, body)
        presentation.save(source)
        parsed = parse_presentation_source(source, source.name, [project])

    project["slides"] = parsed["slides"]
    for key in ("current", "issues", "next"):
        project[key] = "\n".join(dict.fromkeys(item[key] for item in parsed["slides"] if item[key])) or "无"
    result = {"projects": [project], "week_end": "2026-09-26", "issues": [], "stats": {}}
    target = tmp_path / "分页回归成品.pptx"
    build_weekly_presentation(result, {source.name: source}, target)
    return Presentation(target), project, result


def _flow_pages(presentation):
    pages = []
    for slide in presentation.slides:
        bodies = [
            shape for shape in slide.shapes
            if shape.name.startswith("weekly-flow:") and shape.name.endswith(":body")
        ]
        if bodies:
            pages.append((slide, sorted(bodies, key=lambda shape: shape.top)))
    return pages


def _content_text(value: str) -> str:
    return re.sub(r"[\s·\u200b]", "", value)


def _assert_flow_content(presentation, project) -> list:
    pages = _flow_pages(presentation)
    assert pages, "长内容项目应生成带显式栏目标识的单列页面"
    sections = []
    fragment_ids = []
    values = {key: [] for key in ("current", "issues", "next")}
    body_lefts, body_widths, page_tops, page_bottoms = set(), set(), set(), set()
    for slide, bodies in pages:
        names = {shape.name: shape for shape in slide.shapes}
        page_font = bodies[0].text_frame.paragraphs[0].font.size.pt
        assert page_font in (10.5, 10, 9.5, 9)
        page_line_height = Centipoints(round(2200 * page_font / 14))
        assert not any(shape.name.startswith("legacy:") for shape in slide.shapes)
        page_tops.add(bodies[0].top)
        page_bottoms.add(bodies[-1].top + bodies[-1].height)
        for index, body in enumerate(bodies):
            _, key, fragment_id, _ = body.name.split(":")
            sections.append(key)
            fragment_ids.append(int(fragment_id))
            values[key].append(body.text)
            body_lefts.add(body.left)
            body_widths.add(body.width)
            assert body.width > presentation.slide_width * 0.75
            label = names[f"weekly-flow:{key}:{fragment_id}:label"]
            border = names[f"weekly-flow:{key}:{fragment_id}:border"]
            assert label.top == body.top
            assert label.left + label.width <= body.left
            assert border.top == body.top and border.height == body.height
            assert border.left == body.left and border.width == body.width
            assert body.height == label.height
            assert label.height >= Pt(84)
            assert body.top + body.height <= presentation.slide_height
            assert _content_text(body.text)
            if index:
                previous = bodies[index - 1]
                assert body.top - (previous.top + previous.height) == Pt(6)
            frame = body.text_frame
            assert frame.margin_left == frame.margin_right == Pt(8)
            assert frame.margin_top == frame.margin_bottom == Pt(6)
            assert frame.vertical_anchor == MSO_ANCHOR.TOP
            for paragraph in frame.paragraphs:
                assert paragraph.line_spacing == page_line_height
                assert paragraph.space_before == paragraph.space_after == Pt(0)
                assert paragraph.font.size == Pt(page_font)
                end_properties = paragraph._p.find(pptx_qn("a:endParaRPr"))
                assert end_properties is not None
                assert end_properties.get("sz") == str(round(page_font * 100))
            metrics_match, required_height = _flow_shape_metrics(
                body, font_size=page_font, line_height=page_line_height,
            )
            assert metrics_match and required_height is not None and required_height <= body.height
            for paragraph in label.text_frame.paragraphs:
                assert paragraph.line_spacing == Pt(18)
                assert paragraph.space_before == paragraph.space_after == Pt(0)
                assert paragraph.font.size == Pt(14)
                assert all(run.font.size == Pt(14) for run in paragraph.runs)
                end_properties = paragraph._p.find(pptx_qn("a:endParaRPr"))
                assert end_properties is not None and end_properties.get("sz") == "1400"
            runs = [run for paragraph in body.text_frame.paragraphs for run in paragraph.runs if run.text]
            assert runs and all(run.font.size == Pt(page_font) for run in runs)
            if key == "issues":
                assert all(str(run.font.color.rgb) == "FF0000" for run in runs)
    assert len(body_lefts) == len(body_widths) == 1
    assert len(page_tops) == len(page_bottoms) == 1
    assert fragment_ids == list(range(len(fragment_ids)))
    assert sections == sorted(sections, key=("current", "next", "issues").index)
    for key, fragments in values.items():
        assert _content_text("".join(fragments)) == _content_text(project[key])
    return pages


@pytest.mark.parametrize("source_format", ["pptx", "docx"])
def test_long_progress_uses_full_width_until_complete(tmp_path: Path, source_format: str) -> None:
    current = "\n".join(f"{index}. 已完成第{index}批数据核对并交付验收记录，后续跟进业务确认。" for index in range(1, 71))
    output, project, result = _build_flow_fixture(tmp_path, source_format, [
        {"current": current, "issues": "1. 等待联调环境恢复。", "next": "1. 组织验收并完成上线。"},
    ])

    pages = _assert_flow_content(output, project)
    progress_pages = [index for index, (_, bodies) in enumerate(pages) if any(":current:" in body.name for body in bodies)]
    assert len(progress_pages) > 1
    for _, bodies in pages[:progress_pages[-1]]:
        assert all(":current:" in body.name for body in bodies)
    # 每条都是可放入一页的独立段落，分页应保留完整条目。
    for line in current.splitlines():
        assert sum(_content_text(line) in _content_text(body.text) for _, bodies in pages for body in bodies) == 1
    assert not any(issue["code"] == "generated_content_mismatch" for issue in result["qa"]["issues"])


def test_long_issue_column_can_follow_plan_on_the_same_page(tmp_path: Path) -> None:
    issues = "\n".join(f"{index}. 风险第{index}项：等待外部接口恢复后继续核验。" for index in range(1, 81))
    output, project, _ = _build_flow_fixture(tmp_path, "pptx", [
        {"current": "1. 已完成本周开发。", "issues": issues, "next": "1. 继续处理剩余任务。"},
    ])

    pages = _assert_flow_content(output, project)
    assert [body.name.split(":")[1] for body in pages[0][1]][:3] == ["current", "next", "issues"]
    first_issue_page = min(index for index, (_, bodies) in enumerate(pages) if any(":issues:" in body.name for body in bodies))
    assert all(":next:" not in body.name for _, bodies in pages[first_issue_page + 1:] for body in bodies)


def test_project_source_pages_are_grouped_before_flow_pagination(tmp_path: Path) -> None:
    first = "\n".join(f"{index}. 第一部分数据已复核，业务系统联调和验收记录整理完成。" for index in range(1, 36))
    second = "\n".join(f"{index}. 第二部分数据已复核，剩余接口联调和验收记录整理完成。" for index in range(36, 71))
    output, project, _ = _build_flow_fixture(tmp_path, "pptx", [
        {"current": first, "issues": "1. 第一部分依赖待确认。", "next": "1. 第一部分安排上线。"},
        {"current": second, "issues": "2. 第二部分依赖待确认。", "next": "2. 第二部分安排上线。"},
        {"current": second, "issues": "2. 第二部分依赖待确认。", "next": "2. 第二部分安排上线。"},
    ])

    pages = _assert_flow_content(output, project)
    assert project["current"] == f"{first}\n{second}"
    last_current_page = max(index for index, (_, bodies) in enumerate(pages) if any(":current:" in body.name for body in bodies))
    assert all(":current:" in body.name for _, bodies in pages[:last_current_page] for body in bodies)


@pytest.mark.parametrize("section", ["current", "next"])
def test_single_oversized_paragraph_does_not_repeat_numbering(tmp_path: Path, section: str) -> None:
    paragraph = "1. " + "持续核对业务数据并补充接口验收记录，" * 350
    values = {"current": "1. 本周开发完成。", "issues": "无", "next": "1. 下周组织验收。"}
    values[section] = paragraph
    output, project, _ = _build_flow_fixture(tmp_path, "docx", [values])

    pages = _assert_flow_content(output, project)
    bodies = [body for _, page_bodies in pages for body in page_bodies if f":{section}:" in body.name]
    assert len(bodies) > 1
    assert sum(body.text.count("1.") for body in bodies) == 1
    assert all(not body.text.lstrip().startswith(("·", "•", "1.")) for body in bodies[1:])


def test_empty_sections_share_remaining_space_without_extra_empty_page(tmp_path: Path) -> None:
    values = {
        "current": "\n".join(f"{index}. 本批次内容处理完成。" for index in range(1, 76)),
        "issues": "",
        "next": "",
    }
    output, project, _ = _build_flow_fixture(tmp_path, "docx", [values])

    pages = _assert_flow_content(output, project)
    for key in ("issues", "next"):
        assert sum(f":{key}:" in body.name for _, bodies in pages for body in bodies) == 1
    assert all(
        _content_text(body.text) == "无"
        for _, bodies in pages for body in bodies if ":issues:" in body.name or ":next:" in body.name
    )
    assert all(any(_content_text(body.text) for body in bodies) for _, bodies in pages)
    assert sum(body.text == "无" for _, bodies in pages for body in bodies) == 2


def _flow_test_paragraph(value: str) -> dict:
    element = OxmlElement("a:p")
    for index, line in enumerate(value.split("\v")):
        if index:
            element.append(OxmlElement("a:br"))
        run = OxmlElement("a:r")
        text = OxmlElement("a:t")
        text.text = line
        run.append(text)
        element.append(run)
    return {"text": value, "xml": element}


def _assert_flow_plan(pages: list, sections: dict, geometry: dict) -> None:
    actual = {key: [] for key in ("current", "next", "issues")}
    order = []
    for page in pages:
        assert page
        font_size = page[0]["font_size"]
        assert font_size in (10.5, 10, 9.5, 9)
        line_height = Centipoints(round(2200 * font_size / 14))
        cursor = geometry["top"]
        for fragment in page:
            assert fragment["font_size"] == font_size
            assert fragment["line_height"] == line_height
            assert fragment["top"] == cursor
            measured_lines = sum(len(_flow_line_ends(
                paragraph, geometry["body_width"], font_size=font_size,
            )) for paragraph in fragment["paragraphs"])
            assert fragment["line_count"] == measured_lines
            assert fragment["height"] >= max(Pt(84), measured_lines * line_height + Pt(16))
            cursor += fragment["height"] + Pt(6)
            actual[fragment["section"]].extend(item["text"] for item in fragment["paragraphs"])
            order.append(fragment["section"])
        assert cursor - Pt(6) == geometry["bottom"]
    assert order == sorted(order, key=("current", "next", "issues").index)
    for key, paragraphs in sections.items():
        assert _content_text("".join(actual[key])) == _content_text("".join(item["text"] for item in paragraphs))


def test_body_whitespace_trimming_preserves_inner_breaks_and_rich_paragraphs() -> None:
    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    body = slide.shapes.add_textbox(Pt(20), Pt(100), Pt(400), Pt(200))
    body.text = "\n\n\v\v第一段 API\v继续联调\n\n末段验收\v\v\n\n"
    paragraph = body.text_frame.paragraphs[2]
    paragraph.level = 1
    properties = paragraph._p.get_or_add_pPr()
    numbering = OxmlElement("a:buAutoNum")
    numbering.set("type", "arabicPeriod")
    numbering.set("startAt", "3")
    properties.append(numbering)
    for run in paragraph.runs:
        run.font.bold = True

    _trim_body_whitespace(body.text_frame)

    assert body.text == "第一段 API\v继续联调\n\n末段验收"
    assert len(body.text_frame.paragraphs) == 3
    first = body.text_frame.paragraphs[0]
    assert first.level == 1
    assert all(run.font.bold for run in first.runs if run.text)
    preserved_numbering = first._p.find(pptx_qn("a:pPr")).find(pptx_qn("a:buAutoNum"))
    assert preserved_numbering is not None and preserved_numbering.get("startAt") == "3"
    assert body.text_frame.paragraphs[1].text == ""


def test_trimming_an_empty_template_body_keeps_its_geometry() -> None:
    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    body = slide.shapes.add_textbox(Pt(20), Pt(100), Pt(400), Pt(200))
    body.text = "\n\v\n\n"

    _trim_body_whitespace(body.text_frame)

    assert not body.text.strip()
    assert len(body.text_frame.paragraphs) == 1
    assert (body.left, body.top, body.width, body.height) == (Pt(20), Pt(100), Pt(400), Pt(200))


def test_mixed_text_width_uses_character_classes_and_keeps_soft_break_offsets() -> None:
    def ends(value: str):
        return _flow_line_ends(_flow_test_paragraph(value), Pt(140), margin_left=0, margin_right=0, font_size=10)

    assert len(ends("API123 " * 8)) < len(ends("接口联调完成。" * 8))
    assert len(ends("ＡＢＣ１２３" * 8)) > len(ends("ABC123" * 8))
    assert len(ends("a\u0301" * 40)) == len(ends("a" * 40))
    assert len(ends("x\t" * 20)) >= len(ends("x " * 20))
    assert ends("API\v中文123") == [4, 9]
    long_word = "LibreOfficeToken" * 40
    word_ends = ends(long_word)
    assert len(word_ends) > 1 and word_ends[-1] == len(long_word)
    assert word_ends == sorted(set(word_ends))
    assert all(right > left for left, right in zip([0, *word_ends], word_ends))


def test_flow_line_width_respects_paragraph_indents() -> None:
    paragraph = _flow_test_paragraph("接口核验数据" * 12)
    unindented = _flow_line_ends(paragraph, Pt(140), font_size=10)
    properties = OxmlElement("a:pPr")
    properties.set("marL", str(Pt(35)))
    properties.set("marR", str(Pt(20)))
    properties.set("indent", str(Pt(15)))
    paragraph["xml"].insert(0, properties)

    indented = _flow_line_ends(paragraph, Pt(140), font_size=10)

    assert len(indented) > len(unindented)
    assert indented[0] < unindented[0]
    assert indented[-1] == len(paragraph["text"])


def test_flow_page_shrinks_before_splitting_a_complete_paragraph() -> None:
    sections = {
        "current": [
            _flow_test_paragraph("完成开发。\v完成联调。\v完成测试。\v完成验收。"),
            _flow_test_paragraph("后续数据已核对。\v后续记录已归档。"),
        ],
        "next": [],
        "issues": [],
    }
    geometry = {"top": Pt(100), "bottom": Pt(212), "body_width": Pt(600)}

    pages = _flow_page_plan(sections, geometry)

    _assert_flow_plan(pages, sections, geometry)
    assert len(pages) == 1
    assert pages[0][0]["font_size"] == 10
    assert [item["text"] for item in pages[0][0]["paragraphs"]] == [item["text"] for item in sections["current"]]


def test_flow_page_uses_remaining_lines_after_reaching_minimum_font() -> None:
    continued = _flow_test_paragraph("任务一完成。\v任务二完成。\v任务三完成。\v任务四完成。\v任务五完成。\v任务六完成。")
    properties = OxmlElement("a:pPr")
    numbering = OxmlElement("a:buAutoNum")
    numbering.set("type", "arabicPeriod")
    numbering.set("startAt", "3")
    properties.append(numbering)
    continued["xml"].insert(0, properties)
    sections = {
        "current": [_flow_test_paragraph("接口已联调。\v数据已核验。"), continued],
        "next": [],
        "issues": [],
    }
    geometry = {"top": Pt(100), "bottom": Pt(200), "body_width": Pt(600)}

    pages = _flow_page_plan(sections, geometry)

    _assert_flow_plan(pages, sections, geometry)
    assert len(pages) == 2
    assert pages[0][0]["font_size"] == 9
    assert [page[0]["line_count"] for page in pages] == [5, 3]
    assert pages[1][0]["font_size"] == 10.5
    first_properties = pages[0][0]["paragraphs"][-1]["xml"].find(pptx_qn("a:pPr"))
    assert first_properties.find(pptx_qn("a:buAutoNum")).get("startAt") == "3"
    continued_properties = pages[1][0]["paragraphs"][0]["xml"].find(pptx_qn("a:pPr"))
    assert continued_properties.find(pptx_qn("a:buAutoNum")) is None
    assert continued_properties.find(pptx_qn("a:buNone")) is not None


@pytest.mark.parametrize(("pending_lines", "expected_counts"), [(3, [3, 3]), (4, [5, 2])])
def test_normal_paragraph_split_leaves_at_least_two_lines_on_each_page(
    pending_lines: int, expected_counts: list[int],
) -> None:
    sections = {
        "current": [
            _flow_test_paragraph("开发完成。\v联调完成。\v验收完成。"),
            _flow_test_paragraph("\v".join(f"后续任务{index}完成。" for index in range(pending_lines))),
        ],
        "next": [],
        "issues": [],
    }
    geometry = {"top": Pt(100), "bottom": Pt(200), "body_width": Pt(600)}

    pages = _flow_page_plan(sections, geometry)

    _assert_flow_plan(pages, sections, geometry)
    assert [page[0]["line_count"] for page in pages] == expected_counts
    assert len(pages[-1][0]["paragraphs"][0]["text"].split("\v")) >= 2


def test_flow_heading_moves_with_following_body_when_page_has_one_line_left() -> None:
    heading = _flow_test_paragraph("六、数据采集与入库")
    properties = OxmlElement("a:rPr")
    properties.set("b", "1")
    heading["xml"].find(pptx_qn("a:r")).insert(0, properties)
    sections = {
        "current": [
            _flow_test_paragraph("接口已核对。\v数据库已同步。\v环境已就绪。\v权限已确认。"),
            heading,
            _flow_test_paragraph("网站验收完成。\v封装验证完成。\v入库记录已交付。"),
        ],
        "next": [],
        "issues": [],
    }
    geometry = {"top": Pt(100), "bottom": Pt(200), "body_width": Pt(600)}

    pages = _flow_page_plan(sections, geometry)

    _assert_flow_plan(pages, sections, geometry)
    assert len(pages) == 2
    assert all(item["text"] != heading["text"] for item in pages[0][0]["paragraphs"])
    assert pages[1][0]["paragraphs"][0]["text"] == heading["text"]
    assert len(pages[1][0]["paragraphs"]) >= 2


def test_flow_columns_share_extra_height_by_visual_lines() -> None:
    sections = {
        "current": [_flow_test_paragraph("完成交付")],
        "issues": [_flow_test_paragraph("接口待确认\v环境待恢复")],
        "next": [_flow_test_paragraph("核对数据\v回归接口\v组织验收\v安排上线")],
    }

    pages = _flow_page_plan(sections, {
        "top": Pt(100), "bottom": Pt(500), "body_width": Pt(600),
    })

    assert len(pages) == 1
    fragments = pages[0]
    assert [item["section"] for item in fragments] == ["current", "next", "issues"]
    assert [item["line_count"] for item in fragments] == [1, 4, 2]
    # 84pt protects the four-line label; even four 16.5pt body lines need only 82pt.
    # After two 6pt gaps and these minima, 136pt is shared in a 1:4:2 ratio.
    extra_heights = [item["height"] - minimum for item, minimum in zip(
        fragments, [Pt(84), Pt(84), Pt(84)],
    )]
    assert sum(extra_heights) == Pt(136)
    for extra, weight in zip(extra_heights, [1, 4, 2]):
        # Rounding the first two shares can leave less than two EMU for the last.
        assert abs(extra * 7 - Pt(136) * weight) <= 14
    assert fragments[0]["top"] == Pt(100)
    for previous, current in zip(fragments, fragments[1:]):
        assert current["top"] == previous["top"] + previous["height"] + Pt(6)
    assert fragments[-1]["top"] + fragments[-1]["height"] == Pt(500)


@pytest.mark.parametrize(("body_height", "first_page_lines"), [(86.69, 4), (86.70, 5)])
def test_flow_line_capacity_reserves_padding_and_bottom_safety(body_height: float, first_page_lines: int) -> None:
    lines = [f"完成第{index}项" for index in range(1, 7)]
    sections = {
        "current": [_flow_test_paragraph(line) for line in lines],
        "issues": [],
        "next": [],
    }

    pages = _flow_page_plan(sections, {
        "top": Pt(100), "bottom": Pt(100 + body_height), "body_width": Pt(600),
    })

    assert len(pages) == 2
    assert all(len(page) == 1 for page in pages)
    assert [page[0]["line_count"] for page in pages] == [first_page_lines, 6 - first_page_lines]
    assert [paragraph["text"] for page in pages for fragment in page for paragraph in fragment["paragraphs"]] == lines
    assert all(page[0]["top"] + page[0]["height"] == Pt(100 + body_height) for page in pages)


@pytest.mark.parametrize("progress_lines", [6, 11])
def test_no_issues_and_short_plan_tail_fill_the_content_area(progress_lines: int) -> None:
    sections = {
        "current": [_flow_test_paragraph(f"完成第{index}项") for index in range(1, progress_lines + 1)],
        "issues": [_flow_test_paragraph("无")],
        "next": [_flow_test_paragraph("继续处理剩余任务")],
    }

    pages = _flow_page_plan(sections, {
        "top": Pt(100), "bottom": Pt(360), "body_width": Pt(600),
    })

    assert len(pages) == 2
    # With six progress lines, the plan alone fits the first page. It must
    # nevertheless travel with “无”, which would otherwise become a lone tail.
    assert [item["section"] for item in pages[0]] == ["current"]
    assert [item["section"] for item in pages[1]] == ["next", "issues"]
    assert [item["line_count"] for item in pages[1]] == [1, 1]
    assert [item["height"] for item in pages[1]] == [Pt(127), Pt(127)]
    assert pages[1][0]["paragraphs"][0]["text"] == "继续处理剩余任务"
    assert pages[1][1]["paragraphs"][0]["text"] == "无"
    assert pages[1][1]["top"] == Pt(233)
    assert all(page[-1]["top"] + page[-1]["height"] == Pt(360) for page in pages)


def test_pptx_flow_preserves_soft_breaks_bold_levels_and_numbered_continuations(tmp_path: Path) -> None:
    current = "联调交付：API Excel 数据核对。\v" + "持续核对中文记录、LibreOffice 文件及 token 参数，" * 350

    def style_body(key, body):
        if key != "current":
            return
        paragraph = body.text_frame.paragraphs[0]
        paragraph.level = 1
        paragraph.font.size = Pt(28)
        properties = paragraph._p.get_or_add_pPr()
        properties.set("marL", str(Pt(28)))
        properties.set("indent", str(-Pt(14)))
        numbering = OxmlElement("a:buAutoNum")
        numbering.set("type", "arabicPeriod")
        numbering.set("startAt", "3")
        properties.append(numbering)
        end_properties = OxmlElement("a:endParaRPr")
        end_properties.set("sz", "4800")
        paragraph._p.append(end_properties)
        for run in paragraph.runs:
            run.font.size = Pt(28)
            run.font.bold = True

    output, project, _ = _build_flow_fixture(tmp_path, "pptx", [{
        "current": current, "issues": "待确认 API 权限。", "next": "完成回归后安排上线。",
    }], style_pptx_body=style_body)

    pages = _assert_flow_content(output, project)
    bodies = [body for _, page_bodies in pages for body in page_bodies if ":current:" in body.name]
    assert len(bodies) > 1
    paragraphs = [paragraph for body in bodies for paragraph in body.text_frame.paragraphs]
    assert "\v" in bodies[0].text
    assert all(paragraph.level == 1 for paragraph in paragraphs)
    assert all(run.font.bold for paragraph in paragraphs for run in paragraph.runs if run.text)
    numbering = paragraphs[0]._p.find(pptx_qn("a:pPr")).find(pptx_qn("a:buAutoNum"))
    assert numbering is not None and numbering.get("startAt") == "3"
    for paragraph in paragraphs[1:]:
        properties = paragraph._p.find(pptx_qn("a:pPr"))
        assert properties.find(pptx_qn("a:buAutoNum")) is None
        assert properties.find(pptx_qn("a:buNone")) is not None


@pytest.mark.parametrize("source_format", ["pptx", "docx"])
def test_short_project_keeps_existing_layout(tmp_path: Path, source_format: str) -> None:
    values = {"current": "1. 完成接口联调。", "issues": "1. 等待验收反馈。", "next": "1. 安排上线。"}
    output, project, _ = _build_flow_fixture(tmp_path, source_format, [values])

    assert not _flow_pages(output)
    project_slides = [
        slide for slide in output.slides
        if any(project["title"] in entry["text"] and "汇报人" in entry["text"] for entry in _flatten_shapes(slide.shapes))
    ]
    assert len(project_slides) == 1
    entries = _flatten_shapes(project_slides[0].shapes)
    for key in ("current", "issues", "next"):
        assert _content_text(_section_from_entries(entries, key)) == _content_text(values[key])
    if source_format == "pptx":
        names = {shape.name: shape for shape in project_slides[0].shapes}
        assert names["legacy:issues:body"].left > names["legacy:current:body"].left
        assert names["legacy:next:body"].top > names["legacy:current:body"].top


def _project_output_slides(presentation, project):
    return [
        slide for slide in presentation.slides
        if any(project["title"] in entry["text"] and "汇报人" in entry["text"]
               for entry in _flatten_shapes(slide.shapes))
    ]


def _set_compact_body_geometry(body, height: float) -> None:
    body.width = Pt(300)
    body.height = Pt(height)
    body.text_frame.margin_left = body.text_frame.margin_right = Pt(8)
    body.text_frame.margin_top = body.text_frame.margin_bottom = Pt(6)


def _assert_compact_body_font(body, font_size: float) -> None:
    line_height = Centipoints(round(2200 * font_size / 14))
    size_value = str(int(font_size * 100))
    for paragraph in body.text_frame.paragraphs:
        assert paragraph.line_spacing == line_height
        assert paragraph.space_before == paragraph.space_after == Pt(0)
        assert paragraph.font.size == Pt(font_size)
        end_properties = paragraph._p.find(pptx_qn("a:endParaRPr"))
        assert end_properties is not None and end_properties.get("sz") == size_value
        assert all(run.font.size == Pt(font_size) for run in paragraph.runs)


@pytest.mark.parametrize(("height", "expected_font_size"), [
    (83, 14),
    (78, 13),
    (74, 12),
    (69, 11),
    (65.5, 10.5),
    (64, 10),
    (61, 9.5),
    (58.42, 9),
])
def test_pptx_chooses_largest_fitting_font_before_reflow(
    tmp_path: Path, height: float, expected_font_size: float,
) -> None:
    current = "1. 完成接口联调。\n2. 完成数据核对。\n3. 完成验收交付。"

    def style_body(key, body):
        if key == "current":
            _set_compact_body_geometry(body, height)
            # The source's larger defaults must not override the selected size.
            for paragraph in body.text_frame.paragraphs:
                paragraph.font.size = Pt(28)
                end_properties = OxmlElement("a:endParaRPr")
                end_properties.set("sz", "2800")
                paragraph._p.append(end_properties)

    output, project, result = _build_flow_fixture(tmp_path, "pptx", [{
        "current": current, "issues": "无", "next": "1. 安排下周上线。",
    }], style_pptx_body=style_body)

    assert not _flow_pages(output)
    slides = _project_output_slides(output, project)
    assert len(slides) == 1
    names = {shape.name: shape for shape in slides[0].shapes}
    body = names["legacy:current:body"]
    assert body.width == Pt(300) and body.height == Pt(height)
    assert body.text_frame.margin_top == body.text_frame.margin_bottom == Pt(6)
    assert body.text_frame.margin_left == body.text_frame.margin_right == Pt(8)
    assert _content_text(body.text) == _content_text(current)
    _assert_compact_body_font(body, expected_font_size)
    _assert_compact_body_font(names["legacy:next:body"], 14)
    _assert_compact_body_font(names["legacy:issues:body"], 14)
    assert not any(issue["code"] == "generated_content_mismatch" for issue in result["qa"]["issues"])


def test_compact_body_boxes_fit_independently_and_count_soft_breaks(tmp_path: Path) -> None:
    current = "1. 完成开发。\v完成联调。\v完成验收。"
    next_week = "1. 核对数据。\n2. 回归接口。\n3. 安排上线。"

    def style_body(key, body):
        if key in {"current", "next"}:
            _set_compact_body_geometry(body, 78 if key == "current" else 65.5)

    output, project, _ = _build_flow_fixture(tmp_path, "pptx", [{
        "current": current, "issues": "无", "next": next_week,
    }], style_pptx_body=style_body)

    assert not _flow_pages(output)
    slides = _project_output_slides(output, project)
    assert len(slides) == 1
    names = {shape.name: shape for shape in slides[0].shapes}
    _assert_compact_body_font(names["legacy:current:body"], 13)
    _assert_compact_body_font(names["legacy:next:body"], 10.5)
    _assert_compact_body_font(names["legacy:issues:body"], 14)
    assert "\v" in names["legacy:current:body"].text
    assert _content_text(names["legacy:current:body"].text) == _content_text(current)
    assert _content_text(names["legacy:next:body"].text) == _content_text(next_week)


def test_compact_trial_removes_outer_blank_lines_before_choosing_font(tmp_path: Path) -> None:
    current = "已完成 API 联调。\v已完成数据核验。"

    def style_body(key, body):
        if key != "current":
            return
        _set_compact_body_geometry(body, 60)
        body.text = f"\n\n\v{body.text}\v\n\n"
        paragraph = body.text_frame.paragraphs[2]
        paragraph.level = 1
        for run in paragraph.runs:
            run.font.bold = True

    output, project, _ = _build_flow_fixture(tmp_path, "pptx", [{
        "current": current, "issues": "无", "next": "安排上线。",
    }], style_pptx_body=style_body)

    assert not _flow_pages(output)
    slides = _project_output_slides(output, project)
    assert len(slides) == 1
    body = next(shape for shape in slides[0].shapes if shape.name == "legacy:current:body")
    assert body.text == current
    assert body.text_frame.paragraphs[0].level == 1
    assert all(run.font.bold for paragraph in body.text_frame.paragraphs for run in paragraph.runs if run.text)
    _assert_compact_body_font(body, 14)


def test_minimum_font_overflow_on_one_source_page_reflows_whole_project(tmp_path: Path) -> None:
    def style_body(key, body):
        if key == "current":
            # Three 14.14pt lines at 9pt plus padding and safety need 58.42pt.
            _set_compact_body_geometry(body, 58.41)

    source_pages = [
        {"current": "1. 首批接口已交付。", "issues": "1. 首批权限待确认。", "next": "1. 首批安排上线。"},
        {"current": "2. 完成后续开发。\n3. 完成后续联调。\n4. 完成后续验收。",
         "issues": "2. 后续权限待确认。", "next": "2. 后续安排上线。"},
    ]
    output, project, result = _build_flow_fixture(
        tmp_path, "pptx", source_pages, style_pptx_body=style_body,
    )

    pages = _assert_flow_content(output, project)
    assert len(_project_output_slides(output, project)) == len(pages)
    assert not any(issue["code"] == "generated_content_mismatch" for issue in result["qa"]["issues"])


def test_each_fitting_source_page_keeps_layout_despite_large_project_total(tmp_path: Path) -> None:
    def style_body(key, body):
        if key == "current":
            _set_compact_body_geometry(body, 230)

    source_pages = [
        {"current": "\n".join(f"{index}. 第{page}批记录{index}已完成。" for index in range(1, 13)),
         "issues": f"1. 第{page}批权限待确认。", "next": f"1. 第{page}批安排上线。"}
        for page in range(1, 4)
    ]
    output, project, result = _build_flow_fixture(
        tmp_path, "pptx", source_pages, style_pptx_body=style_body,
    )

    assert not _flow_pages(output)
    slides = _project_output_slides(output, project)
    assert len(slides) == len(source_pages)
    for slide, expected in zip(slides, source_pages):
        names = {shape.name: shape for shape in slide.shapes}
        for section in ("current", "issues", "next"):
            body = names[f"legacy:{section}:body"]
            assert _content_text(body.text) == _content_text(expected[section])
            _assert_compact_body_font(body, 11 if section == "current" else 14)
        assert names["legacy:current:body"].height == Pt(230)
        assert names["legacy:issues:body"].left > names["legacy:current:body"].left
        assert names["legacy:next:body"].top > names["legacy:current:body"].top
    assert not any(issue["code"] == "generated_content_mismatch" for issue in result["qa"]["issues"])
