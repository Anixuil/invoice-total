"""Parse reimbursement DOCX files and render printable reimbursement forms."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from functools import lru_cache
from io import BytesIO
from pathlib import Path
import re
import unicodedata

import fitz
from docx import Document
from fontTools.ttLib import TTCollection


LABELS = {
    "报销编号": "reimbursement_number", "报销人": "claimant", "所属部门": "department",
    "工单号": "reimbursement_number", "申请人": "claimant", "申请部门": "department",
    "本次支出金额": "total_amount", "支出类型明细": "expense_type",
    "申请组织": "organization", "费用承担部门": "cost_department", "币别": "currency",
    "填报日期": "report_date", "工作性质": "work_type", "项目类型": "project_type",
    "所属项目": "project", "所属客户": "client", "合同号/立项号": "contract_number",
    "合同号/立项编号": "contract_number", "合同编号/立项号": "contract_number",
    "合同号": "contract_number", "立项号": "contract_number",
    "备注": "notes", "报销总金额": "total_amount",
}
DETAIL_LABELS = {"类型": "type", "用途": "purpose", "类型用途": "purpose", "金额": "amount"}
CN_DIGITS = "零壹贰叁肆伍陆柒捌玖"
CN_UNITS = ("", "拾", "佰", "仟")
CN_BIG_UNITS = ("", "万", "亿")
FIXED_DEPARTMENT = "数据智能部"
TEXT_FONT_NAME = "china-s"
NOTE_FONT_NAME = "notes-cjk"
NOTE_FONT_PATHS = (
    Path("C:/Windows/Fonts/simsun.ttc"),
    Path("C:/Windows/Fonts/NotoSerifSC-VF.ttf"),
    Path("/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc"),
    Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
)


@dataclass
class ReimbursementDetail:
    type: str = ""
    purpose: str = ""
    amount: Decimal = Decimal("0")


@dataclass
class Reimbursement:
    fields: dict[str, str] = field(default_factory=dict)
    details: list[ReimbursementDetail] = field(default_factory=list)

    @property
    def total(self) -> Decimal:
        try:
            return Decimal(self.fields.get("total_amount", "").replace(",", "")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        except InvalidOperation:
            return sum((item.amount for item in self.details), Decimal("0")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def _label(value: str) -> str:
    return re.sub(r"\s+", "", re.sub(r"[：:]$", "", value.strip())).replace("／", "/")


def _money(value: str) -> Decimal:
    try:
        return Decimal(re.sub(r"[^0-9.-]", "", value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except InvalidOperation:
        return Decimal("0")


def _contract_number(value: str) -> str:
    """Treat question-mark placeholders in an unfilled contract field as blank."""
    value = str(value or "").strip()
    return "" if re.fullmatch(r"[?？]*", value) else value


def _looks_like_label(value: str) -> bool:
    """Identify known labels and unrecognised DOCX form labels after empty fields."""
    key = _label(value)
    return (
        key in LABELS
        or key in DETAIL_LABELS
        or key in {"报销明细", "用途明细"}
        or value.strip().endswith((":", "："))
    )


def _parse_reimbursement_lines(lines: list[str]) -> Reimbursement:
    """Parse one reimbursement form from its exported label/value lines."""
    result = Reimbursement()
    detail = None
    details_started = False
    index = 0
    while index < len(lines):
        key = _label(lines[index])
        if key in {"报销明细", "用途明细"}:
            details_started = True
            index += 1
            continue
        if details_started and re.fullmatch(r"(?:报销明细|用途明细)\d+", key):
            if detail and (detail.type or detail.purpose or detail.amount): result.details.append(detail)
            detail = ReimbursementDetail()
            index += 1
            continue
        mapping = DETAIL_LABELS if details_started else LABELS
        target = mapping.get(key)
        if target:
            value = lines[index + 1] if index + 1 < len(lines) else ""
            next_is_label = _looks_like_label(value)
            if next_is_label:
                value = ""
            if details_started:
                detail = detail or ReimbursementDetail()
                if target == "amount": detail.amount = _money(value)
                else: setattr(detail, target, value)
            else:
                result.fields[target] = value
            # Empty fields may be followed by another label. Do not skip it.
            index += 1 if next_is_label else 2
        else:
            index += 1
    if detail and (detail.type or detail.purpose or detail.amount): result.details.append(detail)
    # The expense-request export stores the expense type once above its details.
    for item in result.details:
        if not item.type:
            item.type = result.fields.get("expense_type", "")
    return result


def parse_reimbursement_docx_many(path: str | Path) -> list[Reimbursement]:
    """Read every reimbursement form in a DOCX exported as one continuous document.

    Each occurrence of the standard reimbursement-number label starts a new form.
    This is deliberately detected through ``LABELS`` so supported label punctuation
    variants continue to work.
    """
    lines = [
        paragraph.text.strip().replace("\u00a0", " ")
        for paragraph in Document(str(path)).paragraphs
        if paragraph.text.strip()
    ]
    starts = [
        index for index, line in enumerate(lines)
        if LABELS.get(_label(line)) == "reimbursement_number"
    ]
    if not starts:
        return [_parse_reimbursement_lines(lines)]
    return [
        _parse_reimbursement_lines(lines[start:end])
        for start, end in zip(starts, starts[1:] + [len(lines)])
    ]


def parse_reimbursement_docx(path: str | Path) -> Reimbursement:
    """Read the first reimbursement form for backwards-compatible callers."""
    return parse_reimbursement_docx_many(path)[0]


def amount_to_chinese(value: Decimal) -> str:
    value = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    integer, cents = int(value), int(value * 100) % 100
    if not integer: text = "零"
    else:
        groups = []
        while integer: groups.append(integer % 10000); integer //= 10000
        parts, pending_zero = [], False
        for pos in range(len(groups) - 1, -1, -1):
            group = groups[pos]
            if not group: pending_zero = bool(parts); continue
            if parts and (pending_zero or group < 1000): parts.append("零")
            chars = []
            for unit in range(3, -1, -1):
                digit = group // 10 ** unit % 10
                if digit:
                    if chars and chars[-1] == "零": chars.pop()
                    chars.append(CN_DIGITS[digit] + CN_UNITS[unit])
                elif chars and chars[-1] != "零": chars.append("零")
            if chars[-1] == "零": chars.pop()
            parts.append("".join(chars) + CN_BIG_UNITS[pos]); pending_zero = False
        text = "".join(parts)
    jiao, fen = cents // 10, cents % 10
    return text + "元" + (CN_DIGITS[jiao] + "角" if jiao else "") + (CN_DIGITS[fen] + "分" if fen else "") + ("整" if not cents else "")


def amount_to_template_digits(value: Decimal) -> list[str]:
    """Return the eight handwritten slots on the printed amount-uppercase row.

    The template already prints the units: 拾万仟佰拾元角分. Unused high-order
    slots are crossed out, while a zero inside the actual amount is written as 零.
    """
    value = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if value < 0 or value >= Decimal("1000000"):
        raise ValueError("报销总金额必须在 0 至 999999.99 之间")
    cents = int(value * 100)
    digits = [int(character) for character in f"{cents:08d}"]
    first_used = next((index for index, digit in enumerate(digits[:6]) if digit), 5)
    result = []
    for index, digit in enumerate(digits):
        if index < first_used:
            result.append("×")
        else:
            result.append(CN_DIGITS[digit])
    return result


def _text(
    page,
    rect,
    text,
    color=(.12, .12, .12),
    align=fitz.TEXT_ALIGN_CENTER,
    fontname=TEXT_FONT_NAME,
    fontsize=10,
    preserve_full_text=False,
):
    """Keep required content readable and complete within its cell."""
    text = str(text or "")
    if preserve_full_text:
        if not text:
            return
        # Fill each line by glyph width; textbox word wrapping can move a whole
        # Chinese phrase to the next line and leave most of the first line empty.
        # Built-in china-s drawing uses full-width Latin advances, unlike
        # Font.text_length. Embed the same font used for measuring mixed text.
        fontname, font = _notes_font(page)
        if fontname == TEXT_FONT_NAME:
            fontname = "expense-cjk"
            page.insert_font(fontname=fontname, fontbuffer=font.buffer)
        lines = []
        for paragraph in text.split("\n"):
            line = ""
            for char in paragraph:
                if font.text_length(char, fontsize=fontsize) > rect.width:
                    raise ValueError("报销明细单元格宽度不足以显示正常字号")
                if line and font.text_length(line + char, fontsize=fontsize) > rect.width:
                    lines.append(line)
                    line = ""
                line += char
            lines.append(line)
        line_height = fontsize * 1.15
        baseline = rect.y0 + fontsize * font.ascender
        if baseline + (len(lines) - 1) * line_height - fontsize * font.descender > rect.y1:
            raise ValueError("报销明细用途过长，正常字号换行后仍超出单元格，请拆分明细或精简用途内容")
        for index, line in enumerate(lines):
            width = font.text_length(line, fontsize=fontsize)
            x = rect.x0
            if align == fitz.TEXT_ALIGN_CENTER:
                x += (rect.width - width) / 2
            elif align == fitz.TEXT_ALIGN_RIGHT:
                x += rect.width - width
            page.insert_text(
                (x, baseline + index * line_height), line,
                fontsize=fontsize, fontname=fontname, color=color,
            )
        return
    truncated = text
    while truncated:
        candidate = truncated if truncated == text else truncated + "…"
        if page.insert_textbox(
            rect,
            candidate,
            fontsize=fontsize,
            fontname=fontname,
            color=color,
            align=align,
        ) >= 0:
            return
        truncated = truncated[:-1].rstrip()


@lru_cache(maxsize=4)
def _simplified_chinese_font(path: Path) -> bytes:
    """Extract the SC face instead of using a collection's default locale."""
    collection = TTCollection(str(path), lazy=True)
    try:
        for font in collection.fonts:
            postscript_name = font["name"].getDebugName(6) or ""
            if "cjksc" in postscript_name.lower():
                buffer = BytesIO()
                font.save(buffer)
                return buffer.getvalue()
        raise ValueError("Noto 字体集合中未找到简体中文 SC 字体")
    finally:
        collection.close()


def _notes_font(page) -> tuple[str, fitz.Font]:
    """Register a Song-style mixed CJK font so Latin words are not spaced out."""
    for path in NOTE_FONT_PATHS:
        if path.is_file():
            if path.suffix.lower() == ".ttc" and path.name.startswith("Noto"):
                font_data = _simplified_chinese_font(path)
                page.insert_font(fontname=NOTE_FONT_NAME, fontbuffer=font_data)
                return NOTE_FONT_NAME, fitz.Font(fontbuffer=font_data)
            page.insert_font(fontname=NOTE_FONT_NAME, fontfile=str(path))
            return NOTE_FONT_NAME, fitz.Font(fontfile=str(path))
    return TEXT_FONT_NAME, fitz.Font(fontname=TEXT_FONT_NAME)


def _notes_text(page, rect, text) -> None:
    """Write a readable Song-style note that wraps within its template cell."""
    text = str(text or "").replace("\n", " ").strip()
    if not text:
        return
    fontname, _ = _notes_font(page)
    _text(page, rect, text, align=fitz.TEXT_ALIGN_LEFT, fontname=fontname)


def render_reimbursement_pdf(
    reimbursement: Reimbursement,
    output_path: str | Path,
    template_path: str | Path | None = None,
    generated_at: datetime | None = None,
) -> Path:
    """Fill the supplied template without redrawing or changing its artwork."""
    output_path = Path(output_path)
    template_path = Path(template_path or Path(__file__).parent / "templates" / "reimbursement_template.pdf")
    if not template_path.is_file():
        raise ValueError("未找到报销单空白模板")

    template = fitz.open(template_path)
    if not template.page_count:
        raise ValueError("报销单空白模板没有页面")
    output = fitz.open()
    details = reimbursement.details
    detail_pages = [details[index:index + 4] for index in range(0, len(details), 4)]
    total = reimbursement.total
    generated_at = generated_at or datetime.now()

    for number, page_details in enumerate(detail_pages, start=1):
        # insert_pdf preserves every original template drawing and image unchanged.
        output.insert_pdf(template, from_page=0, to_page=0)
        page = output[-1]
        _text(page, fitz.Rect(135, 119, 285, 136), reimbursement.fields.get("department", ""))
        _text(page, fitz.Rect(510, 23, 700, 41), reimbursement.fields.get("reimbursement_number", ""), fontname="tiro")
        contract_rect = fitz.Rect(510, 42, 700, 60)
        # The supplied template includes placeholder glyphs in this cell.
        page.draw_rect(contract_rect, color=None, fill=(1, 1, 1), overlay=True)
        _text(page, contract_rect, _contract_number(reimbursement.fields.get("contract_number", "")), fontname="tiro")

        _text(page, fitz.Rect(310, 119, 350, 136), str(generated_at.year), fontname="tiro")
        _text(page, fitz.Rect(375, 119, 392, 136), str(generated_at.month), fontname="tiro")
        _text(page, fitz.Rect(410, 119, 432, 136), str(generated_at.day), fontname="tiro")

        _notes_text(page, fitz.Rect(438, 143, 642, 317), reimbursement.fields.get("notes", ""))
        for row, detail in enumerate(page_details):
            top, bottom = 184 + row * 30, 202 + row * 30
            expense = f"{detail.type}（{detail.purpose}）" if detail.type and detail.purpose else detail.type or detail.purpose
            _text(
                page, fitz.Rect(115, top - 5, 307, bottom + 5), expense,
                align=fitz.TEXT_ALIGN_LEFT, preserve_full_text=True,
            )
            _text(page, fitz.Rect(315, top, 397, bottom), f"{detail.amount:.2f}", fontname="tiro")
        _text(page, fitz.Rect(315, 304, 397, 323), f"{total:.2f}", fontname="tiro")
        for x, digit in zip((174, 203, 232, 261, 290, 319, 348, 377), amount_to_template_digits(total)):
            _text(page, fitz.Rect(x, 338, x + 10, 358), digit)
        _text(page, fitz.Rect(565, 366, 627, 386), reimbursement.fields.get("claimant", ""))

    template.close()
    output.set_metadata({"title": "费用报销单", "author": "本地报销单生成工具"})
    # MuPDF's native subsetter can leave Debian's Noto CJK OpenType font intact.
    # The FontTools fallback also handles that font while preserving mixed text.
    output.subset_fonts(fallback=True)
    output.save(str(output_path), garbage=4, deflate=True)
    output.close()
    return output_path


def validate_reimbursement_pdf(
    reimbursement: Reimbursement,
    pdf_path: str | Path,
    generated_at: datetime,
) -> dict[str, object]:
    """Independently verify that a rendered PDF contains every required value."""
    expected_pages = max(1, (len(reimbursement.details) + 3) // 4)
    checks: list[str] = []
    errors: list[str] = []

    def check(name: str, condition: bool) -> None:
        (checks if condition else errors).append(name)

    with fitz.open(str(pdf_path)) as document:
        check("模板尺寸", all(
            abs(page.rect.width - 728.52) < 0.02 and abs(page.rect.height - 515.88) < 0.02
            for page in document
        ))
        check("页数", document.page_count == expected_pages)
        text = "".join(page.get_text("text") for page in document)
        usage_texts = []
        for index in range(len(reimbursement.details)):
            page_index, row = divmod(index, 4)
            if page_index >= document.page_count:
                usage_texts.append("（缺少页面）")
                continue
            top = 184 + row * 30
            usage_texts.append(document[page_index].get_text(
                "text", clip=fitz.Rect(113, top - 6, 310, top + 24),
            ).strip())

    compact_text = re.sub(r"\s+", "", text)

    def normalized_usage(value: str) -> str:
        # PDF fonts may map full-width punctuation to compatibility forms.
        return re.sub(r"\s+", "", unicodedata.normalize("NFKC", value))

    def contains(name: str, value: str, required: bool = True) -> None:
        if not value:
            if required:
                errors.append(name)
            return
        # Optional informational fields must not block generation when a
        # platform-specific PDF font/text extractor cannot recover them.
        # Required fields below still fail closed as before.
        if required:
            check(name, re.sub(r"\s+", "", value) in compact_text)
        elif re.sub(r"\s+", "", value) in compact_text:
            checks.append(name)

    contains("报销部门", reimbursement.fields.get("department", ""))
    contains("报销编号", reimbursement.fields.get("reimbursement_number", ""))
    # Contract/project numbers are informational only and must never block generation.
    contains("生成年份", str(generated_at.year))
    contains("生成月份", str(generated_at.month))
    contains("生成日期", str(generated_at.day))
    contains("备注", reimbursement.fields.get("notes", ""), required=False)
    contains("领款人", reimbursement.fields.get("claimant", ""))
    contains("合计", f"{reimbursement.total:.2f}")
    contains("金额大写", "".join(amount_to_template_digits(reimbursement.total)))
    for index, detail in enumerate(reimbursement.details, start=1):
        usage = f"{detail.type}（{detail.purpose}）" if detail.type and detail.purpose else detail.type or detail.purpose
        # Compare the complete cell, not the page-wide extraction where other
        # columns can interrupt a wrapped description's reading order.
        expected_usage = normalized_usage(usage)
        check(
            f"明细{index}用途",
            bool(expected_usage)
            and expected_usage == normalized_usage(usage_texts[index - 1]),
        )
        if f"明细{index}用途" in errors:
            # Show the evidence needed to distinguish font mapping failures
            # from missing content without weakening the validation.
            errors[-1] = (
                f"明细{index}用途（原文：{usage!r}；"
                f"PDF 单元格提取文字：{usage_texts[index - 1]!r}）"
            )
        contains(f"明细{index}金额", f"{detail.amount:.2f}")

    return {"ok": not errors, "checks": checks, "errors": errors}
