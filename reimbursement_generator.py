"""Parse reimbursement DOCX files and render printable reimbursement forms."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import re

import fitz
from docx import Document


LABELS = {
    "报销编号": "reimbursement_number", "报销人": "claimant", "所属部门": "department",
    "申请组织": "organization", "费用承担部门": "cost_department", "币别": "currency",
    "填报日期": "report_date", "工作性质": "work_type", "项目类型": "project_type",
    "所属项目": "project", "所属客户": "client", "合同号/立项号": "contract_number",
    "合同号/立项编号": "contract_number", "合同编号/立项号": "contract_number",
    "合同号": "contract_number", "立项号": "contract_number",
    "备注": "notes", "报销总金额": "total_amount",
}
DETAIL_LABELS = {"类型": "type", "用途": "purpose", "金额": "amount"}
CN_DIGITS = "零壹贰叁肆伍陆柒捌玖"
CN_UNITS = ("", "拾", "佰", "仟")
CN_BIG_UNITS = ("", "万", "亿")
FIXED_DEPARTMENT = "数据智能部"
TEXT_FONT_NAME = "china-s"


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


def parse_reimbursement_docx(path: str | Path) -> Reimbursement:
    """Read the label/value paragraph export used by the reimbursement platform."""
    lines = [paragraph.text.strip().replace("\u00a0", " ") for paragraph in Document(str(path)).paragraphs if paragraph.text.strip()]
    result = Reimbursement()
    detail = None
    details_started = False
    index = 0
    while index < len(lines):
        key = _label(lines[index])
        if key == "报销明细":
            details_started = True
            index += 1
            continue
        if details_started and re.fullmatch(r"报销明细\d+", key):
            if detail and (detail.type or detail.purpose or detail.amount): result.details.append(detail)
            detail = ReimbursementDetail()
            index += 1
            continue
        mapping = DETAIL_LABELS if details_started else LABELS
        target = mapping.get(key)
        if target:
            value = lines[index + 1] if index + 1 < len(lines) else ""
            next_is_label = _label(value) in LABELS or _label(value) in DETAIL_LABELS or _label(value) == "报销明细"
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
    return result


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
):
    """Write 10pt Song text without allowing content to escape its template cell."""
    text = str(text or "")
    truncated = text
    while truncated:
        candidate = truncated if truncated == text else truncated + "…"
        if page.insert_textbox(
            rect,
            candidate,
            fontsize=10,
            fontname=fontname,
            color=color,
            align=align,
        ) >= 0:
            return
        truncated = truncated[:-1].rstrip()


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
        _text(page, fitz.Rect(510, 42, 700, 60), reimbursement.fields.get("contract_number", ""), fontname="tiro")

        _text(page, fitz.Rect(310, 119, 350, 136), str(generated_at.year), fontname="tiro")
        _text(page, fitz.Rect(375, 119, 392, 136), str(generated_at.month), fontname="tiro")
        _text(page, fitz.Rect(410, 119, 432, 136), str(generated_at.day), fontname="tiro")

        _text(
            page,
            fitz.Rect(438, 143, 642, 317),
            reimbursement.fields.get("notes", ""),
            align=fitz.TEXT_ALIGN_LEFT,
        )
        for row, detail in enumerate(page_details):
            top, bottom = 184 + row * 30, 202 + row * 30
            expense = f"{detail.type}（{detail.purpose}）" if detail.type and detail.purpose else detail.type or detail.purpose
            _text(page, fitz.Rect(115, top, 310, bottom), expense, align=fitz.TEXT_ALIGN_LEFT)
            _text(page, fitz.Rect(315, top, 397, bottom), f"{detail.amount:.2f}", fontname="tiro")
        _text(page, fitz.Rect(315, 304, 397, 323), f"{total:.2f}", fontname="tiro")
        for x, digit in zip((174, 203, 232, 261, 290, 319, 348, 377), amount_to_template_digits(total)):
            _text(page, fitz.Rect(x, 338, x + 10, 358), digit)
        _text(page, fitz.Rect(565, 366, 627, 386), reimbursement.fields.get("claimant", ""))

    template.close()
    output.set_metadata({"title": "费用报销单", "author": "本地报销单生成工具"})
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

    compact_text = re.sub(r"\s+", "", text)

    def contains(name: str, value: str, required: bool = True) -> None:
        if not value:
            if required:
                errors.append(name)
            return
        check(name, re.sub(r"\s+", "", value) in compact_text)

    contains("报销部门", reimbursement.fields.get("department", ""))
    contains("报销编号", reimbursement.fields.get("reimbursement_number", ""))
    contract_number = reimbursement.fields.get("contract_number", "")
    if contract_number:
        contains("合同号/立项号", contract_number)
    contains("生成年份", str(generated_at.year))
    contains("生成月份", str(generated_at.month))
    contains("生成日期", str(generated_at.day))
    contains("备注", reimbursement.fields.get("notes", ""), required=False)
    contains("领款人", reimbursement.fields.get("claimant", ""))
    contains("合计", f"{reimbursement.total:.2f}")
    contains("金额大写", "".join(amount_to_template_digits(reimbursement.total)))
    for index, detail in enumerate(reimbursement.details, start=1):
        usage = f"{detail.type}（{detail.purpose}）" if detail.type and detail.purpose else detail.type or detail.purpose
        contains(f"明细{index}用途", usage)
        contains(f"明细{index}金额", f"{detail.amount:.2f}")

    return {"ok": not errors, "checks": checks, "errors": errors}
