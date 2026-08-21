#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
invoice_total.py — PDF 发票总金额提取器（多路核算 + 置信度投票）

四路独立核算:
  A. 文本「价税合计」字段提取
  B. 表格结构提取合计
  C. 中文大写金额转数字
  D. 明细行金额+税额加总（还原价税合计）
勾稽验证: 价税合计 ≈ 金额列合计 + 税额列合计
最终: 加权投票 → 总金额 + 置信度等级

用法:
  python invoice_total.py 发票.pdf
  python invoice_total.py ./发票目录 [--json]
"""

import argparse
import json
import re
import sys
from collections import Counter
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from io import BytesIO
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import pymupdf as fitz

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import zxingcpp
    from PIL import Image
except ImportError:
    zxingcpp = None
    Image = None

# ---------------------------------------------------------------- 基础工具

MONEY_RE = re.compile(
    r"[¥￥]\s*([-]?\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|[-]?\d+(?:\.\d{1,2})?)"
)
PLAIN_MONEY_RE = re.compile(
    r"(?<![\d.¥￥])([-]?\d{1,3}(?:,\d{3})+|\d+)\.[0-9]{2}(?![\d.])"
)
MONEY_QUANTUM = Decimal("0.01")
QR_TOTAL_INVOICE_TYPES = {"31", "32"}  # 数电专票、数电普票，二维码金额为价税合计

DIGITS = {
    "零": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9,
    "壹": 1, "贰": 2, "叁": 3, "肆": 4, "伍": 5, "陆": 6, "柒": 7, "捌": 8, "玖": 9,
}
SMALL_UNITS = {"拾": 10, "佰": 100, "仟": 1000, "十": 10, "百": 100, "千": 1000}
BIG_UNITS = {"万": 10000, "亿": 100000000}
CN_DIGIT_CHARS = "零壹贰叁肆伍陆柒捌玖一二三四五六七八九"
CN_INTEGER_CHARS = CN_DIGIT_CHARS + "拾佰仟万亿十百千"
CN_MONEY_RE = re.compile(
    rf"[{CN_INTEGER_CHARS}]+[圆元]"
    rf"(?:[整正]|(?:[{CN_DIGIT_CHARS}]角(?:零?[{CN_DIGIT_CHARS}]分)?|零?[{CN_DIGIT_CHARS}]分))?"
)


def parse_money(s: str):
    """解析 '1,234.56' / '-89.00' 这类字符串为 float。失败返回 None。"""
    if s is None:
        return None
    s = s.replace(",", "").replace("¥", "").replace("￥", "").strip()
    m = re.fullmatch(r"(-?\d+)(\.\d{1,2})?", s)
    if not m:
        return None
    try:
        return float(Decimal(s).quantize(MONEY_QUANTUM, rounding=ROUND_HALF_UP))
    except InvalidOperation:
        return None


def sum_money(values) -> float:
    """使用 Decimal 累加金额，并将对外结果固定为两位小数的 float。"""
    total = sum((Decimal(str(value)) for value in values), Decimal("0"))
    return float(total.quantize(MONEY_QUANTUM, rounding=ROUND_HALF_UP))


def money_difference(left, right) -> Decimal:
    """返回两个金额的绝对差，避免用二进制浮点数判断容差。"""
    return abs(Decimal(str(left)) - Decimal(str(right))).quantize(
        MONEY_QUANTUM, rounding=ROUND_HALF_UP
    )


def parse_invoice_qr(text: str):
    """解析数电发票二维码；仅接受明确以价税合计编码的票种。"""
    parts = [part.strip() for part in text.split(",")]
    if len(parts) < 6 or parts[0] != "01" or parts[1] not in QR_TOTAL_INVOICE_TYPES:
        return None
    if not parts[3] or not re.fullmatch(r"\d{8}", parts[5]):
        return None
    total = parse_money(parts[4])
    if total is None:
        return None
    return {
        "invoice_type": parts[1],
        "invoice_no": parts[3],
        "date": parts[5],
        "total": total,
    }


def _parse_cn_int(part: str) -> int:
    total = 0
    section = 0
    current = 0
    for ch in part:
        if ch in DIGITS:
            current = DIGITS[ch]
        elif ch in SMALL_UNITS:
            if current == 0:
                current = 1  # 「拾」单独出现 = 10
            section += current * SMALL_UNITS[ch]
            current = 0
        elif ch in BIG_UNITS:
            section += current
            current = 0
            if ch == "万":
                total += section * 10000
            else:  # 亿
                total = (total + section) * 100000000
            section = 0
    return total + section + current


def cn_upper_to_number(s: str):
    """中文大写金额 → float。例: 壹仟叁佰伍拾陆元整 → 1356.00。失败返回 None。"""
    s = s.replace("人民币", "").replace("RMB", "").strip()
    m = re.search(
        r"([零一二三四五六七八九壹贰叁肆伍陆柒捌玖拾佰仟万亿十百千]+)[圆元](.*)", s
    )
    if not m:
        return None
    int_part, rest = m.groups()
    value = _parse_cn_int(int_part)
    frac_cents = 0
    jiao = re.search(r"([零壹贰叁肆伍陆柒捌玖一二三四五六七八九])角", rest)
    fen = re.search(r"([零壹贰叁肆伍陆柒捌玖一二三四五六七八九])分", rest)
    if jiao:
        frac_cents += DIGITS[jiao.group(1)] * 10
    if fen:
        frac_cents += DIGITS[fen.group(1)]
    return sum_money([value, Decimal(frac_cents) / 100])


# ---------------------------------------------------------------- 表格分析

def _clean_cell(c):
    return (c or "").strip().replace("\n", " ")


def _analyze_table(tbl):
    """
    返回 (header_idx, cols, total_row_idx, is_summary_row)
    cols: {'amount': 列号, 'tax': 列号, 'total': 列号}
    """
    cols = {}
    header_idx = None
    # 表头：前 3 行内含「金额」「税额」等
    for i, row in enumerate(tbl[:4]):
        cells = [_clean_cell(c) for c in row]
        joined = re.sub(r"\s+", "", "".join(cells))
        if ("金额" in joined or "税额" in joined) and ("项目" in joined or "名称" in joined or "合计" in joined or True):
            header_idx = i
            if "金额" in joined and "税额" in joined and sum(bool(cell) for cell in cells) == 1:
                cols["combined"] = True
            for j, c in enumerate(cells):
                if "价税合计" in c:
                    cols["total"] = j
                elif "税额" in c:
                    cols["tax"] = j
                elif "金额" in c and "税额" not in c and "单价" not in c:
                    cols["amount"] = j
            break
    # 合计行 / 价税合计行
    total_row = None
    is_summary = False
    for i, row in enumerate(tbl):
        if i == header_idx and not cols.get("combined"):
            continue
        cells = [_clean_cell(c) for c in row]
        joined = re.sub(r"\s+", "", "".join(cells))
        if i == header_idx and cols.get("combined"):
            if "合计" in joined and len(_row_numbers(row)) >= 2:
                total_row, is_summary = i, False
                break
            continue
        if "价税合计" in joined:
            total_row, is_summary = i, True
            break
        if ("合计" in joined or "总计" in joined) and any(_cell_money_values(c) for c in cells):
            total_row, is_summary = i, False
            break
    return header_idx, cols, total_row, is_summary


def _cell_money_values(cell):
    value = parse_money(cell)
    if value is not None:
        return [value]
    return [parsed for raw in MONEY_RE.findall(cell) if (parsed := parse_money(raw)) is not None]


def _row_numbers(row, col=None):
    vals = []
    cells = [_clean_cell(c) for c in row] if col is None else [_clean_cell(c) for c in row[col : col + 1]]
    for c in cells:
        vals.extend(_cell_money_values(c))
    return vals


# ---------------------------------------------------------------- 四路核算

class Extractor:
    def __init__(self, pdf_path: str):
        self.path = Path(pdf_path)
        self.text = ""
        self.page_texts = []
        self.image_only_pages = []
        self.qr_results = []
        self.qr_errors = {}
        self.tables = []  # 每页的表格: list[list[list]]
        self.table_meta = []  # 与 tables 一一对应: (header_idx, cols, total_row, is_summary)
        self.table_pages = []  # 与 tables 一一对应的页码

    # ---- 加载 ----
    def load(self) -> None:
        if fitz is None:
            raise RuntimeError("缺少 PyMuPDF，请先 pip install pymupdf")
        with fitz.open(str(self.path)) as doc:
            # 按页面坐标排序，确保「(小写)」与右侧金额保持在同一文本行。
            for page_number, page in enumerate(doc, start=1):
                page_text = page.get_text("text", sort=True)
                self.page_texts.append(page_text)
                if not page_text.strip() and page.get_images(full=True):
                    self.image_only_pages.append(page_number)
                    self._load_page_qr(page, page_number)
            self.text = "\n".join(self.page_texts)
            self.page_count = doc.page_count
        if pdfplumber is not None:
            with pdfplumber.open(str(self.path)) as pdf:
                for page_number, page in enumerate(pdf.pages, start=1):
                    for tbl in page.extract_tables():
                        self.tables.append(tbl)
                        self.table_meta.append(_analyze_table(tbl))
                        self.table_pages.append(page_number)

    def _load_page_qr(self, page, page_number: int) -> None:
        """从无文本页渲染并读取数电发票二维码；失败时保留页面号供人工核对。"""
        if zxingcpp is None or Image is None:
            self.qr_errors[page_number] = "缺少 zxing-cpp 二维码依赖"
            return
        try:
            pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            with Image.open(BytesIO(pixmap.tobytes("png"))) as image:
                barcodes = zxingcpp.read_barcodes(image)
            for barcode in barcodes:
                result = parse_invoice_qr(barcode.text)
                if result is not None:
                    result["page"] = page_number
                    self.qr_results.append(result)
                    return
            self.qr_errors[page_number] = "未找到支持的数电发票二维码"
        except Exception as exc:
            self.qr_errors[page_number] = f"二维码读取失败: {exc}"

    @property
    def is_scanned(self) -> bool:
        return not self.text.strip()

    # ---- 路 A: 文本「价税合计」 ----
    @staticmethod
    def _text_totals(text):
        totals = []
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        for i, line in enumerate(lines):
            if "价税合计" not in line and "税价合计" not in line:
                continue
            found = MONEY_RE.findall(line)
            if found:
                v = parse_money(found[-1])
                if v is not None:
                    totals.append(v)
                    continue
            for nl in lines[i + 1 : i + 3]:
                m = MONEY_RE.findall(nl)
                if m:
                    v = parse_money(m[-1])
                    if v is not None:
                        totals.append(v)
                        break
        return totals

    def route_text_total(self):
        """返回每处『价税合计』行的金额列表。"""
        return self._text_totals(self.text)

    def route_text_total_by_page(self):
        return {page: self._text_totals(text) for page, text in enumerate(self.page_texts, start=1)}

    # ---- 路 B: 表格结构 ----
    def route_table_total(self):
        """每张表提取价税合计。优先『价税合计』行，其次合计行金额+税额。"""
        totals = []
        for tbl, (hdr, cols, total_row, is_summary) in zip(self.tables, self.table_meta):
            if total_row is None:
                continue
            row = tbl[total_row]
            if cols.get("combined"):
                nums = _row_numbers(row)
                if len(nums) >= 2:
                    totals.append(sum_money(nums[-2:]))
                continue
            if is_summary:
                # 价税合计行: 取该行所有数字，取最后一个（小写金额在行尾）
                nums = _row_numbers(row)
                if nums:
                    totals.append(nums[-1])
                continue
            # 普通合计行: 金额列 + 税额列
            a = cols.get("amount")
            t = cols.get("tax")
            vals = []
            if a is not None:
                vals += _row_numbers(row, a)
            if t is not None:
                vals += _row_numbers(row, t)
            if len(vals) >= 2:
                totals.append(sum_money(vals[:2]))
            elif len(vals) == 1:
                totals.append(vals[0])
        return totals

    def route_table_total_by_page(self):
        totals = {}
        for tbl, meta, page in zip(self.tables, self.table_meta, self.table_pages):
            hdr, cols, total_row, is_summary = meta
            if total_row is None:
                continue
            row = tbl[total_row]
            value = None
            if cols.get("combined"):
                nums = _row_numbers(row)
                value = sum_money(nums[-2:]) if len(nums) >= 2 else None
            elif is_summary:
                nums = _row_numbers(row)
                value = nums[-1] if nums else None
            else:
                vals = []
                if cols.get("amount") is not None:
                    vals += _row_numbers(row, cols["amount"])
                if cols.get("tax") is not None:
                    vals += _row_numbers(row, cols["tax"])
                value = sum_money(vals[:2]) if len(vals) >= 2 else vals[0] if vals else None
            if value is not None:
                totals.setdefault(page, []).append(value)
        return totals

    # ---- 路 C: 中文大写 ----
    def route_cn_upper(self):
        vals = []
        for m in CN_MONEY_RE.finditer(self.text):
            v = cn_upper_to_number(m.group())
            if v is not None and 0 <= v < 1e12:
                vals.append(round(v, 2))
        return vals

    def route_cn_upper_by_page(self):
        by_page = {}
        for page, text in enumerate(self.page_texts, start=1):
            values = []
            for match in CN_MONEY_RE.finditer(text):
                value = cn_upper_to_number(match.group())
                if value is not None:
                    values.append(value)
            by_page[page] = values
        return by_page

    def route_qr_total(self):
        """返回无文本图片页中由数电发票二维码得到的价税合计。"""
        return [result["total"] for result in self.qr_results]

    # ---- 路 D: 明细加总 ----
    def route_items_sum(self):
        """每张表: 明细金额列之和 + 税额列之和 = 价税合计。无税额列则只返回金额和（不参与投票）。"""
        out = []  # (价税合计 or None, 金额和, 税额和, 有税额列?)
        for tbl, (hdr, cols, total_row, is_summary) in zip(self.tables, self.table_meta):
            if hdr is None:
                out.append((None, None, None, False))
                continue
            a = cols.get("amount")
            t = cols.get("tax")
            if a is None:
                out.append((None, None, None, False))
                continue
            amount_values = []
            tax_values = []
            for i, row in enumerate(tbl):
                if i <= hdr or i == total_row:
                    continue
                cells = [_clean_cell(c) for c in row]
                if "合计" in "".join(cells) or "总计" in "".join(cells):
                    continue
                va = _row_numbers(row, a)
                if va:
                    amount_values.append(va[0])
                if t is not None:
                    vt = _row_numbers(row, t)
                    if vt:
                        tax_values.append(vt[0])
            amt_sum = sum_money(amount_values)
            tax_sum = sum_money(tax_values)
            if t is not None:
                out.append((sum_money([amt_sum, tax_sum]), amt_sum, tax_sum, True))
            else:
                out.append((None, amt_sum, None, False))
        return out

    def route_items_sum_by_page(self):
        totals = {}
        for tbl, (hdr, cols, total_row, _), page in zip(self.tables, self.table_meta, self.table_pages):
            if hdr is None or cols.get("amount") is None or cols.get("tax") is None:
                continue
            amount_values, tax_values = [], []
            for i, row in enumerate(tbl):
                if i <= hdr or i == total_row:
                    continue
                cells = [_clean_cell(c) for c in row]
                if "合计" in "".join(cells) or "总计" in "".join(cells):
                    continue
                va = _row_numbers(row, cols["amount"])
                vt = _row_numbers(row, cols["tax"])
                if va:
                    amount_values.append(va[0])
                if vt:
                    tax_values.append(vt[0])
            if amount_values or tax_values:
                totals.setdefault(page, []).append(sum_money([sum_money(amount_values), sum_money(tax_values)]))
        return totals

    # ---- 勾稽 ----
    def reconcile(self, total_candidate, detail=""):
        """金额列合计 + 税额列合计 是否 ≈ 价税合计；不可用时 passed 为 None。"""
        amount_values = []
        tax_values = []
        for tbl, (hdr, cols, total_row, is_summary) in zip(self.tables, self.table_meta):
            if hdr is None:
                continue
            a, t = cols.get("amount"), cols.get("tax")
            if a is not None:
                for i, row in enumerate(tbl):
                    if i == hdr or i == total_row:
                        continue
                    cells = [_clean_cell(c) for c in row]
                    if "合计" in "".join(cells) or "总计" in "".join(cells):
                        continue
                    va = _row_numbers(row, a)
                    if va:
                        amount_values.append(va[0])
            if t is not None:
                for i, row in enumerate(tbl):
                    if i == hdr or i == total_row:
                        continue
                    vt = _row_numbers(row, t)
                    if vt:
                        tax_values.append(vt[0])
        amount_sum = sum_money(amount_values)
        tax_sum = sum_money(tax_values)
        if total_candidate is None:
            return None, "无合计候选值，无法勾稽"
        if not amount_values or not tax_values:
            return None, "未识别到可用于勾稽的金额列和税额列"
        detail_total = sum_money([amount_sum, tax_sum])
        diff = money_difference(total_candidate, detail_total)
        passed = diff <= Decimal("0.02")
        desc = f"金额列合计 {amount_sum:.2f} + 税额列合计 {tax_sum:.2f} = {detail_total:.2f} vs 候选 {total_candidate:.2f}，差 {diff:.2f}"
        return passed, ("✅ " if passed else "❌ ") + desc

    def reconcile_by_page(self, page_amounts):
        """按页勾稽；图片页或缺少金额/税额列的页面不强行参与。"""
        page_amount_values = {}
        page_tax_values = {}
        for tbl, (hdr, cols, total_row, _), page in zip(self.tables, self.table_meta, self.table_pages):
            if hdr is None or total_row is None:
                continue
            if page in self.image_only_pages:
                continue
            if cols.get("combined"):
                summary_values = _row_numbers(tbl[total_row])
                if len(summary_values) >= 2:
                    page_amount_values.setdefault(page, []).append(summary_values[-2])
                    page_tax_values.setdefault(page, []).append(summary_values[-1])
                continue
            if cols.get("amount") is None or cols.get("tax") is None:
                continue
            for i, row in enumerate(tbl):
                if i <= hdr or i == total_row:
                    continue
                cells = [_clean_cell(c) for c in row]
                if "合计" in "".join(cells) or "总计" in "".join(cells):
                    continue
                amount_values = _row_numbers(row, cols["amount"])
                tax_values = _row_numbers(row, cols["tax"])
                if amount_values:
                    page_amount_values.setdefault(page, []).append(amount_values[0])
                if tax_values:
                    page_tax_values.setdefault(page, []).append(tax_values[0])

        candidates = {item["page"]: item["amount"] for item in page_amounts if item.get("amount") is not None}
        checked_pages = []
        passed_pages = []
        failed_pages = []
        for page in sorted(set(page_amount_values) & set(page_tax_values) & set(candidates)):
            amount_sum = sum_money(page_amount_values[page])
            tax_sum = sum_money(page_tax_values[page])
            detail_total = sum_money([amount_sum, tax_sum])
            diff = money_difference(candidates[page], detail_total)
            checked_pages.append(page)
            (passed_pages if diff <= Decimal("0.02") else failed_pages).append(page)

        skipped_pages = [
            page for page in range(1, getattr(self, "page_count", 0) + 1)
            if page not in checked_pages
        ]
        if not checked_pages:
            return {
                "available": False,
                "passed": False,
                "partial": bool(skipped_pages),
                "checked_pages": [],
                "passed_pages": [],
                "failed_pages": [],
                "skipped_pages": skipped_pages,
                "desc": "未识别到同时包含金额列和税额列的可勾稽页面",
            }

        passed = not failed_pages
        skipped_desc = f"；另有 {len(skipped_pages)} 页未纳入" if skipped_pages else ""
        detail = (
            f"已检查 {len(checked_pages)} 页：{len(passed_pages)} 页通过，"
            f"{len(failed_pages)} 页未通过{skipped_desc}"
        )
        return {
            "available": True,
            "passed": passed,
            "partial": bool(skipped_pages),
            "checked_pages": checked_pages,
            "passed_pages": passed_pages,
            "failed_pages": failed_pages,
            "skipped_pages": skipped_pages,
            "desc": ("✅ " if passed else "❌ ") + detail,
        }

    def page_amounts(self):
        """按页合并各核算路线，供票面预览逐页对照金额。"""
        by_route = {
            "text": self.route_text_total_by_page(),
            "table": self.route_table_total_by_page(),
            "cn": self.route_cn_upper_by_page(),
            "items": self.route_items_sum_by_page(),
            "qr": {},
        }
        for result in self.qr_results:
            by_route["qr"].setdefault(result["page"], []).append(result["total"])

        pages = []
        for page in range(1, getattr(self, "page_count", len(self.page_texts)) + 1):
            candidates = {}
            for route, values_by_page in by_route.items():
                values = values_by_page.get(page, [])
                if values:
                    candidates[route] = sum_money(values)
            votes = Counter(round(value, 2) for value in candidates.values()).most_common()
            amount = votes[0][0] if votes else None
            top_count = votes[0][1] if votes else 0
            if amount is None:
                confidence = "unknown"
            elif top_count >= 3:
                confidence = "high"
            elif top_count >= 2:
                confidence = "medium"
            elif len(candidates) == 1:
                confidence = "single"
            else:
                confidence = "low"
            pages.append({
                "page": page,
                "amount": amount,
                "confidence": confidence,
                "routes": candidates,
            })
        return pages

    # ---- 汇总决策 ----
    def run(self):
        r_qr = self.route_qr_total()
        if self.is_scanned and not r_qr:
            page_amounts = self.page_amounts()
            return {
                "ok": False,
                "error": "文件没有文本层，且未能从图片页二维码读取发票金额。",
                "total": None,
                "confidence": "unknown",
                "image_only_pages": self.image_only_pages,
                "unresolved_image_pages": self.image_only_pages,
                "notices": [],
                "warnings": ["图片发票未能识别金额，请查看票面"],
                "page_amounts": page_amounts,
            }
        r_text = self.route_text_total()
        r_table = self.route_table_total()
        r_cn = self.route_cn_upper()
        r_items = self.route_items_sum()
        page_amounts = self.page_amounts()
        items_totals = [x[0] for x in r_items if x[0] is not None]
        qr_total = sum_money(r_qr) if r_qr else None
        qr_pages = {result["page"] for result in self.qr_results}
        unresolved_image_pages = [
            page for page in self.image_only_pages if page not in qr_pages
        ]

        cand = {
            "text": r_text,
            "table": r_table,
            "cn": r_cn,
            "items": items_totals,
        }
        totals = {
            k: (sum_money(v) if v else None) for k, v in cand.items()
        }

        # 投票（数值 round 2 后统计）
        votes = []
        vote_src = {}
        for k, v in totals.items():
            if v is not None:
                votes.append(round(v, 2))
                vote_src.setdefault(round(v, 2), []).append(k)
        consensus = Counter(votes).most_common() if votes else []

        base_final = None
        if consensus:
            base_final = consensus[0][0]

        if unresolved_image_pages:
            final = None
        elif base_final is not None and qr_total is not None:
            final = sum_money([base_final, qr_total])
        elif base_final is not None:
            final = base_final
        else:
            final = qr_total

        # 按页勾稽；图片页单独走二维码核验，不阻塞文本页的表格校验。
        reconcile_result = self.reconcile_by_page(page_amounts)
        rec_passed = reconcile_result["passed"] if reconcile_result["available"] else None
        rec_desc = reconcile_result["desc"]

        # 置信度
        if unresolved_image_pages:
            conf = "low"
            pages = "、".join(str(page) for page in unresolved_image_pages)
            conf_desc = f"第 {pages} 页为图片且金额无法识别，未输出不完整总额"
        elif r_qr:
            conf = "medium"
            conf_desc = (
                f"{consensus[0][1]}/{len(votes)} 路文本核算一致；另有 "
                f"{len(r_qr)} 张图片发票已通过二维码识别"
                if votes else f"{len(r_qr)} 张图片发票已通过二维码识别"
            )
        elif not votes:
            conf, conf_desc = "low", "四路全部失败"
        else:
            top_count = consensus[0][1]
            n_routes = len(votes)
            if (top_count >= 3 and n_routes >= 3) or (top_count >= 2 and rec_passed is True and n_routes >= 2):
                conf = "high"
            elif top_count >= 2:
                conf = "medium"
            else:
                conf = "low"
            if rec_passed is True:
                rec_status = "勾稽通过"
            elif rec_passed is False:
                rec_status = "勾稽未通过"
            else:
                rec_status = "勾稽不可用"
            conf_desc = f"{top_count}/{n_routes} 路一致，{rec_status}"

        text_invoice_count = max(len(v) for v in cand.values()) if any(cand.values()) else 0
        invoice_count = text_invoice_count + len(self.image_only_pages)
        notices = []
        warnings = []
        if r_qr:
            details = "、".join(
                f"第 {result['page']} 页 ¥{result['total']:,.2f}"
                for result in self.qr_results
            )
            notices.append(f"图片发票已识别：{details}，金额已纳入合计")
        if unresolved_image_pages:
            pages = "、".join(str(page) for page in unresolved_image_pages)
            warnings.append(f"第 {pages} 页图片发票未能识别金额，已暂停汇总")

        return {
            "ok": True,
            "file": str(self.path),
            "pages": getattr(self, "page_count", 0),
            "invoice_count": invoice_count,
            "total": final,
            "confidence": conf,
            "confidence_desc": conf_desc,
            "routes": {
                "text": {"values": r_text, "sum": totals["text"]},
                "table": {"values": r_table, "sum": totals["table"]},
                "cn": {"values": r_cn, "sum": totals["cn"]},
                "items": {"values": items_totals, "sum": totals["items"]},
                "qr": {"values": r_qr, "sum": qr_total},
            },
            "reconcile": reconcile_result,
            "image_only_pages": self.image_only_pages,
            "unresolved_image_pages": unresolved_image_pages,
            "notices": notices,
            "warnings": warnings,
            "page_amounts": page_amounts,
        }


# ---------------------------------------------------------------- 输出

def fmt_money(v):
    return "—" if v is None else f"¥{v:,.2f}"


def print_report(r):
    if not r["ok"]:
        print(f"[!] {r.get('file', '?')}: {r['error']}")
        return
    print("=" * 52)
    print(f"文件      : {r['file']}")
    print(f"检测发票  : {r['invoice_count']} 张")
    print("-" * 52)
    print("各路核算:")
    route_names = [
        ("text", "① 文本合计字段"),
        ("table", "② 表格结构合计"),
        ("cn", "③ 中文大写金额"),
        ("items", "④ 明细行加总"),
    ]
    if "qr" in r["routes"]:
        route_names.append(("qr", "⑤ 图片页二维码"))
    for k, name in route_names:
        vals = r["routes"][k]["values"]
        s = fmt_money(r["routes"][k]["sum"])
        print(f"  {name}: {s}  ({len(vals)} 处) {vals}")
    print("-" * 52)
    print(f"勾稽验证 : {r['reconcile']['desc']}")
    print("-" * 52)
    print(f"最终总额 : {fmt_money(r['total'])}")
    print(f"置信度   : {r['confidence'].upper()}  ({r['confidence_desc']})")
    for notice in r.get("notices", []):
        print(f"  ℹ️  {notice}")
    for warning in r.get("warnings", []):
        print(f"  ⚠️  {warning}")
    if r["confidence"] == "low" and not r.get("warnings"):
        print("  ⚠️  低置信度结果，请人工核对上述数值")
    print("=" * 52)


def main():
    ap = argparse.ArgumentParser(description="PDF 发票总金额提取（多路核算+置信度）")
    ap.add_argument("target", help="PDF 文件或目录")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    args = ap.parse_args()

    target = Path(args.target)
    if not target.exists():
        print(f"[!] 路径不存在: {target}")
        return
    if target.is_dir():
        files = sorted(
            p for p in target.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"
        )
        if not files:
            print("目录中没有 PDF")
            return
        results = []
        for f in files:
            ex = Extractor(f)
            try:
                ex.load()
            except Exception as e:
                print(f"[!] {f}: {e}")
                continue
            r = ex.run()
            results.append(r)
            if not args.json:
                print_report(r)
        if args.json:
            print(json.dumps(results, ensure_ascii=False, indent=2))
        else:
            ok = [r for r in results if r["ok"] and r["total"] is not None]
            if ok:
                grand = sum_money(r["total"] for r in ok)
                print(f"\n>>> 全部 {len(ok)} 个 PDF 发票总金额: ¥{grand:,.2f}")
    else:
        if target.suffix.lower() != ".pdf":
            print(f"[!] 仅支持 PDF 文件: {target}")
            return
        ex = Extractor(target)
        try:
            ex.load()
        except Exception as e:
            print(f"[!] {e}")
            return
        r = ex.run()
        if args.json:
            print(json.dumps(r, ensure_ascii=False, indent=2))
        else:
            print_report(r)


if __name__ == "__main__":
    main()
