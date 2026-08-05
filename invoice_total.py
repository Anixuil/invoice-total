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
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import fitz  # PyMuPDF

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

# ---------------------------------------------------------------- 基础工具

MONEY_RE = re.compile(
    r"[¥￥]\s*([-]?\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|[-]?\d+(?:\.\d{1,2})?)"
)
PLAIN_MONEY_RE = re.compile(
    r"(?<![\d.¥￥])([-]?\d{1,3}(?:,\d{3})+|\d+)\.[0-9]{2}(?![\d.])"
)

DIGITS = {
    "零": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9,
    "壹": 1, "贰": 2, "叁": 3, "肆": 4, "伍": 5, "陆": 6, "柒": 7, "捌": 8, "玖": 9,
}
SMALL_UNITS = {"拾": 10, "佰": 100, "仟": 1000, "十": 10, "百": 100, "千": 1000}
BIG_UNITS = {"万": 10000, "亿": 100000000}
CN_MONEY_RE = re.compile(
    r"[零壹贰叁肆伍陆柒捌玖拾佰仟万亿十百千]+[圆元]"
    r"(?:[零壹贰叁肆伍陆柒捌玖拾佰仟万亿十百千]*[角分整正])?"
)


def parse_money(s: str) -> float:
    """解析 '1,234.56' / '-89.00' 这类字符串为 float。失败返回 None。"""
    if s is None:
        return None
    s = s.replace(",", "").replace("¥", "").replace("￥", "").strip()
    m = re.fullmatch(r"(-?\d+)(\.\d{1,2})?", s)
    if not m:
        return None
    return round(float(s), 2)


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
    frac = 0.0
    jiao = re.search(r"([零壹贰叁肆伍陆柒捌玖一二三四五六七八九])角", rest)
    fen = re.search(r"([零壹贰叁肆伍陆柒捌玖一二三四五六七八九])分", rest)
    if jiao:
        frac += DIGITS[jiao.group(1)] * 0.1
    if fen:
        frac += DIGITS[fen.group(1)] * 0.01
    return round(value + frac, 2)


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
        joined = "".join(cells)
        if ("金额" in joined or "税额" in joined) and ("项目" in joined or "名称" in joined or "合计" in joined or True):
            header_idx = i
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
        if i == header_idx:
            continue
        cells = [_clean_cell(c) for c in row]
        joined = "".join(cells)
        if "价税合计" in joined:
            total_row, is_summary = i, True
            break
        if ("合计" in joined or "总计" in joined) and any(
            parse_money(c) is not None for c in cells
        ):
            total_row, is_summary = i, False
            break
    return header_idx, cols, total_row, is_summary


def _row_numbers(row, col=None):
    vals = []
    cells = [_clean_cell(c) for c in row] if col is None else [_clean_cell(c) for c in row[col : col + 1]]
    for c in cells:
        v = parse_money(c)
        if v is not None:
            vals.append(v)
    return vals


# ---------------------------------------------------------------- 四路核算

class Extractor:
    def __init__(self, pdf_path: str):
        self.path = Path(pdf_path)
        self.text = ""
        self.tables = []  # 每页的表格: list[list[list]]
        self.table_meta = []  # 与 tables 一一对应: (header_idx, cols, total_row, is_summary)

    # ---- 加载 ----
    def load(self) -> None:
        if fitz is None:
            raise RuntimeError("缺少 PyMuPDF，请先 pip install pymupdf")
        with fitz.open(str(self.path)) as doc:
            self.text = "\n".join(page.get_text("text") for page in doc)
            self.page_count = doc.page_count
        if pdfplumber is not None:
            with pdfplumber.open(str(self.path)) as pdf:
                for page in pdf.pages:
                    for tbl in page.extract_tables():
                        self.tables.append(tbl)
                        self.table_meta.append(_analyze_table(tbl))

    @property
    def is_scanned(self) -> bool:
        return not self.text.strip()

    # ---- 路 A: 文本「价税合计」 ----
    def route_text_total(self):
        """返回每处『价税合计』行的金额列表。"""
        totals = []
        lines = [ln.strip() for ln in self.text.splitlines() if ln.strip()]
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

    # ---- 路 B: 表格结构 ----
    def route_table_total(self):
        """每张表提取价税合计。优先『价税合计』行，其次合计行金额+税额。"""
        totals = []
        for tbl, (hdr, cols, total_row, is_summary) in zip(self.tables, self.table_meta):
            if total_row is None:
                continue
            row = tbl[total_row]
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
                totals.append(round(vals[0] + vals[1], 2))
            elif len(vals) == 1:
                totals.append(vals[0])
        return totals

    # ---- 路 C: 中文大写 ----
    def route_cn_upper(self):
        vals = []
        for m in CN_MONEY_RE.finditer(self.text):
            v = cn_upper_to_number(m.group())
            if v is not None and 0 <= v < 1e12:
                vals.append(round(v, 2))
        return vals

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
            amt_sum = 0.0
            tax_sum = 0.0
            for i, row in enumerate(tbl):
                if i <= hdr or i == total_row:
                    continue
                cells = [_clean_cell(c) for c in row]
                if "合计" in "".join(cells) or "总计" in "".join(cells):
                    continue
                va = _row_numbers(row, a)
                if va:
                    amt_sum += va[0]
                if t is not None:
                    vt = _row_numbers(row, t)
                    if vt:
                        tax_sum += vt[0]
            amt_sum = round(amt_sum, 2)
            tax_sum = round(tax_sum, 2)
            if t is not None:
                out.append((round(amt_sum + tax_sum, 2), amt_sum, tax_sum, True))
            else:
                out.append((None, amt_sum, None, False))
        return out

    # ---- 勾稽 ----
    def reconcile(self, total_candidate, detail=""):
        """金额列合计 + 税额列合计 是否 ≈ 价税合计。返回 (passed, desc)。"""
        amount_sum, tax_sum = 0.0, 0.0
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
                        amount_sum += va[0]
            if t is not None:
                for i, row in enumerate(tbl):
                    if i == hdr or i == total_row:
                        continue
                    vt = _row_numbers(row, t)
                    if vt:
                        tax_sum += vt[0]
        amount_sum = round(amount_sum, 2)
        tax_sum = round(tax_sum, 2)
        if total_candidate is None:
            return False, f"无合计候选值，无法勾稽（金额列合计 {amount_sum:.2f} + 税额列合计 {tax_sum:.2f}）"
        diff = abs(total_candidate - (amount_sum + tax_sum))
        passed = diff <= 0.02
        desc = f"金额列合计 {amount_sum:.2f} + 税额列合计 {tax_sum:.2f} = {amount_sum + tax_sum:.2f} vs 候选 {total_candidate:.2f}，差 {diff:.2f}"
        return passed, ("✅ " if passed else "❌ ") + desc

    # ---- 汇总决策 ----
    def run(self):
        if self.is_scanned:
            return {
                "ok": False,
                "error": "疑似扫描件（无文本层），需要 OCR 引擎，当前未安装。",
                "total": None,
                "confidence": "unknown",
            }
        r_text = self.route_text_total()
        r_table = self.route_table_total()
        r_cn = self.route_cn_upper()
        r_items = self.route_items_sum()
        items_totals = [x[0] for x in r_items if x[0] is not None]

        cand = {
            "text": r_text,
            "table": r_table,
            "cn": r_cn,
            "items": items_totals,
        }
        totals = {
            k: (round(sum(v), 2) if v else None) for k, v in cand.items()
        }

        # 投票（数值 round 2 后统计）
        votes = []
        vote_src = {}
        for k, v in totals.items():
            if v is not None:
                votes.append(round(v, 2))
                vote_src.setdefault(round(v, 2), []).append(k)
        consensus = Counter(votes).most_common() if votes else []

        final = None
        if consensus:
            final = consensus[0][0]

        # 勾稽
        rec_passed, rec_desc = self.reconcile(final, "")

        # 置信度
        if not votes:
            conf, conf_desc = "low", "四路全部失败"
        else:
            top_count = consensus[0][1]
            n_routes = len(votes)
            if (top_count >= 3 and n_routes >= 3) or (top_count >= 2 and rec_passed and n_routes >= 2):
                conf = "high"
            elif top_count >= 2:
                conf = "medium"
            else:
                conf = "low"
            conf_desc = f"{top_count}/{n_routes} 路一致" + ("，勾稽通过" if rec_passed else "，勾稽未通过/不可用")

        invoice_count = max(len(v) for v in cand.values()) if any(cand.values()) else 0

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
            },
            "reconcile": {"passed": rec_passed, "desc": rec_desc},
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
    for k, name in (("text", "① 文本合计字段"), ("table", "② 表格结构合计"),
                    ("cn", "③ 中文大写金额"), ("items", "④ 明细行加总")):
        vals = r["routes"][k]["values"]
        s = fmt_money(r["routes"][k]["sum"])
        print(f"  {name}: {s}  ({len(vals)} 处) {vals}")
    print("-" * 52)
    print(f"勾稽验证 : {r['reconcile']['desc']}")
    print("-" * 52)
    print(f"最终总额 : {fmt_money(r['total'])}")
    print(f"置信度   : {r['confidence'].upper()}  ({r['confidence_desc']})")
    if r["confidence"] == "low":
        print("  ⚠️  各路不一致，请人工核对上述数值")
    print("=" * 52)


def main():
    ap = argparse.ArgumentParser(description="PDF 发票总金额提取（多路核算+置信度）")
    ap.add_argument("target", help="PDF 文件或目录")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    args = ap.parse_args()

    target = Path(args.target)
    if target.is_dir():
        files = sorted(target.glob("*.pdf"))
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
                grand = round(sum(r["total"] for r in ok), 2)
                print(f"\n>>> 全部 {len(ok)} 个 PDF 发票总金额: ¥{grand:,.2f}")
    else:
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
