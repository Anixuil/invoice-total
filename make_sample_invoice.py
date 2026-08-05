#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_sample_invoice.py — 生成模拟数电发票 PDF（用于测试 invoice_total.py）

生成 2 张发票:
  发票1: 办公用品 500.00(13%,65.00) + 差旅费 700.00(13%,91.00)
         → 金额 1200.00, 税额 156.00, 价税合计 1356.00, 大写「壹仟叁佰伍拾陆元整」
  发票2: 技术服务 1000.00(6%,60.00)
         → 金额 1000.00, 税额 60.00, 价税合计 1060.00, 大写「壹仟零陆拾元整」
总计: 2416.00
"""

import fitz

INVOICES = [
    {
        "no": "011001900111",
        "date": "2026年08月01日",
        "buyer": "某某科技有限公司",
        "items": [
            ("办公用品", "批", "1", "1", "500.00", "500.00", "13%", "65.00"),
            ("差旅费", "次", "1", "1", "700.00", "700.00", "13%", "91.00"),
        ],
        "amt": "1200.00", "tax": "156.00", "grand": "1356.00",
        "grand_cn": "壹仟叁佰伍拾陆元整",
    },
    {
        "no": "011001900112",
        "date": "2026年08月02日",
        "buyer": "某某科技有限公司",
        "items": [
            ("技术服务费", "项", "1", "1", "1000.00", "1000.00", "6%", "60.00"),
        ],
        "amt": "1000.00", "tax": "60.00", "grand": "1060.00",
        "grand_cn": "壹仟零陆拾元整",
    },
]

HEADERS = ["项目名称", "规格型号", "单位", "数量", "单价", "金额", "税率", "税额"]
COLS_X = [50, 165, 205, 235, 265, 335, 405, 445, 515]


def text_w(s, size=9):
    return fitz.get_text_length(s, fontname="china-s", fontsize=size)


def draw_table(page, y0, row_ys, y1):
    """画表格：外框 + 竖线 + 行分隔线"""
    for x in COLS_X:
        page.draw_line((x, y0), (x, y1), color=(0.6, 0.6, 0.6), width=0.6)
    for yy in [y0] + row_ys + [y1]:
        page.draw_line((COLS_X[0], yy), (COLS_X[-1], yy), color=(0.4, 0.4, 0.4), width=0.8)


def cell_left(page, x_idx, y, s, size=9):
    page.insert_text((COLS_X[x_idx] + 6, y), s, fontname="china-s", fontsize=size)


def make_pdf(path):
    doc = fitz.open()
    for inv in INVOICES:
        page = doc.new_page(width=595, height=842)  # A4
        page.insert_text((200, 60), "电子发票（普通发票）", fontname="china-s", fontsize=16)
        page.insert_text((380, 60), f"发票号码: {inv['no']}", fontname="china-s", fontsize=9)
        page.insert_text((380, 78), f"开票日期: {inv['date']}", fontname="china-s", fontsize=9)
        page.insert_text((50, 110), f"购买方: {inv['buyer']}", fontname="china-s", fontsize=10)
        page.insert_text((50, 128), "纳税人识别号: 91440101MA5XXXXXX", fontname="china-s", fontsize=10)

        n = len(inv["items"])
        row_h = 24
        y0 = 160
        row_ys = [y0 + row_h * (i + 1) for i in range(n + 1)]  # 表头后每行一条线 + 合计线
        y1 = y0 + row_h * (n + 2)
        draw_table(page, y0, row_ys, y1)

        for j, h in enumerate(HEADERS):
            cell_left(page, j, y0 + 16, h)
        for i, row in enumerate(inv["items"]):
            y = y0 + row_h * (i + 1)
            for j, v in enumerate(row):
                cell_left(page, j, y + 16, v)

        y_sum = y0 + row_h * (n + 1)
        cell_left(page, 0, y_sum + 16, "合计")
        cell_left(page, 5, y_sum + 16, inv["amt"])
        cell_left(page, 7, y_sum + 16, inv["tax"])

        y_g = y1 + 40
        page.insert_text((50, y_g), f"价税合计（大写）{inv['grand_cn']}", fontname="china-s", fontsize=10)
        page.insert_text((50, y_g + 20), f"（小写）¥{inv['grand']}", fontname="china-s", fontsize=10)
        page.insert_text((50, y_g + 55), "销售方: 某某软件有限公司", fontname="china-s", fontsize=10)

    doc.save(str(path))
    doc.close()


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "sample_invoice.pdf"
    make_pdf(out)
    print(f"已生成 {out}（2 张发票，总金额应为 2416.00）")
