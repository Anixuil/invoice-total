# invoice-total — PDF 发票总金额提取器

多路核算 + 交叉验证 + 置信度投票，确保金额提取准确率最高。

## 用法

```bash
# 单文件
python invoice_total.py 发票.pdf

# 目录批量（自动汇总所有 PDF 的总金额）
python invoice_total.py ./发票目录

# JSON 输出（便于程序化调用）
python invoice_total.py 发票.pdf --json
```

## 依赖

```
pip install pymupdf pdfplumber
```

## 四路独立核算

| 路 | 方法 | 说明 |
|----|------|------|
| ① 文本合计字段 | 正则扫描 | 定位「价税合计」行，提取（小写）金额，支持多张发票 |
| ② 表格结构合计 | pdfplumber | 表格中定位合计行/价税合计行，金额列+税额列求和 |
| ③ 中文大写金额 | 大写转数字 | 壹贰叁…圆角分整 → 数字，与阿拉伯数字互证 |
| ④ 明细行加总 | 表格明细 | 逐行金额+税额累加，还原价税合计 |

## 勾稽验证

- 金额列合计 + 税额列合计 ≈ 价税合计（容差 0.02）
- 通过 → 高置信度加分；未通过 → 置信度降级并提示人工核对

## 置信度规则

- **HIGH**：≥3 路一致，或 2 路一致 + 勾稽通过
- **MEDIUM**：2 路一致
- **LOW**：各路不一致 → 输出各路数值，绝不硬报一个数

## 生成模拟发票测试

```bash
python make_sample_invoice.py sample_invoice.pdf   # 2 张发票，总金额 2416.00
python invoice_total.py sample_invoice.pdf
```
