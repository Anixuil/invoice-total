from decimal import Decimal
from datetime import datetime
from pathlib import Path
import tempfile
import unittest
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

from docx import Document
import fitz

from reimbursement_generator import (
    Reimbursement,
    ReimbursementDetail,
    _contract_number,
    amount_to_chinese,
    amount_to_template_digits,
    parse_reimbursement_docx,
    parse_reimbursement_docx_many,
    render_reimbursement_pdf,
    validate_reimbursement_pdf,
)


class ReimbursementGeneratorTests(unittest.TestCase):
    def test_question_mark_contract_placeholder_is_blank(self):
        self.assertEqual(_contract_number("???"), "")
        self.assertEqual(_contract_number("？？？"), "")
        self.assertEqual(_contract_number("PRDP2024063"), "PRDP2024063")

    def test_contract_number_does_not_block_validation(self):
        reimbursement = Reimbursement(
            fields={
                "department": "department",
                "reimbursement_number": "number",
                "claimant": "claimant",
                "contract_number": "unrendered-contract-number",
                "total_amount": "10.00",
            },
            details=[ReimbursementDetail(type="expense", amount=Decimal("10.00"))],
        )
        page = MagicMock()
        page.rect = SimpleNamespace(width=728.52, height=515.88)
        page.get_text.return_value = (
            "department number 2026 8 14 claimant 10.00 expense"
            + "".join(amount_to_template_digits(Decimal("10.00")))
        )
        document = MagicMock()
        document.page_count = 1
        document.__enter__.return_value = document
        document.__iter__.side_effect = lambda: iter([page])

        with patch("reimbursement_generator.fitz.open", return_value=document):
            validation = validate_reimbursement_pdf(reimbursement, "unused.pdf", datetime(2026, 8, 14))

        self.assertTrue(validation["ok"], validation["errors"])

    def test_amount_to_chinese(self):
        self.assertEqual(amount_to_chinese(Decimal("321.15")), "叁佰贰拾壹元壹角伍分")
        self.assertEqual(amount_to_chinese(Decimal("100.00")), "壹佰元整")

    def test_amount_to_template_digits(self):
        self.assertEqual(
            amount_to_template_digits(Decimal("321.15")),
            ["×", "×", "×", "叁", "贰", "壹", "壹", "伍"],
        )

    def test_contract_number_label_accepts_spacing_and_variants(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "contract.docx"
            document = Document()
            for value in ("合同号 / 立项号：", "PRDP2024063"):
                document.add_paragraph(value)
            document.save(source)
            reimbursement = parse_reimbursement_docx(source)
            self.assertEqual(reimbursement.fields["contract_number"], "PRDP2024063")

    def test_contract_number_after_empty_customer_field_is_not_skipped(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "consecutive-labels.docx"
            document = Document()
            for value in ("所属客户:", "合同号/立项号:", "PRDP2024063"):
                document.add_paragraph(value)
            document.save(source)
            reimbursement = parse_reimbursement_docx(source)
            self.assertEqual(reimbursement.fields["client"], "")
            self.assertEqual(reimbursement.fields["contract_number"], "PRDP2024063")

    def test_contract_number_is_empty_before_an_unrecognised_form_label(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "empty-contract.docx"
            document = Document()
            for value in ("合同号/立项号:", "产品线:", "备注:", "测试备注"):
                document.add_paragraph(value)
            document.save(source)

            reimbursement = parse_reimbursement_docx(source)

            self.assertEqual(reimbursement.fields["contract_number"], "")
            self.assertEqual(reimbursement.fields["notes"], "测试备注")
        self.assertEqual(
            amount_to_template_digits(Decimal("101.00")),
            ["×", "×", "×", "壹", "零", "壹", "零", "零"],
        )

    def test_parses_multiple_reimbursements_in_one_docx(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "multiple.docx"
            document = Document()
            for number, claimant, amount in (("TEST-001", "张三", "12.50"), ("TEST-002", "李四", "20.00")):
                for value in (
                    "报销编号:", number, "报销人:", claimant, "所属部门:", "研发部",
                    "报销总金额:", amount, "报销明细", "报销明细1", "类型", "交通费",
                    "用途:", "出行", "金额:", amount,
                ):
                    document.add_paragraph(value)
            document.save(source)

            reimbursements = parse_reimbursement_docx_many(source)

            self.assertEqual(len(reimbursements), 2)
            self.assertEqual(reimbursements[0].fields["reimbursement_number"], "TEST-001")
            self.assertEqual(reimbursements[0].fields["claimant"], "张三")
            self.assertEqual(reimbursements[0].total, Decimal("12.50"))
            self.assertEqual(reimbursements[1].fields["reimbursement_number"], "TEST-002")
            self.assertEqual(reimbursements[1].fields["claimant"], "李四")
            self.assertEqual(reimbursements[1].total, Decimal("20.00"))

    def test_parses_docx_and_renders_multiple_pages(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.docx"
            document = Document()
            for value in ("报销人:", "测试人员", "所属部门:", "研发部", "报销编号:", "TEST-001", "报销总金额:", "50.00", "报销明细"):
                document.add_paragraph(value)
            for index in range(5):
                for value in (f"报销明细{index + 1}", "类型", "交通费", "用途:", f"出行{index + 1}", "金额:", "10.00"):
                    document.add_paragraph(value)
            document.save(source)
            reimbursement = parse_reimbursement_docx(source)
            self.assertEqual(len(reimbursement.details), 5)
            self.assertEqual(reimbursement.total, Decimal("50.00"))
            self.assertEqual(reimbursement.fields.get("contract_number", ""), "")
            output = Path(directory) / "output.pdf"
            generated_at = __import__("datetime").datetime(2026, 8, 14)
            render_reimbursement_pdf(reimbursement, output, generated_at=generated_at)
            validation = validate_reimbursement_pdf(reimbursement, output, generated_at)
            self.assertTrue(validation["ok"], validation["errors"])
            self.assertNotIn("合同号/立项号", validation["errors"])
            with fitz.open(output) as pdf:
                self.assertEqual(pdf.page_count, 2)
                self.assertAlmostEqual(pdf[0].rect.width, 728.52, places=2)
                self.assertAlmostEqual(pdf[0].rect.height, 515.88, places=2)
                self.assertIn("测试人员", pdf[0].get_text())


if __name__ == "__main__":
    unittest.main()
