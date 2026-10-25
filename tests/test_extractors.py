"""Unit tests for invoice extraction and validation."""

import unittest
from pathlib import Path

from main import (
    build_record,
    extract_customer,
    extract_description,
    extract_invoice_number,
    extract_pdf_text,
    extract_total_amount,
    flag_duplicate_invoices,
)


SAMPLE_TEXT = """Invoice Number: INV-1001
Customer: Acme Retail Demo
Description Amount
Amount (INR)
1
PDF invoice extraction and Excel reporting automation
Includes text extraction, field parsing, and report formatting
18,500.00
Subtotal
INR 18,500.00
Tax (sample: 0%)
INR 0.00
Total Amount: INR 18,500.00
"""


class TestInvoiceExtractors(unittest.TestCase):
    def test_invoice_number(self):
        self.assertEqual(extract_invoice_number(SAMPLE_TEXT), "INV-1001")

    def test_customer(self):
        self.assertEqual(extract_customer(SAMPLE_TEXT), "Acme Retail Demo")

    def test_total_amount(self):
        self.assertEqual(extract_total_amount(SAMPLE_TEXT), 18500.0)

    def test_description_excludes_amount_and_totals(self):
        self.assertEqual(
            extract_description(SAMPLE_TEXT),
            "PDF invoice extraction and Excel reporting automation "
            "Includes text extraction, field parsing, and report formatting",
        )

    def test_missing_amount_is_none(self):
        self.assertIsNone(extract_total_amount("Invoice Number: INV-1002"))

    def test_record_is_valid_for_example(self):
        record = build_record(Path("INV-1001.pdf"), SAMPLE_TEXT)
        self.assertEqual(record["status"], "OK")
        self.assertEqual(record["total_amount"], 18500.0)

    def test_duplicate_invoice_numbers_are_flagged(self):
        records = [
            build_record(Path("one.pdf"), SAMPLE_TEXT),
            build_record(Path("two.pdf"), SAMPLE_TEXT),
        ]
        flag_duplicate_invoices(records)
        self.assertTrue(all(r["status"] == "REVIEW" for r in records))
        self.assertTrue(all("Duplicate invoice number; verify manually" in r["review_notes"] for r in records))

    def test_sample_pdf_contains_extractable_text(self):
        project_root = Path(__file__).resolve().parents[1]
        pdf_path = project_root / "examples" / "INV-1001.pdf"
        text = extract_pdf_text(pdf_path)
        self.assertIn("INV-1001", text)
        self.assertIn("Acme Retail Demo", text)


if __name__ == "__main__":
    unittest.main()
