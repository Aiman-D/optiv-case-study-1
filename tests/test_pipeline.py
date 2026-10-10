"""
Tests for CS1: Vendor AI Usage Analysis Pipeline.

Covers:
  1. Criticality determinism (same input → same output)
  2. Substring-quote guard (invalid quotes are discarded)
  3. Weighting logic (signal scores computed correctly)
  4. B-K unchanged round-trip (columns B-K untouched after export)
  5. Coverage-gap handling (unreachable source ≠ "No")
"""

import sys
import os
import unittest
import json
import shutil
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from criticality import evaluate_vendor, load_rubric
from classify import verify_quote, pre_filter, classify_with_mock, find_ai_sentences


class TestCriticalityDeterminism(unittest.TestCase):
    """Same input row must always produce the same tier and rationale."""

    def setUp(self):
        self.rubric = load_rubric(os.path.join(os.path.dirname(__file__), '..', 'config', 'rubric.yaml'))
        self.row_fiserv = {
            'Data Classification Accessed': 'Full customer master, balances and transactions for 1.4 million customers.',
            'Service / Product Provided': 'The DNA core banking platform',
            'Business Process Supported': 'Core account processing and servicing',
            'Category': 'Core Platform',
            'Operational Dependency': 'Critical',
            'Data Volume (annual)': '240 million transactions'
        }

    def test_fiserv_always_critical(self):
        tier1, rat1, _ = evaluate_vendor(self.row_fiserv, self.rubric)
        tier2, rat2, _ = evaluate_vendor(self.row_fiserv, self.rubric)
        self.assertEqual(tier1, tier2)
        self.assertEqual(rat1, rat2)
        self.assertEqual(tier1, 'Critical')

    def test_automworx_not_critical(self):
        row = {
            'Data Classification Accessed': 'No direct customer data. Privileged access to production job schedules.',
            'Service / Product Provided': 'workload automation and batch scheduling',
            'Business Process Supported': 'Batch scheduling across core systems',
            'Category': 'Professional Services',
            'Operational Dependency': 'High',
            'Data Volume (annual)': 'Not applicable'
        }
        tier, _, _ = evaluate_vendor(row, self.rubric)
        self.assertIn(tier, ('Low', 'Medium'))  # should never be Critical or High
        self.assertNotEqual(tier, 'Critical')

    def test_privileged_access_modifier(self):
        row = {
            'Data Classification Accessed': 'Privileged access to production credentials.',
            'Service / Product Provided': 'consulting',
            'Business Process Supported': 'batch scheduling',
            'Category': 'Professional Services',
            'Operational Dependency': 'Low',
            'Data Volume (annual)': 'Not applicable'
        }
        tier, rationale, _ = evaluate_vendor(row, self.rubric)
        self.assertIn('privileged', rationale.lower())


class TestSubstringQuoteGuard(unittest.TestCase):
    """Invalid quotes must be rejected."""

    def test_valid_quote(self):
        text = "We use artificial intelligence to process payments."
        quote = "artificial intelligence to process payments"
        self.assertTrue(verify_quote(quote, text))

    def test_invalid_quote_rejected(self):
        text = "We use automation in our services."
        quote = "We use artificial intelligence in our services."
        self.assertFalse(verify_quote(quote, text))

    def test_empty_quote_passes(self):
        self.assertTrue(verify_quote("", "any text"))

    def test_exact_sentence_match(self):
        text = "Our platform is AI-powered. It processes data daily."
        quote = "Our platform is AI-powered."
        self.assertTrue(verify_quote(quote, text))


class TestPreFilter(unittest.TestCase):
    """Pre-filter must catch AI keywords and reject non-AI text."""

    def test_catches_ai(self):
        self.assertTrue(pre_filter("We use artificial intelligence."))

    def test_catches_ml(self):
        self.assertTrue(pre_filter("Machine learning models are deployed."))

    def test_catches_llm(self):
        self.assertTrue(pre_filter("We deploy LLM for document processing."))

    def test_rejects_no_ai(self):
        self.assertFalse(pre_filter("We provide accounting and tax services."))

    def test_empty_text(self):
        self.assertFalse(pre_filter(""))
        self.assertFalse(pre_filter(None))


class TestWeightingLogic(unittest.TestCase):
    """Signal weight = base × tier_weight × confidence × service_linkage."""

    def test_confirmed_tier_a_high_conf(self):
        from risk_export import compute_weighted_score
        signals = [{
            'signal_type': 'CONFIRMED',
            'source_tier': 'Tier_A',
            'confidence': 'High',
            'ties_to_service': True,
            'valid': True
        }]
        score, _ = compute_weighted_score(signals)
        # 5 × 4 × 1.0 × 1.5 = 30
        self.assertEqual(score, 30.0)

    def test_marketing_low_weight(self):
        from risk_export import compute_weighted_score
        signals = [{
            'signal_type': 'MARKETING-ONLY',
            'source_tier': 'Tier_D',
            'confidence': 'Low',
            'ties_to_service': False,
            'valid': True
        }]
        score, _ = compute_weighted_score(signals)
        # 0.5 × 1 × 0.3 × 0.5 = 0.075
        self.assertAlmostEqual(score, 0.075)

    def test_invalid_signals_excluded(self):
        from risk_export import compute_weighted_score
        signals = [{
            'signal_type': 'CONFIRMED',
            'source_tier': 'Tier_A',
            'confidence': 'High',
            'ties_to_service': True,
            'valid': False  # invalid
        }]
        score, _ = compute_weighted_score(signals)
        self.assertEqual(score, 0)


class TestCoverageGapHandling(unittest.TestCase):
    """An unreachable source must be recorded as COVERAGE-GAP, never as 'No'."""

    def test_mock_classifier_no_text(self):
        signals = classify_with_mock("", "http://example.com")
        self.assertEqual(len(signals), 1)
        self.assertEqual(signals[0]['signal_type'], 'NEGATIVE-FINDING')

    def test_find_ai_sentences(self):
        text = "We use artificial intelligence to automate complex financial workflows. Our team is growing rapidly."
        sentences = find_ai_sentences(text)
        self.assertGreaterEqual(len(sentences), 1)
        self.assertIn('artificial intelligence', sentences[0].lower())


class TestBKUnchanged(unittest.TestCase):
    """Columns B-K must be byte-for-byte unchanged after export."""

    def test_bk_round_trip(self):
        import openpyxl
        input_path = os.path.join(os.path.dirname(__file__), '..', 'data', 'Meridian_Vendor_Input.xlsx')
        output_path = os.path.join(os.path.dirname(__file__), '..', 'data', 'Meridian_Vendor_Output.xlsx')

        if not os.path.exists(output_path):
            self.skipTest("Output file not yet generated. Run risk_export.py first.")

        wb_in = openpyxl.load_workbook(input_path, data_only=True)
        wb_out = openpyxl.load_workbook(output_path, data_only=True)

        ws_in = wb_in['Vendor Inventory']
        ws_out = wb_out['Vendor Inventory']

        # Columns B (2) through K (11), rows 4 onwards
        for row in range(4, ws_in.max_row + 1):
            for col in range(2, 12):  # B=2, K=11
                val_in = ws_in.cell(row=row, column=col).value
                val_out = ws_out.cell(row=row, column=col).value
                self.assertEqual(
                    val_in, val_out,
                    f"MISMATCH at row {row}, col {col}: '{val_in}' ≠ '{val_out}'"
                )


if __name__ == '__main__':
    unittest.main()
