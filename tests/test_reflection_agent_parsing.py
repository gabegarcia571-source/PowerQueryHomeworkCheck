"""Tests for reflection parsing and Q2 extraction behavior."""

from __future__ import annotations

import unittest
from pathlib import Path
from unittest import mock

from grader.constants import FOR_REVIEW
from grader.agents.reflection_agent import (
    extract_text,
    run_reflection_agent,
    score_q2,
    split_q1_q2_with_pages,
)


class ReflectionAgentParsingTests(unittest.TestCase):
    @mock.patch("grader.agents.reflection_agent._extract_docx_text_with_pages")
    @mock.patch("grader.agents.reflection_agent._convert_pdf_to_docx")
    def test_extract_text_converts_pdf_to_docx_before_reading(
        self,
        mock_convert: mock.Mock,
        mock_extract_docx: mock.Mock,
    ) -> None:
        mock_extract_docx.return_value = (
            "Question 1:\n- client objective\nQuestion 2:\n- notes contain @ symbols",
            ["Question 1:\n- client objective\nQuestion 2:\n- notes contain @ symbols"],
        )

        text = extract_text(Path("reflection.pdf"))

        self.assertIn("- client objective", text)

        mock_convert.assert_called_once()
        src_path, converted_path = mock_convert.call_args.args
        self.assertEqual(src_path, Path("reflection.pdf"))
        self.assertEqual(converted_path.suffix, ".docx")
        mock_extract_docx.assert_called_once_with(converted_path)

    @mock.patch("grader.agents.reflection_agent._convert_pdf_to_docx")
    @mock.patch("grader.agents.reflection_agent.Document")
    def test_extract_text_docx_does_not_convert_pdf(
        self,
        mock_document: mock.Mock,
        mock_convert: mock.Mock,
    ) -> None:
        paragraph1 = mock.Mock()
        paragraph1.text = "Question 1: client objective"
        paragraph2 = mock.Mock()
        paragraph2.text = "Question 2: downstream impact"
        mock_document.return_value.paragraphs = [paragraph1, paragraph2]

        text = extract_text(Path("reflection.docx"))

        self.assertIn("Question 1: client objective", text)
        self.assertIn("Question 2: downstream impact", text)
        mock_convert.assert_not_called()

    @mock.patch(
        "grader.agents.reflection_agent._convert_pdf_to_docx",
        side_effect=RuntimeError("PDF to DOCX conversion failed for 'reflection.pdf': converter error"),
    )
    def test_run_reflection_agent_returns_for_review_on_conversion_failure(self, _: mock.Mock) -> None:
        result = run_reflection_agent(Path("reflection.pdf"))

        self.assertEqual(result.q1.score, FOR_REVIEW)
        self.assertEqual(result.q2.score, FOR_REVIEW)
        self.assertTrue(any("conversion failed" in flag.lower() for flag in result.review_flags))

    def test_split_detects_embedded_q2_and_page_range(self) -> None:
        page_texts = [
            "Question 1: I would ask the client about intended use and join keys.",
            "In Question 2: I found notes with @ symbols and dollars/cents split across columns.",
            "This causes bad downstream report aggregation and lookup mismatches.\nAppendix A\nScreenshot captions only.",
        ]
        text = "\n\f\n".join(page_texts)

        result = split_q1_q2_with_pages(text, page_texts)

        self.assertIn("notes with @ symbols", result.q2_text.lower())
        self.assertNotIn("appendix a", result.q2_text.lower())
        self.assertEqual(result.q2_start_page, 2)
        self.assertEqual(result.q2_end_page, 3)
        self.assertTrue(result.segment_confident)

    def test_split_ignores_unrelated_appendix_when_q2_absent(self) -> None:
        page_texts = [
            "Question 1: I would inspect the worksheet names and join identifiers first.",
            "Appendix A\nScanned screenshot descriptions.",
        ]
        text = "\n\f\n".join(page_texts)

        result = split_q1_q2_with_pages(text, page_texts)

        self.assertEqual(result.q2_text, "")
        self.assertIsNone(result.q2_start_page)
        self.assertIsNone(result.q2_end_page)

    def test_score_q2_only_returns_missing_when_blank(self) -> None:
        missing = score_q2("   ")
        present_but_short = score_q2("Notes include @ characters.")

        self.assertEqual(missing.score, 0.0)
        self.assertEqual(missing.label, "Missing")
        self.assertEqual(present_but_short.score, 0.75)
        self.assertNotEqual(present_but_short.label, "Missing")


if __name__ == "__main__":
    unittest.main()
