"""Smoke tests for COMM2003 HW4 grader core behavior."""

from __future__ import annotations

import unittest

from grader.agents.report_agent import compute_total_score, render_report
from grader.models import (
    FinalGrade,
    ReflectionEvaluation,
    ReflectionQuestionResult,
    XlsxEvaluation,
)
from grader.utils import extract_student_key


class GraderSmokeTests(unittest.TestCase):
    def test_extract_student_key(self) -> None:
        name = "jopatrick_1285366_18930213_Patrick Jo - COMM 2003 HW 4.xlsx"
        self.assertEqual(extract_student_key(name), "jopatrick")

    def test_total_score_for_review_propagates(self) -> None:
        total = compute_total_score([1.0, 4.0, "FOR REVIEW", 0.5])
        self.assertEqual(total, "FOR REVIEW")

    def test_report_contains_required_sections(self) -> None:
        xlsx_eval = XlsxEvaluation(
            score_1a=1.0,
            notes_1a="All seven worksheets present.",
            score_1b=4.0,
            errors_1b=["No errors found."],
            score_1c=0.5,
            notes_1c="Pivot object exists and required dimensions/value field checks passed.",
            score_1d=0.5,
            notes_1d="Pivot object exists and required dimensions/value field checks passed.",
            review_flags=[],
        )
        reflection_eval = ReflectionEvaluation(
            q1=ReflectionQuestionResult(
                score=2.0,
                label="Excellent",
                rationale="Direct and specific response.",
            ),
            q2=ReflectionQuestionResult(
                score=2.0,
                label="Excellent",
                rationale="Direct and specific response.",
            ),
            review_flags=[],
        )
        grade = FinalGrade(
            student_name="Test Student",
            total_score=10.0,
            xlsx_eval=xlsx_eval,
            reflection_eval=reflection_eval,
            flags_for_review=[],
        )

        report = render_report(grade)
        self.assertIn("SECTION 1: XLSX EVALUATION", report)
        self.assertIn("SECTION 2: WRITTEN REFLECTION", report)
        self.assertIn("FLAGS FOR INSTRUCTOR REVIEW", report)


if __name__ == "__main__":
    unittest.main()
