"""Agent D: final score aggregation and report rendering."""

from __future__ import annotations

from grader.constants import FOR_REVIEW
from grader.models import FinalGrade, ScoreValue
from grader.utils import format_score


def compute_total_score(scores: list[ScoreValue]) -> ScoreValue:
    if any(isinstance(score, str) for score in scores):
        return FOR_REVIEW
    return float(sum(score for score in scores if isinstance(score, (int, float))))


def render_report(grade: FinalGrade) -> str:
    x = grade.xlsx_eval
    r = grade.reflection_eval

    score_1a = score_text(x.score_1a)
    score_1b = score_text(x.score_1b)
    score_1c = score_text(x.score_1c)
    score_1d = score_text(x.score_1d)
    score_q1 = score_text(r.q1.score)
    score_q2 = score_text(r.q2.score)
    total = score_text(grade.total_score)

    errors_1b_lines = ["    - " + item for item in x.errors_1b] if x.errors_1b else ["    - No errors found."]
    errors_1b_text = "\n".join(errors_1b_lines)

    flags = grade.flags_for_review if grade.flags_for_review else ["None"]
    flags_text = "; ".join(flags)

    return (
        f"STUDENT: {grade.student_name}\n"
        f"TOTAL SCORE: {total} / 10\n\n"
        "────────────────────────────────────────────────────────\n"
        "SECTION 1: XLSX EVALUATION\n"
        "────────────────────────────────────────────────────────\n\n"
        f"1A — Worksheets Present: {score_1a} / 1.0\n"
        f"  Notes: {x.notes_1a}\n\n"
        f"1B — Cleaned Dataset: {score_1b} / 4.0\n"
        "  Errors found (−0.5 each):\n"
        f"{errors_1b_text}\n\n"
        f"1C — Report 1 Pivot Table: {score_1c} / 0.5\n"
        f"  Notes: {x.notes_1c}\n\n"
        f"1D — Report 2 Pivot Table: {score_1d} / 0.5\n"
        f"  Notes: {x.notes_1d}\n\n"
        "────────────────────────────────────────────────────────\n"
        "SECTION 2: WRITTEN REFLECTION\n"
        "────────────────────────────────────────────────────────\n\n"
        f"Q1: {score_q1} / 2.0 — {r.q1.label}\n"
        f"  Rationale: {r.q1.rationale}\n\n"
        f"Q2: {score_q2} / 2.0 — {r.q2.label}\n"
        f"  Rationale: {r.q2.rationale}\n\n"
        "────────────────────────────────────────────────────────\n"
        f"FLAGS FOR INSTRUCTOR REVIEW: {flags_text}\n"
        "────────────────────────────────────────────────────────\n"
    )


def score_text(score: ScoreValue) -> str:
    if isinstance(score, str):
        return score
    return format_score(float(score))
