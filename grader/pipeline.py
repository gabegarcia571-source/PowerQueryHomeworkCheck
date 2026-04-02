"""Pipeline orchestration for COMM2003 HW4 grading."""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from grader.feedback_sheet_writer import (
    write_feedback_sheet_for_student,
    write_feedback_sheets_for_batch,
)
from grader.agents.integrity_agent import discover_pairs, run_integrity_agent
from grader.agents.reflection_agent import run_reflection_agent
from grader.agents.report_agent import compute_total_score, render_report
from grader.agents.xlsx_agent import run_xlsx_agent
from grader.constants import FOR_REVIEW
from grader.models import (
    FinalGrade,
    ReflectionEvaluation,
    ReflectionQuestionResult,
    XlsxEvaluation,
)
from grader.utils import extract_student_display_name


def grade_single_submission(student_key: str, xlsx_path: Path | None, written_path: Path | None) -> FinalGrade:
    integrity = run_integrity_agent(student_key, xlsx_path, written_path)

    if not integrity.passed:
        return build_for_review_grade(integrity.student_display_name, integrity.reasons)

    assert integrity.xlsx_path is not None
    assert integrity.written_path is not None

    with ThreadPoolExecutor(max_workers=2) as executor:
        xlsx_future = executor.submit(run_xlsx_agent, integrity.xlsx_path)
        reflection_future = executor.submit(run_reflection_agent, integrity.written_path)
        xlsx_eval = xlsx_future.result()
        reflection_eval = reflection_future.result()

    total = compute_total_score([
        xlsx_eval.score_1a,
        xlsx_eval.score_1b,
        xlsx_eval.score_1c,
        xlsx_eval.score_1d,
        reflection_eval.q1.score,
        reflection_eval.q2.score,
    ])

    review_flags = xlsx_eval.review_flags + reflection_eval.review_flags

    return FinalGrade(
        student_name=integrity.student_display_name,
        total_score=total,
        xlsx_eval=xlsx_eval,
        reflection_eval=reflection_eval,
        flags_for_review=review_flags,
    )


def grade_batch(submissions_dir: Path) -> dict[str, FinalGrade]:
    pairs = discover_pairs(submissions_dir)
    results: dict[str, FinalGrade] = {}

    for student_key in sorted(pairs):
        pair = pairs[student_key]
        results[student_key] = grade_single_submission(
            student_key=student_key,
            xlsx_path=pair.get("xlsx"),
            written_path=pair.get("written"),
        )

    return results


def write_single_report(report_path: Path, grade: FinalGrade) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(render_report(grade), encoding="utf-8")


def write_batch_reports(output_dir: Path, grades: dict[str, FinalGrade]) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)

    for key, grade in grades.items():
        report_file = output_dir / f"{key}_grading_report.txt"
        report_file.write_text(render_report(grade), encoding="utf-8")

    summary_path = output_dir / "batch_summary.csv"
    lines = ["student_key,student_name,total_score,flag_count"]
    for key, grade in sorted(grades.items()):
        score = grade.total_score if isinstance(grade.total_score, str) else f"{grade.total_score:.2f}"
        lines.append(f"{key},{sanitize_csv(grade.student_name)},{sanitize_csv(str(score))},{len(grade.flags_for_review)}")

    summary_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return summary_path


def write_single_feedback_sheet(
    student_key: str,
    feedback_sheets_dir: Path,
    output_path: Path,
    grade: FinalGrade,
) -> bool:
    template = write_feedback_sheet_for_student(
        student_key=student_key,
        grade=grade,
        feedback_sheets_dir=feedback_sheets_dir,
        output_path=output_path,
    )
    return template is not None


def write_batch_feedback_sheets(
    feedback_sheets_dir: Path,
    output_dir: Path,
    grades: dict[str, FinalGrade],
) -> tuple[int, list[str]]:
    return write_feedback_sheets_for_batch(
        grades=grades,
        feedback_sheets_dir=feedback_sheets_dir,
        output_dir=output_dir,
    )


def build_for_review_grade(student_name: str, reasons: list[str]) -> FinalGrade:
    xlsx_eval = XlsxEvaluation(
        score_1a=FOR_REVIEW,
        notes_1a="Integrity check failed before workbook scoring.",
        score_1b=FOR_REVIEW,
        errors_1b=["Skipped due to integrity failure."],
        score_1c=FOR_REVIEW,
        notes_1c="Skipped due to integrity failure.",
        score_1d=FOR_REVIEW,
        notes_1d="Skipped due to integrity failure.",
        review_flags=reasons,
    )

    reflection_eval = ReflectionEvaluation(
        q1=ReflectionQuestionResult(
            score=FOR_REVIEW,
            label="Missing",
            rationale="Skipped due to integrity failure.",
        ),
        q2=ReflectionQuestionResult(
            score=FOR_REVIEW,
            label="Missing",
            rationale="Skipped due to integrity failure.",
        ),
        review_flags=reasons,
    )

    return FinalGrade(
        student_name=student_name,
        total_score=FOR_REVIEW,
        xlsx_eval=xlsx_eval,
        reflection_eval=reflection_eval,
        flags_for_review=reasons,
    )


def sanitize_csv(value: str) -> str:
    text = value.replace('"', '""')
    if "," in text or "\n" in text:
        return f'"{text}"'
    return text
