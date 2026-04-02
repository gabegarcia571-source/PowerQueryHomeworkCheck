#!/usr/bin/env python3
"""CLI for COMM2003 HW4 grader."""

from __future__ import annotations

import argparse
from pathlib import Path

from grader.pipeline import (
    write_batch_feedback_sheets,
    grade_batch,
    grade_single_submission,
    write_single_feedback_sheet,
    write_batch_reports,
    write_single_report,
)
from grader.utils import extract_student_key


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Grade COMM2003 HW4 submissions.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    single = subparsers.add_parser("single", help="Grade a single student submission pair.")
    single.add_argument("--xlsx", required=True, type=Path, help="Path to student XLSX file.")
    single.add_argument("--written", required=True, type=Path, help="Path to student DOCX/PDF reflection file.")
    single.add_argument("--output", required=True, type=Path, help="Output text report path.")
    single.add_argument(
        "--feedback-sheets-dir",
        default=Path("Feedback Sheets"),
        type=Path,
        help="Directory containing feedback sheet templates.",
    )
    single.add_argument(
        "--filled-feedback-output",
        default=None,
        type=Path,
        help="Optional output DOCX path for populated feedback sheet.",
    )

    batch = subparsers.add_parser("batch", help="Grade all submission pairs in a directory.")
    batch.add_argument(
        "--submissions-dir",
        default=Path("Submissions"),
        type=Path,
        help="Directory containing student submissions.",
    )
    batch.add_argument(
        "--output-dir",
        default=Path("Generated Reports"),
        type=Path,
        help="Directory for generated reports.",
    )
    batch.add_argument(
        "--feedback-sheets-dir",
        default=Path("Feedback Sheets"),
        type=Path,
        help="Directory containing feedback sheet templates.",
    )
    batch.add_argument(
        "--filled-feedback-dir",
        default=None,
        type=Path,
        help="Directory for populated feedback sheet DOCX files. Default: <output-dir>/Filled Feedback Sheets",
    )

    return parser.parse_args()


def main() -> int:
    args = parse_args()

    if args.command == "single":
        student_key = extract_student_key(args.xlsx.name)
        grade = grade_single_submission(student_key, args.xlsx, args.written)
        write_single_report(args.output, grade)
        print(f"Wrote report: {args.output}")

        if args.filled_feedback_output is not None:
            wrote_feedback = write_single_feedback_sheet(
                student_key=student_key,
                feedback_sheets_dir=args.feedback_sheets_dir,
                output_path=args.filled_feedback_output,
                grade=grade,
            )
            if wrote_feedback:
                print(f"Wrote feedback sheet: {args.filled_feedback_output}")
            else:
                print(
                    f"Skipped feedback sheet: no template found for key '{student_key}' in {args.feedback_sheets_dir}"
                )
        return 0

    grades = grade_batch(args.submissions_dir)
    summary = write_batch_reports(args.output_dir, grades)
    filled_feedback_dir = args.filled_feedback_dir or (args.output_dir / "Filled Feedback Sheets")
    filled_count, missing_templates = write_batch_feedback_sheets(
        feedback_sheets_dir=args.feedback_sheets_dir,
        output_dir=filled_feedback_dir,
        grades=grades,
    )

    print(f"Processed {len(grades)} student pair(s).")
    print(f"Batch summary: {summary}")
    print(f"Filled feedback sheets: {filled_count}")
    print(f"Feedback sheet output dir: {filled_feedback_dir}")
    if missing_templates:
        preview = ", ".join(missing_templates[:10])
        remainder = len(missing_templates) - min(len(missing_templates), 10)
        more_text = "" if remainder == 0 else f" (+{remainder} more)"
        print(f"Missing feedback templates for keys: {preview}{more_text}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
