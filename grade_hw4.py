#!/usr/bin/env python3
"""CLI for COMM2003 HW4 grader."""

from __future__ import annotations

import argparse
from pathlib import Path

from grader.pipeline import (
    grade_batch,
    grade_single_submission,
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

    return parser.parse_args()


def main() -> int:
    args = parse_args()

    if args.command == "single":
        student_key = extract_student_key(args.xlsx.name)
        grade = grade_single_submission(student_key, args.xlsx, args.written)
        write_single_report(args.output, grade)
        print(f"Wrote report: {args.output}")
        return 0

    grades = grade_batch(args.submissions_dir)
    summary = write_batch_reports(args.output_dir, grades)
    print(f"Processed {len(grades)} student pair(s).")
    print(f"Batch summary: {summary}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
