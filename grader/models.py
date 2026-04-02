"""Dataclasses shared across grading agents."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Union

ScoreValue = Union[float, str]


@dataclass
class IntegrityResult:
    student_key: str
    student_display_name: str
    xlsx_path: Path | None
    written_path: Path | None
    passed: bool
    reasons: list[str] = field(default_factory=list)


@dataclass
class XlsxEvaluation:
    score_1a: ScoreValue
    notes_1a: str
    score_1b: ScoreValue
    errors_1b: list[str] = field(default_factory=list)
    review_flags_1b: list[str] = field(default_factory=list)
    score_1c: ScoreValue = 0.0
    notes_1c: str = ""
    score_1d: ScoreValue = 0.0
    notes_1d: str = ""
    review_flags: list[str] = field(default_factory=list)


@dataclass
class ReflectionQuestionResult:
    score: ScoreValue
    label: str
    rationale: str


@dataclass
class ReflectionEvaluation:
    q1: ReflectionQuestionResult
    q2: ReflectionQuestionResult
    q2_start_page: int | None = None
    q2_end_page: int | None = None
    review_flags: list[str] = field(default_factory=list)


@dataclass
class FinalGrade:
    student_name: str
    total_score: ScoreValue
    xlsx_eval: XlsxEvaluation
    reflection_eval: ReflectionEvaluation
    flags_for_review: list[str] = field(default_factory=list)
