"""Agent C: written reflection scoring for Q1 and Q2."""

from __future__ import annotations

import logging
import re
from pathlib import Path

from docx import Document
from pypdf import PdfReader

from grader.constants import FOR_REVIEW
from grader.models import ReflectionEvaluation, ReflectionQuestionResult
from grader.utils import label_from_score, normalize_whitespace


logging.getLogger("pypdf").setLevel(logging.ERROR)


def run_reflection_agent(written_path: Path) -> ReflectionEvaluation:
    try:
        text = extract_text(written_path)
    except Exception as exc:  # noqa: BLE001
        review_flag = f"Unable to parse written response for scoring: {exc}"
        for_review = ReflectionQuestionResult(
            score=FOR_REVIEW,
            label="Missing",
            rationale="Written response could not be parsed confidently; manual review required.",
        )
        return ReflectionEvaluation(q1=for_review, q2=for_review, review_flags=[review_flag])

    q1_text, q2_text, segment_confident = split_q1_q2(text)

    q1 = score_q1(q1_text)
    q2 = score_q2(q2_text)

    review_flags = []
    if not segment_confident:
        review_flags.append("Written response Q1/Q2 segmentation confidence is low.")

    return ReflectionEvaluation(q1=q1, q2=q2, review_flags=review_flags)


def extract_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        doc = Document(path)
        text = "\n".join(p.text for p in doc.paragraphs)
        return normalize_whitespace(text)
    if suffix == ".pdf":
        reader = PdfReader(str(path))
        parts = [page.extract_text() or "" for page in reader.pages]
        return normalize_whitespace("\n".join(parts))
    raise ValueError(f"Unsupported written response type: {suffix}")


def split_q1_q2(text: str) -> tuple[str, str, bool]:
    lower = text.lower()

    q1_pattern = re.compile(
        r"(?:question\s*1|q\s*1|^\s*1[\).:-])\s*(.*?)(?=(?:question\s*2|q\s*2|\n\s*2[\).:-])|$)",
        flags=re.IGNORECASE | re.DOTALL,
    )
    q2_pattern = re.compile(
        r"(?:question\s*2|q\s*2|^\s*2[\).:-])\s*(.*)$",
        flags=re.IGNORECASE | re.DOTALL,
    )

    q1_match = q1_pattern.search(text)
    q2_match = q2_pattern.search(text)

    if q1_match and q2_match:
        q1_text = normalize_whitespace(q1_match.group(1))
        q2_text = normalize_whitespace(q2_match.group(1))
        return q1_text, q2_text, True

    # Fallback: split by paragraphs and assume first half is Q1 and second half is Q2.
    chunks = [c.strip() for c in re.split(r"\n{2,}", text) if c.strip()]
    if len(chunks) >= 2:
        midpoint = max(1, len(chunks) // 2)
        q1_text = normalize_whitespace(" ".join(chunks[:midpoint]))
        q2_text = normalize_whitespace(" ".join(chunks[midpoint:]))
        return q1_text, q2_text, False

    # Last fallback: no confident segmentation.
    return normalize_whitespace(text), "", False


def score_q1(answer: str) -> ReflectionQuestionResult:
    wc = word_count(answer)
    if wc < 8:
        return ReflectionQuestionResult(
            score=0.0,
            label="Missing",
            rationale="Response is missing or too limited to evaluate Question 1.",
        )

    criteria = {
        "client_questions": contains_any(answer, [
            "client", "intended use", "business question", "ask", "objective", "goal",
        ]),
        "data_clues": contains_any(answer, [
            "column", "header", "data type", "pattern", "inconsisten", "standardiz", "clue",
        ]),
        "join_identifier": contains_any(answer, [
            "customerid", "identifier", "join", "merge", "key",
        ]),
        "reasoning_depth": wc >= 70,
    }

    coverage = sum(1 for passed in criteria.values() if passed)

    if coverage >= 4:
        score = 2.0
        rationale = "Directly addresses client questions and data-derived clues with specific reasoning about joins and standardization."
    elif coverage >= 2 and wc >= 40:
        score = 1.25
        rationale = "Direction is generally correct, but the explanation lacks depth on how and why decisions would be made."
    else:
        score = 0.75
        rationale = "Response is vague or incomplete on client inquiry and data-as-clue reasoning required by Question 1."

    return ReflectionQuestionResult(score=score, label=label_from_score(score), rationale=rationale)


def score_q2(answer: str) -> ReflectionQuestionResult:
    wc = word_count(answer)
    if wc < 8:
        return ReflectionQuestionResult(
            score=0.0,
            label="Missing",
            rationale="Response is missing or too limited to evaluate Question 2.",
        )

    issue_hits = {
        "notes": contains_any(answer, ["notes", "@", "error", "corrupt character"]),
        "dollars_cents": contains_any(answer, ["dollars", "cents", "split", "total value"]),
        "null": contains_any(answer, ["null", "blank", "missing value"]),
        "transaction_id": contains_any(answer, ["transactionid", "tr", "format", "identifier"]),
    }
    issue_count = sum(1 for hit in issue_hits.values() if hit)

    has_specific_example = bool(re.search(r"\bT\d{4}\b|\bTR\d{4}\b", answer, flags=re.IGNORECASE)) or contains_any(
        answer,
        ["for example", "such as", "example"],
    )

    has_downstream_link = contains_any(
        answer,
        ["calculation", "lookup", "match", "report", "aggregate", "downstream", "analyst", "trust", "accuracy"],
    )

    if issue_count >= 2 and has_specific_example and has_downstream_link and wc >= 80:
        score = 2.0
        rationale = "Addresses two distinct issues with specific examples and clearly links each issue to downstream impact."
    elif issue_count >= 2 and (has_specific_example or has_downstream_link) and wc >= 45:
        score = 1.25
        rationale = "Identifies relevant issues, but examples or downstream impact explanation are not consistently specific."
    else:
        score = 0.75
        rationale = "Does not substantively connect two issue examples to concrete downstream consequences."

    return ReflectionQuestionResult(score=score, label=label_from_score(score), rationale=rationale)


def contains_any(text: str, keywords: list[str]) -> bool:
    lowered = text.lower()
    return any(keyword.lower() in lowered for keyword in keywords)


def word_count(text: str) -> int:
    return len(re.findall(r"\b\w+\b", text))
