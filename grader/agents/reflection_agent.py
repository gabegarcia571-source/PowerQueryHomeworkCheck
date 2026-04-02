"""Agent C: written reflection scoring for Q1 and Q2."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document

from grader.constants import FOR_REVIEW
from grader.models import ReflectionEvaluation, ReflectionQuestionResult
from grader.utils import label_from_score, normalize_whitespace


logging.getLogger("pdf2docx").setLevel(logging.WARNING)


PAGE_SEPARATOR = "\n\f\n"


@dataclass
class Marker:
    start: int
    end: int
    confident: bool


@dataclass
class SplitResult:
    q1_text: str
    q2_text: str
    segment_confident: bool
    q2_start_page: int | None
    q2_end_page: int | None


def run_reflection_agent(written_path: Path) -> ReflectionEvaluation:
    try:
        text, page_texts = extract_text_with_pages(written_path)
    except Exception as exc:  # noqa: BLE001
        review_flag = f"Unable to parse written response for scoring: {exc}"
        for_review = ReflectionQuestionResult(
            score=FOR_REVIEW,
            label="Missing",
            rationale="Written response could not be parsed confidently; manual review required.",
        )
        return ReflectionEvaluation(
            q1=for_review,
            q2=for_review,
            q2_start_page=None,
            q2_end_page=None,
            review_flags=[review_flag],
        )

    split = split_q1_q2_with_pages(text, page_texts)

    q1 = score_q1(split.q1_text)
    q2 = score_q2(split.q2_text)

    review_flags = []
    if not split.segment_confident:
        review_flags.append("Written response Q1/Q2 segmentation confidence is low.")

    return ReflectionEvaluation(
        q1=q1,
        q2=q2,
        q2_start_page=split.q2_start_page,
        q2_end_page=split.q2_end_page,
        review_flags=review_flags,
    )


def extract_text(path: Path) -> str:
    text, _ = extract_text_with_pages(path)
    return text


def extract_text_with_pages(path: Path) -> tuple[str, list[str]]:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return _extract_docx_text_with_pages(path)
    if suffix == ".pdf":
        return _extract_pdf_text_via_docx(path)
    raise ValueError(f"Unsupported written response type: {suffix}")


def _extract_docx_text_with_pages(path: Path) -> tuple[str, list[str]]:
    doc = Document(path)
    text = "\n\n".join(p.text for p in doc.paragraphs)
    page_texts = [_preserve_layout(text)]
    return _join_pages(page_texts), page_texts


def _extract_pdf_text_via_docx(pdf_path: Path) -> tuple[str, list[str]]:
    with TemporaryDirectory(prefix="grading-pdf-docx-") as temp_dir:
        converted_docx_path = Path(temp_dir) / f"{pdf_path.stem}.docx"
        _convert_pdf_to_docx(pdf_path, converted_docx_path)
        return _extract_docx_text_with_pages(converted_docx_path)


def _convert_pdf_to_docx(pdf_path: Path, output_docx_path: Path) -> None:
    try:
        from pdf2docx import Converter
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(
            "PDF conversion dependency missing. Install 'pdf2docx' to grade PDF reflections."
        ) from exc

    root_logger = logging.getLogger()
    previous_level = root_logger.level
    root_logger.setLevel(logging.WARNING)

    converter = None
    try:
        converter = Converter(str(pdf_path))
        converter.convert(str(output_docx_path))
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(f"PDF to DOCX conversion failed for '{pdf_path.name}': {exc}") from exc
    finally:
        if converter is not None:
            converter.close()
        root_logger.setLevel(previous_level)

    if not output_docx_path.exists():
        raise RuntimeError(f"PDF to DOCX conversion did not produce output for '{pdf_path.name}'.")


def split_q1_q2(text: str) -> tuple[str, str, bool]:
    result = split_q1_q2_with_pages(text, [text])
    return result.q1_text, result.q2_text, result.segment_confident


def split_q1_q2_with_pages(text: str, page_texts: list[str]) -> SplitResult:
    preserved_text = _preserve_layout(text)
    if preserved_text == "":
        return SplitResult("", "", False, None, None)

    q1_marker = _find_question_marker(preserved_text, question_number=1, search_start=0)
    q2_search_start = q1_marker.end if q1_marker else 0
    q2_marker = _find_question_marker(
        preserved_text,
        question_number=2,
        search_start=q2_search_start,
    )

    if q2_marker:
        q2_start = q2_marker.start
        q2_body_start = q2_marker.end
        q2_end = _find_q2_end_boundary(preserved_text, q2_body_start)

        q1_start = q1_marker.end if q1_marker and q1_marker.start < q2_start else 0
        q1_text = normalize_whitespace(preserved_text[q1_start:q2_start])
        q2_text = normalize_whitespace(preserved_text[q2_body_start:q2_end])

        start_page, end_page = _pages_for_span(
            page_texts=page_texts,
            span_start=q2_start,
            span_end=max(q2_start, q2_end - 1),
        )
        return SplitResult(q1_text, q2_text, q2_marker.confident, start_page, end_page)

    # Fallback: split by paragraph blocks from non-appendix text only.
    text_without_appendix = _truncate_before_appendix(preserved_text)
    chunks = [c.strip() for c in re.split(r"\n{2,}", text_without_appendix) if c.strip()]
    if len(chunks) >= 2:
        midpoint = max(1, len(chunks) // 2)
        q1_text = normalize_whitespace("\n\n".join(chunks[:midpoint]))
        q2_text = normalize_whitespace("\n\n".join(chunks[midpoint:]))

        q2_start = text_without_appendix.find(chunks[midpoint])
        q2_end = len(text_without_appendix)
        start_page, end_page = _pages_for_span(
            page_texts=page_texts,
            span_start=q2_start,
            span_end=max(q2_start, q2_end - 1),
        ) if q2_start >= 0 and q2_text else (None, None)
        return SplitResult(q1_text, q2_text, False, start_page, end_page)

    # Last fallback: no confident segmentation.
    return SplitResult(normalize_whitespace(text_without_appendix), "", False, None, None)


def _join_pages(page_texts: list[str]) -> str:
    return PAGE_SEPARATOR.join(page_texts)


def _preserve_layout(value: str) -> str:
    text = (value or "").replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{4,}", "\n\n\n", text)
    return text.strip()


def _truncate_before_appendix(text: str) -> str:
    appendix_pattern = re.compile(r"(?im)^\s*(?:appendix|references|works cited|bibliography)\b")
    appendix_match = appendix_pattern.search(text)
    if appendix_match:
        return text[:appendix_match.start()].rstrip()
    return text


def _find_q2_end_boundary(text: str, search_start: int) -> int:
    pattern = re.compile(
        r"(?im)^\s*(?:"
        r"(?:question|q|ques)\s*[:.\-]?\s*(?:3|iii|three)\b\s*[:.\-]?"
        r"|(?:3|iii)\s*[\).:-]\s*"
        r"|(?:appendix|references|works cited|bibliography)\b"
        r")",
    )
    match = pattern.search(text, pos=search_start)
    return match.start() if match else len(text)


def _find_question_marker(text: str, question_number: int, search_start: int) -> Marker | None:
    question_tokens = {
        1: "(?:1|i|one)",
        2: "(?:2|ii|two)",
    }
    numbered_tokens = {
        1: "(?:1|i)",
        2: "(?:2|ii)",
    }
    question_token = question_tokens[question_number]
    numbered_token = numbered_tokens[question_number]

    heading_patterns = [
        re.compile(
            rf"(?im)^\s*(?:question|q|ques)\s*[:.\-]?\s*{question_token}\b\s*[:.\-]?",
        ),
        re.compile(rf"(?im)^\s*{numbered_token}\s*[\).:-]\s*"),
        re.compile(
            rf"(?im)^\s*(?:part|section)\s*(?:{'a|1' if question_number == 1 else 'b|2'})\b\s*[:.\-]?",
        ),
    ]
    embedded_patterns = [
        re.compile(
            rf"(?i)\b(?:for\s+)?(?:question|q)\s*[:.\-]?\s*{question_token}\b\s*[:.\-]?",
        ),
    ]

    best = _earliest_marker(text, search_start, heading_patterns)
    if best:
        return best

    return _earliest_marker(text, search_start, embedded_patterns)


def _earliest_marker(text: str, search_start: int, patterns: list[re.Pattern[str]]) -> Marker | None:
    best: Marker | None = None
    for pattern in patterns:
        match = pattern.search(text, pos=search_start)
        if not match:
            continue
        marker = Marker(start=match.start(), end=match.end(), confident=True)
        if best is None or marker.start < best.start:
            best = marker

    return best


def _page_starts(page_texts: list[str]) -> list[int]:
    starts: list[int] = []
    cursor = 0
    for index, page_text in enumerate(page_texts):
        starts.append(cursor)
        cursor += len(page_text)
        if index < len(page_texts) - 1:
            cursor += len(PAGE_SEPARATOR)
    return starts


def _offset_to_page(offset: int, starts: list[int]) -> int | None:
    if not starts:
        return None
    if offset <= starts[0]:
        return 1

    for index, start in enumerate(starts):
        next_start = starts[index + 1] if index + 1 < len(starts) else None
        if next_start is None or offset < next_start:
            return index + 1

    return len(starts)


def _pages_for_span(page_texts: list[str], span_start: int, span_end: int) -> tuple[int | None, int | None]:
    starts = _page_starts(page_texts)
    start_page = _offset_to_page(span_start, starts)
    end_page = _offset_to_page(span_end, starts)
    return start_page, end_page


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
    normalized = normalize_whitespace(answer)
    if normalized == "":
        return ReflectionQuestionResult(
            score=0.0,
            label="Missing",
            rationale="Response is missing or too limited to evaluate Question 2.",
        )

    wc = word_count(normalized)

    issue_hits = {
        "notes": contains_any(normalized, ["notes", "@", "error", "corrupt character"]),
        "dollars_cents": contains_any(normalized, ["dollars", "cents", "split", "total value"]),
        "null": contains_any(normalized, ["null", "blank", "missing value"]),
        "transaction_id": contains_any(normalized, ["transactionid", "tr", "format", "identifier"]),
    }
    issue_count = sum(1 for hit in issue_hits.values() if hit)

    has_specific_example = bool(re.search(r"\bT\d{4}\b|\bTR\d{4}\b", normalized, flags=re.IGNORECASE)) or contains_any(
        normalized,
        ["for example", "such as", "example"],
    )

    has_downstream_link = contains_any(
        normalized,
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
