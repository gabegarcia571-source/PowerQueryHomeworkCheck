"""Utility functions for parsing and normalization."""

from __future__ import annotations

import re
from datetime import date, datetime, time
from pathlib import Path


def compact_text(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (value or "").lower())


def resolve_required_columns(
    headers: list[str],
    required_columns: list[str],
    aliases: dict[str, list[str]] | None = None,
) -> tuple[dict[str, str | None], list[str], list[str], list[str]]:
    """Resolve required columns against extracted headers.

    Matching is done on compacted text and accepts trailing-character noise
    in actual headers (for example, ``TransactionIDx`` still matches
    ``TransactionID``). Aliases can be supplied per required column to
    allow alternate accepted names. Returns:
    - required_to_header: mapping of required column name to matched header or None
    - missing_required: required columns that could not be resolved
    - extra_headers: actual headers that were not used by any required column
    - ambiguous_required: required columns that had multiple possible matches
    """

    normalized_headers: list[tuple[int, str, str]] = [
        (index, header, compact_text(header))
        for index, header in enumerate(headers)
    ]

    required_to_header: dict[str, str | None] = {}
    used_headers: set[str] = set()
    ambiguous_required: list[str] = []
    alias_map = aliases or {}

    for required in required_columns:
        required_keys = [compact_text(required)]
        required_keys.extend(compact_text(alias) for alias in alias_map.get(required, []))
        required_keys = [key for key in dict.fromkeys(required_keys) if key]

        candidates: list[tuple[int, int, int, str]] = []

        for index, header, header_key in normalized_headers:
            best_candidate: tuple[int, int, int, str] | None = None
            for required_key in required_keys:
                if header_key == required_key:
                    # Prefer exact compact-match first.
                    candidate = (0, 0, index, header)
                elif header_key.startswith(required_key):
                    trailing_len = len(header_key) - len(required_key)
                    candidate = (1, trailing_len, index, header)
                else:
                    continue

                if best_candidate is None or candidate[:2] < best_candidate[:2]:
                    best_candidate = candidate

            if best_candidate is not None:
                candidates.append(best_candidate)

        candidates.sort(key=lambda item: (item[0], item[1], item[2], item[3].lower()))
        available = [item for item in candidates if item[3] not in used_headers]

        if not available:
            required_to_header[required] = None
            if candidates:
                ambiguous_required.append(required)
            continue

        if len(available) > 1:
            ambiguous_required.append(required)

        chosen_header = available[0][3]
        required_to_header[required] = chosen_header
        used_headers.add(chosen_header)

    missing_required = [name for name in required_columns if required_to_header.get(name) is None]
    extra_headers = [header for header in headers if header not in used_headers]
    return required_to_header, missing_required, extra_headers, ambiguous_required


def normalize_whitespace(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip())


def is_blank(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    return False


def safe_str(value: object) -> str:
    if value is None:
        return ""
    return str(value)


def has_time_component(value: object) -> bool:
    if isinstance(value, datetime):
        return value.time() != time(0, 0)
    if isinstance(value, date):
        return False
    if isinstance(value, str):
        text = value.strip().lower()
        if re.search(r"\b(am|pm)\b", text):
            return True
        if re.search(r"\d{1,2}:\d{2}", text):
            return True
    return False


def is_t_id(value: object) -> bool:
    return bool(re.fullmatch(r"T\d{4}", safe_str(value).strip()))


def looks_unsplit_merchant(value: object) -> bool:
    text = safe_str(value).strip()
    if text == "":
        return False
    if re.search(r"\s-\s.+,\s*[A-Za-z]{2}\s*$", text):
        return True
    if re.search(r",\s*[A-Za-z]{2}\s*$", text):
        return True
    return False


def is_explicit_null_marker(value: object) -> bool:
    if not isinstance(value, str):
        return False
    marker = value.strip().lower()
    return marker in {
        "null",
        "none",
        "nan",
        "n/a",
        "na",
        "<null>",
        "(null)",
    }


def is_date_only_number_format(number_format: object) -> bool:
    fmt = safe_str(number_format).strip().lower()
    if fmt == "" or fmt == "general":
        return False
    if "h" in fmt or "s" in fmt or "am/pm" in fmt:
        return False
    return "y" in fmt and "m" in fmt and "d" in fmt


def is_datetime_number_format(number_format: object) -> bool:
    fmt = safe_str(number_format).strip().lower()
    if fmt == "" or fmt == "general":
        return False

    # Remove quoted literals and escaped characters so token checks focus on
    # actual Excel date/time specifiers.
    fmt = re.sub(r'"[^"]*"', "", fmt)
    fmt = re.sub(r"\\.", "", fmt)

    if "am/pm" in fmt:
        return True
    return "h" in fmt or "s" in fmt


def is_currency_number_format(number_format: object) -> bool:
    fmt = safe_str(number_format).strip().lower()
    if fmt == "" or fmt == "general":
        return False
    if "$" in fmt or "€" in fmt or "£" in fmt or "¥" in fmt:
        return True
    if "currency" in fmt or "accounting" in fmt:
        return True
    if "_($" in fmt or "[$" in fmt:
        return True
    return False


def is_numeric_cell(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def format_score(value: float) -> str:
    if abs(value - round(value)) < 1e-9:
        return f"{int(round(value))}.0"
    return f"{value:.2f}".rstrip("0").rstrip(".")


def extract_student_key(filename: str) -> str:
    match = re.match(r"([^_]+)_\d+_\d+_.*", filename)
    if match:
        return match.group(1).lower()
    stem = Path(filename).stem
    return compact_text(stem.split("_")[0]) or compact_text(stem)


def extract_student_display_name(filename: str, fallback_key: str) -> str:
    stem = Path(filename).stem
    parts = stem.split("_")
    candidate = parts[3] if len(parts) > 3 else ""
    candidate = re.sub(r"\b(comm|homework|hw|data|cleaning|aggregation|using|power|query|responses?|answers?)\b", "", candidate, flags=re.IGNORECASE)
    candidate = re.sub(r"[-–_]+", " ", candidate)
    candidate = normalize_whitespace(candidate)
    words = [w for w in candidate.split() if re.search(r"[A-Za-z]", w)]
    if len(words) >= 2:
        # Keep first two alphabetic chunks for a stable display name.
        return f"{words[0].title()} {words[1].title()}"
    return fallback_key.title()


def label_from_score(score: float) -> str:
    if abs(score - 2.0) < 1e-9:
        return "Excellent"
    if abs(score - 1.25) < 1e-9:
        return "Meets Expectations"
    if abs(score - 0.75) < 1e-9:
        return "Does Not Meet Expectations"
    return "Missing"
