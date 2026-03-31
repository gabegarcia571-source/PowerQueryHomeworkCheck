"""Utility functions for parsing and normalization."""

from __future__ import annotations

import re
from datetime import date, datetime, time
from pathlib import Path


def compact_text(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (value or "").lower())


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
