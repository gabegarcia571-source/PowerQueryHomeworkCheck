"""Agent A: file pairing and integrity checks."""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from grader.constants import FOR_REVIEW
from grader.models import IntegrityResult
from grader.utils import extract_student_display_name, extract_student_key

XLSX_EXTS = {".xlsx", ".xlsm", ".xls"}
WRITTEN_EXTS = {".docx", ".pdf"}


def discover_pairs(submissions_dir: Path) -> dict[str, dict[str, Path | None]]:
    """Discover one xlsx and one written file per student key.

    If multiple files exist for a type, picks latest modified file deterministically.
    """
    by_key: dict[str, dict[str, list[Path]]] = {}

    for path in sorted(submissions_dir.iterdir()):
        if not path.is_file():
            continue
        key = extract_student_key(path.name)
        if not key:
            continue
        item = by_key.setdefault(key, {"xlsx": [], "written": []})
        suffix = path.suffix.lower()
        if suffix in XLSX_EXTS:
            item["xlsx"].append(path)
        elif suffix in WRITTEN_EXTS:
            item["written"].append(path)

    resolved: dict[str, dict[str, Path | None]] = {}
    for key, files in by_key.items():
        xlsx = _latest(files["xlsx"])
        written = _latest(files["written"])
        resolved[key] = {"xlsx": xlsx, "written": written}

    return resolved


def run_integrity_agent(student_key: str, xlsx_path: Path | None, written_path: Path | None) -> IntegrityResult:
    reasons: list[str] = []

    display_from = xlsx_path.name if xlsx_path else (written_path.name if written_path else student_key)
    display_name = extract_student_display_name(display_from, student_key)

    xlsx_key = extract_student_key(xlsx_path.name) if xlsx_path else None
    written_key = extract_student_key(written_path.name) if written_path else None

    if xlsx_path is None:
        reasons.append("XLSX file missing.")
    if written_path is None:
        reasons.append("Written response file missing.")

    if xlsx_path is not None:
        try:
            wb = load_workbook(xlsx_path, read_only=True, data_only=True)
            wb.close()
        except Exception as exc:  # noqa: BLE001
            reasons.append(f"XLSX failed to open: {exc}")

    if written_path is not None:
        try:
            # Read a small chunk to verify file is openable.
            with written_path.open("rb") as handle:
                _ = handle.read(16)
        except Exception as exc:  # noqa: BLE001
            reasons.append(f"Written response failed to open: {exc}")

    if xlsx_key and written_key and xlsx_key != written_key:
        reasons.append(
            f"Submission pair mismatch: XLSX key '{xlsx_key}' does not match written key '{written_key}'."
        )

    return IntegrityResult(
        student_key=student_key,
        student_display_name=display_name,
        xlsx_path=xlsx_path,
        written_path=written_path,
        passed=len(reasons) == 0,
        reasons=reasons,
    )


def for_review_integrity_result(student_key: str, reason: str) -> IntegrityResult:
    return IntegrityResult(
        student_key=student_key,
        student_display_name=student_key.title(),
        xlsx_path=None,
        written_path=None,
        passed=False,
        reasons=[reason or FOR_REVIEW],
    )


def _latest(paths: list[Path]) -> Path | None:
    if not paths:
        return None
    return sorted(paths, key=lambda p: (p.stat().st_mtime, p.name))[-1]
