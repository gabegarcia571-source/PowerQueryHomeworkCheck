# PowerQueryHomeworkCheck

Automated grader for COMM2003 Homework 4 submissions (Power Query cleaning + reflection questions).

## TA Handoff

For a full step-by-step guide for another TA using this repository in their own GitHub Codespace, see:

- `TA_CODESPACE_HANDOFF.md`

## What Is Implemented

- Agent A: integrity + submission pairing
- Agent B: XLSX grading for Steps 1A-1D
- Agent C: reflection grading for Q1/Q2 (DOCX and PDF converted to DOCX first)
- Agent D: score aggregation + exact report template rendering
- Parallel execution of Agent B and Agent C after integrity gate

## Setup

1. Install dependencies:

```bash
python -m pip install -r requirements.txt
```

2. Run single submission grading:

```bash
python grade_hw4.py single \
	--xlsx "Submissions/student_file.xlsx" \
	--written "Submissions/student_reflection.pdf" \
	--output "Generated Reports/student_grading_report.txt" \
	--feedback-sheets-dir "Feedback Sheets" \
	--filled-feedback-output "Generated Reports/student_feedback_sheet.docx"
```

3. Run batch grading for all discovered pairs:

```bash
python grade_hw4.py batch \
	--submissions-dir "Submissions" \
	--output-dir "Generated Reports" \
	--feedback-sheets-dir "Feedback Sheets" \
	--filled-feedback-dir "Generated Reports/Filled Feedback Sheets"
```

## Outputs

- Per-student report files: `Generated Reports/<student_key>_grading_report.txt`
- Batch summary CSV: `Generated Reports/batch_summary.csv`
- Populated feedback sheets: `Generated Reports/Filled Feedback Sheets/Feedback_Sheet_<Last>_<First>.docx`

## Quick Validation

Run smoke tests:

```bash
python -m unittest discover -s tests -p "test_*.py"
```

## Notes

- Conservative policy is enforced: uncertain checks raise instructor review flags.
- If Step 0 integrity fails, all score fields are set to `FOR REVIEW` and grading short-circuits.
- PDF reflections are converted to DOCX in temporary directories before scoring.
- OCR is not included in this version. Image-only/scanned PDFs may still require manual review.