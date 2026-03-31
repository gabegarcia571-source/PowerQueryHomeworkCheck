# PowerQueryHomeworkCheck

Automated grader for COMM2003 Homework 4 submissions (Power Query cleaning + reflection questions).

## What Is Implemented

- Agent A: integrity + submission pairing
- Agent B: XLSX grading for Steps 1A-1D
- Agent C: reflection grading for Q1/Q2 (DOCX/PDF)
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
	--output "Generated Reports/student_grading_report.txt"
```

3. Run batch grading for all discovered pairs:

```bash
python grade_hw4.py batch \
	--submissions-dir "Submissions" \
	--output-dir "Generated Reports"
```

## Outputs

- Per-student report files: `Generated Reports/<student_key>_grading_report.txt`
- Batch summary CSV: `Generated Reports/batch_summary.csv`

## Quick Validation

Run smoke tests:

```bash
python -m unittest discover -s tests -p "test_*.py"
```

## Notes

- Conservative policy is enforced: uncertain checks raise instructor review flags.
- If Step 0 integrity fails, all score fields are set to `FOR REVIEW` and grading short-circuits.