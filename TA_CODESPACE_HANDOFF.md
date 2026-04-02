# TA Codespace Handoff Guide

This guide is for TAs who will run this grader in their own GitHub Codespace using this repository.

## Goal

By the end of this guide, you should be able to:

- Open your own Codespace for this repo.
- Install dependencies once.
- Run grading in batch or single-student mode.
- Generate grading reports and filled feedback sheets.
- Avoid common issues with file naming and pairing.

## 1) Create Your Own Codespace

1. Open the repository on GitHub.
2. Click Code.
3. Open the Codespaces tab.
4. Click Create codespace on main.
5. Wait for the container to finish starting.

You should land in a terminal at this project root:

`/workspaces/PowerQueryHomeworkCheck`

## 2) One-Time Environment Setup

Run these commands once per new Codespace:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m unittest discover -s tests -p "test_*.py"
```

If tests pass, your environment is ready.

## 3) Daily Start Routine

Each time you return to the Codespace:

```bash
cd /workspaces/PowerQueryHomeworkCheck
source .venv/bin/activate
git pull origin main
```

## 4) Choose Your Working Folder Strategy

You can run grading in either of these ways.

### Option A: Use Default Project Folders

Use existing folders in the repo:

- `Submissions`
- `Feedback Sheets`
- `Generated Reports`

This is simplest if your files are already there.

### Option B (Recommended): Use TA-Local Folders

This avoids mixing your run artifacts with shared repository folders.

```bash
mkdir -p "TA Workspace/Submissions" "TA Workspace/Feedback Sheets" "TA Workspace/Output"
```

Then upload your own student files into:

- `TA Workspace/Submissions`
- `TA Workspace/Feedback Sheets`

Optional local-only git ignore (applies only in your Codespace):

```bash
cat >> .git/info/exclude <<'EOF'
TA Workspace/
EOF
```

## 5) Input File Rules (Important)

The grader pairs files by student key from filename.

- Workbook extensions accepted: `.xlsx`, `.xlsm`, `.xls`
- Written response extensions accepted: `.docx`, `.pdf`
- Pairing key is parsed from filename, typically the prefix before the first underscore.
- If multiple XLSX or written files exist for the same key, the latest modified file is chosen.

Practical naming example for one student key:

- `doejane_12345_67890_OscorpCleaned.xlsx`
- `doejane_12345_67891_Reflection.pdf`

Feedback sheet template matching expects names like:

- `Feedback_Sheet_Doe_Jane.docx`

Matching is case-insensitive and punctuation-insensitive after normalization.

## 6) Run Batch Grading

### Batch with default folders

```bash
python grade_hw4.py batch \
  --submissions-dir "Submissions" \
  --output-dir "Generated Reports" \
  --feedback-sheets-dir "Feedback Sheets" \
  --filled-feedback-dir "Generated Reports/Filled Feedback Sheets"
```

### Batch with TA-local folders

```bash
python grade_hw4.py batch \
  --submissions-dir "TA Workspace/Submissions" \
  --output-dir "TA Workspace/Output" \
  --feedback-sheets-dir "TA Workspace/Feedback Sheets" \
  --filled-feedback-dir "TA Workspace/Output/Filled Feedback Sheets"
```

## 7) Run a Single Student

```bash
python grade_hw4.py single \
  --xlsx "TA Workspace/Submissions/student_file.xlsx" \
  --written "TA Workspace/Submissions/student_reflection.pdf" \
  --output "TA Workspace/Output/student_grading_report.txt" \
  --feedback-sheets-dir "TA Workspace/Feedback Sheets" \
  --filled-feedback-output "TA Workspace/Output/student_feedback_sheet.docx"
```

## 8) Where Results Appear

For a batch run:

- Per-student reports: `<output-dir>/<student_key>_grading_report.txt`
- Batch summary: `<output-dir>/batch_summary.csv`
- Filled feedback sheets: `<filled-feedback-dir>/Feedback_Sheet_<Last>_<First>.docx`

## 9) Interpreting Flags and Review Outcomes

- `FOR REVIEW` means grading could not be safely completed for that section.
- If Step 0 integrity fails (missing/corrupt/mismatched files), all score fields are set to `FOR REVIEW`.
- `NEEDS MANUAL REVIEW` indicates uncertainty where manual instructor check is required.

## 10) Common Troubleshooting

### Problem: ModuleNotFoundError or import error

Fix:

```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
```

### Problem: Many students show FOR REVIEW

Checks:

1. Confirm each student has one workbook and one written file.
2. Confirm filenames produce matching student keys.
3. Confirm files are not corrupted and can open normally.

### Problem: Missing feedback templates warning

Checks:

1. Ensure template DOCX files exist in your feedback sheets directory.
2. Ensure template names follow `Feedback_Sheet_<Last>_<First>.docx`.
3. Ensure the student key in submissions matches the same student identity.

### Problem: PDF reflection content not extracted well

Notes:

- Some scanned/image-heavy PDFs may still require manual review.
- This version does not include OCR.

## 11) Before You Commit or Push

Recommended checks:

```bash
git status
```

Only commit source-code/documentation changes you intend to share.
Avoid committing private student artifacts unless your course policy explicitly allows it.
