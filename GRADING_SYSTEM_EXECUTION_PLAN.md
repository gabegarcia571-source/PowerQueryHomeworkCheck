# COMM2003 HW4 Grading System Execution Plan

## 1) Scope and Success Criteria

This system grades one student submission using:

- One XLSX workbook
- One written response file (DOCX or PDF)

It must follow the grading prompt exactly:

- Conservative scoring (when uncertain, do not award points)
- Step 0 integrity gate before all other checks
- Exact report format output
- Explicit instructor review flags for uncertainty

## 2) Requirements Traceability (Prompt -> System Behavior)

### Step 0 Integrity Check

Required by prompt:

- Attempt to open both files.
- If either missing/unreadable/corrupted, all scores are `FOR REVIEW`.
- Stop grading and output review report with reason.

System requirement:

- `IntegrityAgent` performs hard gate and blocks all downstream agents on failure.

### Step 1A Worksheets Present (1.0)

Required by prompt:

- 7 required worksheet intents must be present.
- Fuzzy/semantic matching allowed.
- Score is binary: 1 or 0.

System requirement:

- `WorkbookStructureAgent` maps sheet names to required intents using normalized tokens and alias dictionary.

### Step 1B Cleaned Dataset (4.0)

Required by prompt:

- Evaluate the cleaned transactions sheet (not merged customer sheet).
- Validate 13 required columns (order not required).
- Deduct 0.5 per error type, minimum 0.
- No double counting same root issue.
- If confidence is insufficient, do not deduct; flag instructor review.

System requirement:

- `CleanedDataAgent` runs criterion-level checks and records evidence + confidence.
- Deduction engine enforces one deduction per root issue class.

### Step 1C and 1D Pivot Reports (0.5 each)

Required by prompt:

- Real pivot object verification via `ws._pivots`.
- Correct dimensions and aggregation field intent.
- If pivot object cannot be verified programmatically, mark manual review.

System requirement:

- `PivotAgent` checks object existence and field intent; emits `NEEDS MANUAL REVIEW` when verification cannot be trusted.

### Step 2 Written Reflection (Q1/Q2, 2 points each)

Required by prompt:

- Score bins are fixed: 2.0 / 1.25 / 0.75 / 0.0.
- Missing answer is 0.
- Rationales must be brief and no direct quotes.

System requirement:

- `ReflectionAgent` extracts text, isolates Q1/Q2, applies rubric rubric-by-rubric, and emits label + score + short rationale.

### Step 3 and Step 4 Finalization

Required by prompt:

- Total score out of 10.
- Emit report in exact provided layout.

System requirement:

- `ReportAgent` builds exact template text and validates all required sections before output.

## 3) Clarifications Resolved from Prompt and Resource Files

Sources consulted:

- `COMM2003_HW4_Grading_Prompt_v4.md` (provided in user attachment context)
- `Resource/Assignment 4 - Grading Guidelines.docx`
- `Resource/Homework 4 - Instructor's Guide.docx`

Resolved decisions:

- For Step 1B, run checks on `AllTransactions`-equivalent sheet, not merged `AllTransactionsAndCustomers`.
- Sheet names are intent-based, not exact literal strings.
- Known duplicate IDs (`T0010`, `T0026`) are a direct check target.
- `TransactionID` must be `T####`; any `TR####` remaining is a criterion failure.
- `TransactionType` abbreviations (`FEE/DEP/WIT/TRA/PAY`) must be normalized to title-case full words.
- `Merchant` must not include city/state fragments; those belong in separate `City` and `State` columns.
- Pivot verification must prefer real object inspection (`ws._pivots`) over visual table inference.
- If certainty is insufficient for a criterion, no deduction is applied; instructor review flag is added.

## 4) What We Have in This Workspace

Validated repository assets:

- `Submissions/` contains student XLSX + DOCX/PDF files.
- `Resource/` contains grading guideline and instructor guide docs.
- `Oscorp-answerkey.xlsx` exists for sanity checks and reference alignment.
- `Feedback Sheets/` exists for downstream output workflows.

Inventory result snapshot from automated scan:

- Parsed student keys: 66
- Paired keys with both XLSX and written file: 66
- Only-XLSX keys: 0
- Only-written keys: 0

## 5) MECE Agent Assignment (Senior Engineering Split)

### Agent A: Integrity and Pairing Agent

Owns:

- File pair discovery (xlsx + docx/pdf)
- Open/readability checks
- Student identity extraction from filename
- Step 0 gate result

Output contract:

- `student_key`
- `xlsx_path`
- `written_path`
- `integrity_passed`
- `failure_reasons[]`

### Agent B: XLSX Rubric Agent

Owns:

- Step 1A worksheet intent coverage
- Step 1B cleaned dataset checks
- Step 1C/1D pivot checks including object verification

Output contract:

- `score_1a`
- `score_1b`
- `score_1c`
- `score_1d`
- `errors_1b[]`
- `notes_1a/1c/1d`
- `review_flags[]`

### Agent C: Reflection Rubric Agent

Owns:

- Written file text extraction (docx/pdf)
- Q1 and Q2 identification
- Rubric label selection and scoring
- Short rationale generation per question

Output contract:

- `score_q1`, `label_q1`, `rationale_q1`
- `score_q2`, `label_q2`, `rationale_q2`
- `review_flags[]`

### Agent D: Report and Adjudication Agent

Owns:

- Merge outputs from A/B/C
- Compute total out of 10
- Emit exact report template
- Enforce `FOR REVIEW` behavior when Step 0 fails

Output contract:

- Final report text
- Structured JSON companion for auditing (optional)

## 6) Parallelization Strategy and Dependencies

Execution graph:

1. Agent A executes first (hard dependency).
2. If Step 0 passes, Agent B and Agent C run in parallel.
3. Agent D runs after B and C complete.

This is mutually exclusive and collectively exhaustive:

- No overlap in ownership boundaries.
- Every rubric item is owned by exactly one grading agent.
- Aggregation logic is isolated from grading logic.

## 7) Quality Gates and Best-Practice Controls

- Gate G0: File integrity pass required before scoring.
- Gate G1: Evidence-backed deductions only.
- Gate G2: Anti-double-count validation in Step 1B.
- Gate G3: Pivot object confidence check (`ws._pivots`) with manual review fallback.
- Gate G4: Report schema validation against exact required template.

Auditability best practices:

- Save criterion-by-criterion evidence and reason for each deduction.
- Separate scoring logic from report formatting logic.
- Keep deterministic outputs for same inputs.

## 8) Open Questions and Status

Questions resolved directly by prompt/resources:

- Rubric scale, scoring bins, and deduction model: resolved.
- Required worksheet/content expectations: resolved.
- Pivot verification method and manual-review fallback: resolved.

No blocking specification gaps remain for implementation.

Non-blocking operational policy choices (implementation defaults recommended):

- Filename parsing fallback behavior when malformed names appear.
- Batch output location and file naming convention for generated reports.
- Whether to emit both text report and machine-readable JSON.

## 9) Immediate Next Implementation Steps

1. Build Agent A first (integrity gate, pairing, metadata extraction).
2. Build Agent B criterion checks with confidence tracking.
3. Build Agent C reflection extraction/scoring for docx/pdf.
4. Build Agent D report renderer with strict template validation.
5. Run pilot on 3-5 students and review flagged items for calibration.
