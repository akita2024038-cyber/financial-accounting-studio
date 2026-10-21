# Architecture plan

## Processing pipeline
1. Teacher uploads a solved XLSX workbook and a PDF statement.
2. A parser extracts sheets, formulas, named ranges, cell roles, and question metadata.
3. The variant engine creates a deterministic seed per version and changes only configured input cells.
4. The solver recalculates the workbook and stores a separate answer key outside the student download.
5. The grader compares normalized cell values, formula presence, journal-entry rules, debit/credit balance, and statement totals.
6. Results are written as question-level scores with error explanations.

## Accounting grading model
- Numeric cells: absolute/relative tolerance and rounding policy.
- Formula cells: expected formula pattern plus calculated result.
- Journal entries: account, debit, credit, date, and balancing validation.
- Financial statements: subtotal and total reconciliation.
- Partial credit: points split across input, calculation, and final result.
- Manual review: ambiguous formatting or unmatched account names are flagged.

## Recommended production stack
- Next.js + TypeScript frontend
- FastAPI + Python processing service using openpyxl and LibreOffice recalculation
- PostgreSQL/Supabase for users, assignments, variants, submissions, and grading results
- Object storage for source workbooks, generated files, and protected answer keys
- Background queue for workbook processing