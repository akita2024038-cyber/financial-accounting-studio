# Financial Accounting Studio

A platform for generating, solving, and grading Financial Accounting assignments in Excel.

## MVP scope

- Teacher workspace for assignment templates, variants, answer keys, and grading
- Student submission workflow
- Accounting-aware grading rules: calculations, journal entries, trial balances, and statements
- Separate answer-key storage boundary
- Import/export integration points for Excel and PDF processing

## Stack

The initial UI is a zero-dependency browser MVP so it can be opened directly from GitHub Pages or any static host. The domain model is designed to be extended with a Next.js/FastAPI/Supabase backend.

## Current demo

Open `index.html` to explore the teacher dashboard. The demo includes sample assignment data and a grading preview.

## Roadmap

1. Excel/PDF upload and parsing
2. Variant generation with deterministic seeds
3. Protected answer-key storage
4. Excel formula and workbook grading
5. Authentication, student submissions, and reports
6. Production deployment
