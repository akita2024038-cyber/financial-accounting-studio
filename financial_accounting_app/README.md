# Financial Accounting Studio

Real FastAPI browser application for generating consistent `.xlsx` Financial Accounting assignments and solutions, publishing variants to students, automatically grading submitted workbooks, and exporting reports. UI language: Russian. Works in Google Chrome.

## Local installation

Python 3.11–3.13, LibreOffice Calc and its CLI (`libreoffice`), and pip are required. To start:

```bash
cd financial_accounting_app
python -m venv .venv
source .venv/bin/activate     # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env           # Windows: copy .env.example .env
# Edit .env and set APP_SECRET to 32+ random chars, SETUP_TOKEN to another secret
python run.py
```

Open **http://127.0.0.1:8000**. On the first visit choose *Первичная настройка* and supply `SETUP_TOKEN` to create a teacher account; then sign in. Google Chrome runs only the frontend; the Python server and LibreOffice run on your own computer/server. Keep the terminal running.

Run tests from the project root:

```bash
pytest -v tests
```

### Docker Compose (PostgreSQL)

Copy `.env.example` to `.env`, set strong `APP_SECRET`, `SETUP_TOKEN`, `DB_PASSWORD`, then:

```bash
docker compose up --build -d
```

Open http://127.0.0.1:8000. Compose binds only to localhost; put a HTTPS reverse proxy with access controls in front if deploying beyond local network. Set `COOKIE_SECURE=1` behind HTTPS. Configure rate limiting at the reverse proxy.

## Database migrations

On a new installation `run.py` automatically executes `alembic upgrade head`. To apply manually:

```bash
alembic upgrade head
alembic current
```

## Teacher workflow

1. Log in; upload **task.xlsx** (blank answers) and **key.xlsx** (formulas and correct answers). Identical sheet names/order are required.
2. Review suggested numerical input cells and confirm their min / max / step. Select only independent Financial Accounting variables; if assets/liabilities or other accounting identities must balance, specify constraint rules.
3. Review each answer mapping (student sheet/cell, key sheet/cell), grading weights, formula requirement and tolerance. Add rules for intermediate cells.
4. Save configuration, enter a variant count and generate. Download each task, solution, combined answer workbook, tasks-only ZIP, keys-only ZIP, or complete ZIP. Answers are never provided to student routes.
5. Publish the assignment and share the displayed student link. Each student enters Student ID and group, receives one variant, uploads a completed XLSX and sees their own results.
6. On Grading, inspect class records, upload individual XLSX workbooks or ZIP with files `StudentID__work.xlsx`, override ambiguous scores, and export XLSX or CSV.

### Accounting-specific constraints

In the *Генерация* tab, constraints are JSON objects evaluated **after actual formula recalculation** (all must be satisfied). Example:

```json
[{"sheet":"Balance","left":"B10","right":"C10","operator":"eq","tolerance":0.01}]
```

Supported operators: `eq`, `le`, `ge`. Right operand: a cell address (`right`) or a number (`value`). A failed constraint triggers re-sampling up to 200 attempts per variant; if impossible, the generation fails clearly rather than outputting wrong financial statements.

**Caution**: Automatic identification of semantically independent accounting numbers is undecidable from formatting alone. Approval of editable cells / domain constraints is mandatory. Numerical constants embedded inside prose, Excel figures/charts, conditional formulas, external data connections, and complex dynamic arrays are NOT rewritten. Currency formatting is preserved in task and key OOXML except for cache updates. If a workbook has unsupported formulas or dependencies, the app rejects it, rather than silently inventing answers.

## Grading behavior

Exact formulas and PySheetGrader-supported symbolic arithmetic comparisons receive full formula credit when numeric results agree. LibreOffice recalculates workbooks; correct numeric results with unproved alternative formulas are flagged **manual review**, not automatically asserted equivalent. Missing formulas receive partial credit only if the numeric result matches (25%); a valid expected formula with wrong result gets 50%; alternate reference strategies with matching results get provisional 50–75% pending teacher review. Weights and tolerances vary by answer cell; teacher can override and record reasons. Spreadsheet errors and altered template input cells are detected. Errors are shown per cell in Russian; the student's original submission is retained unchanged. The separate annotated workbook includes colored cells, notes, and a summary sheet.

### Structural + privacy considerations

Files: `.xlsx` only, 8MB/file, up to 45,000 cells/workbook. Password-protected workbooks, macro-enabled workbooks, external Excel links, embedded objects, and unsupported LibreOffice expressions are rejected. Browser upload validation is backed by server-side checks of OOXML and ZIP sizes; no arbitrary ZIP paths are extracted. Teacher cookie: HttpOnly signed JWT, samesite strict, optional Secure. Student result access: high-entropy secret token stored hashed in database; token is held in browser sessionStorage and **lost on session end**. ID/group alone do not recover access; ask teacher if lost. Without an institutional identity provider, registration is **first-claim** by Student ID; it is not strong proof of identity and should not be deployed in high-stakes examinations without external authentication. The original apps are not security sandboxes; isolate the server/container and restrict upload volume. For public deployment add reverse-proxy rate limits and malware scanning.

### Architecture

- FastAPI + static HTML/CSS/JS: no Node build step, responsive Chrome interface; deliberately chosen over React/TS for a directly executable self-contained app.
- SQLAlchemy database: SQLite for local use; Docker Compose starts PostgreSQL for deployments. Alembic `0001_initial` creates the initial schema; `run.py` and Docker run `alembic upgrade head` on startup.
- openpyxl: retains workbook structure and writes changed literal inputs.
- LibreOffice Calc (headless, private temporary profile): calculates formula caches. Cached results are injected into original OOXML to avoid LibreOffice rewriting the workbook itself. Formula errors stop generation. Excel 365-specific formulas may not be compatible.
- PySheetGrader: unmodified original BSD-3 source vendored in `third_party/pysheetgrader-core-main`; formula parser used for symbolic comparison when supported. Special `SheetGradingOrder` and comment rubrics were not forced on teachers; this app implements JSON-based auto rubrics instead.
- ExcelLabGrader: unmodified original MIT source retained in `third_party/ExcelLabGrader-main`; adapted its completeness/status + CSV reporting pattern, not its format-specific grading routine (which explicitly skips XLSX).

## Project structure

```
app/db.py          SQLAlchemy tables and DB session
app/security.py    teacher password, cookies, student tokens
app/excel.py       safe XLSX validation, field detection, generation, LibreOffice calculation
app/grading.py     grading, PySheetGrader formula parser, detailed Excel/CSV reports
app/main.py        application API, teacher and student routes
app/static/*       responsive frontend (no build tools)
tests/*            integration and formula tests
migrations/*       Alembic initial database revision
examples/*         sample workbooks and README usage
third_party/*      original source code, LICENSE and readmes (unmodified)
Dockerfile, docker-compose.yml, run.py, requirements.txt, .env.example
```

### Known limitations

- Cannot infer all valid financial/accounting invariants; teacher must select mutable inputs and define constraints.
- Equivalence of arbitrary Excel formulas cannot be formally proved; uncertain equivalent formulas are escalated for manual review.
- Advanced Excel features (pivot tables, macros, structured table formulas, cross-workbook links, charts with numeric references) are not certified for generated documents. Excel `.xls` and `.xlsm` are unsupported.
- No SSO or strong identity verification; Student ID claim must be corroborated externally in high-stakes settings.
- Synchronous LibreOffice calculation is suited to modest class sizes; for heavy concurrent use add a job queue and resource-limited isolated formula service.
- The initial Alembic migration is included; future schema changes need new migrations. A job queue, antivirus scanning, and reverse-proxy rate limiting are not included.
- Per-variant question selection/reordering is not supported: the strict same-questions requirement is prioritized.
- When an answer is a hardcoded number in the solution key (not a formula), changing inputs cannot regenerate that answer; use formulas in keys or add independent input validation. 

## Licensing

Bundled code stays in `third_party/`; both `LICENSE` files are retained. BSD-3-Clause (PySheetGrader) and MIT (ExcelLabGrader). We did not install their frozen historic requirement pins, which conflict with a modern FastAPI/Python environment; only current compatible libraries are installed.
