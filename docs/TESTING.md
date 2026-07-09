# Testing & Bug-Fix Report

**Environment:** Python 3.14.6, openpyxl 3.1.5

## Methodology

The reconciliation engine (`reconcile_core.py`) is tested end-to-end with realistic
Excel fixtures built with openpyxl, each matching the gateway header schemas exactly.
Every run's output workbook is opened and its cells asserted against
independently-computed expected values. The Streamlit UI (`app.py`) is not
headless-driveable and is covered by compile-checks and manual review rather than
automated tests.

## Scenario results — 9 / 9 passing

| # | Scenario | Result | Evidence |
|---|----------|--------|----------|
| 1 | Multi-gateway merchant | ✅ PASS | Exceptions 0 (was 4 before fix); matched 4 |
| 2 | Metadata preamble skipping (incl. Pelpay loader) | ✅ PASS | Header auto-found at row 4; preamble skipped |
| 3 | Mixed-currency single file | ✅ PASS | Split NGN 3,000 / USD 120; matched 4 |
| 4 | Out-of-range "Other Dates" | ✅ PASS | 2/3,000 + 2/7,000 = 4/10,000, reconciles exactly |
| 5 | Unmatched / exceptions | ✅ PASS | Both missing-side lists correct |
| 6 | `detect_date_range` | ✅ PASS | Correct (min,max); ValueError when no date column |
| 7 | Edge cases | ✅ PASS | Empty file, unfindable header, ragged rows, mixed date formats |
| 8 | MPGS gateway | ✅ PASS | Distinct schema parsed; matched 2; total 4,000 |
| 9 | Genuine miss + multi-gateway (regression) | ✅ PASS | Real unsettled ref reported exactly once |

No exceptions or tracebacks occurred across any scenario.

## BUG-1 — Cross-gateway phantom exceptions (HIGH, resolved)

**Symptom.** For any merchant settling through more than one gateway in the same
currency, every fully-matched transaction was falsely reported as "missing from
settlement." A 100%-reconciled dataset showed all transactions as discrepancies. The
bug produced incorrect output silently — no error was raised.

**Plain-English.** Pelpay is the master list of transactions; the gateway files are the
receipts of what settled. The app checked each gateway in isolation, so a transaction
settled via Cybersource was correctly matched there but wrongly reported "missing" by
the ChoicePay check — and vice-versa. Like two door scanners each calling the other's
legitimate guest a gate-crasher because they weren't sharing logs.

**Cause.** Each gateway section only knew its own gateway's settled references, so the
exception check flagged transactions settled under other gateways as missing.

**Fix.** Pool settled references per (merchant, currency) across all gateways, and dedup
genuine misses so each is reported once.

**Before → after (merchant `acme`, NGN):**

| | Before | After |
|---|--------|-------|
| Matched | 4 | 4 |
| Exceptions (grand total) | 4 / ₦10,000 | **0 / 0** |

Verified: scenario 1 exceptions 4 → 0, while genuine misses (scenario 9) still surface
exactly once.

## Coverage & caveats

- **Streamlit UI (`app.py`)** not exercised by automated tests (not headless-driveable).
  Recommendation: one manual pass in the live app with a multi-gateway merchant.
- **CSV input path** of `load_file_rows` not tested (only `.xlsx`, the stated format).
- **Multi-sheet workbooks** not tested (only the first sheet is read by design).

## Commit history for this work

| Commit | Description |
|--------|-------------|
| `9441224` | Auto-detect header row; skip metadata preamble rows in xlsx settlement files |
| `d115000` | Apply header-row detection to Pelpay loader; dedupe header finder |
| `3148f01` | Fix cross-gateway phantom exceptions (BUG-1) |
