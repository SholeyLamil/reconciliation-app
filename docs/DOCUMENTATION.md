# Gateway ↔ Pelpay Settlement Reconciliation — Documentation

| | |
|---|---|
| **Application** | Reconciliation App (Streamlit) |
| **Deployment** | Streamlit Community Cloud (`share.streamlit.io`) |
| **Entry point** | `app.py` |
| **Language / runtime** | Python (Streamlit UI + openpyxl engine) |

---

## 1. Overview

### 1.1 Purpose
The application reconciles **payment-gateway settlement files** against a **Pelpay
transaction export**. It confirms that every successful transaction recorded by Pelpay
has a corresponding settlement record from the gateway that processed it, and surfaces
any discrepancies — transactions missing on either side, and amount differences.

It replaces a manual spreadsheet process with a single upload-and-run workflow that
produces a formatted, multi-sheet Excel reconciliation workbook.

### 1.2 Who it's for
Finance / operations staff performing periodic settlement reconciliation across
multiple payment processors. No technical knowledge is required to operate it.

### 1.3 Supported gateways & currencies
| Gateway | Approved status | ref / amount / date fields |
|---------|-----------------|----------------------------|
| **Cybersource** | `BATCHED` | `merchant_ref_number` / `amount` / `batch_date` |
| **ChoicePay** | `Captured` | `Order Reference` / `Order Amount (amount only)` / `Order Date` |
| **MPGS** | `Successful` | `Processor Reference` / `Settlement amount` / `Transaction Date` |

Currencies: **NGN** and **USD**.

---

## 2. Architecture

The application separates a thin presentation layer from a self-contained engine.

```
app.py  — Streamlit UI
  • File upload & auto-classification
  • Header-row / currency / date detection
  • "Files Required" checklist + detection table
  • Validation, Run button, result download
        │  calls run(...)
        ▼
reconcile_core.py  — Engine
  • Schema definitions (DEFAULT_SCHEMAS)
  • File loading + header detection
  • Matching & exception computation
  • Excel workbook builders (5 sheet types)
```

**Design principle:** all business logic lives in `reconcile_core.py` and is
importable/testable in isolation. `app.py` is glue — file handling and display only.

---

## 3. Processing pipeline

1. **Upload** — Pelpay export + one or more gateway settlement files, `.xlsx` or `.csv`.
2. **Classify** — headers inspected to determine file type (`classify_file()`).
3. **Header-row detection** — `_find_header_row()` skips metadata / preamble rows
   (report titles, near-empty rows) to locate the true header row. Applied to both
   settlement files and the Pelpay file.
4. **Currency detection** — `read_currencies_from_file()` scans the currency column;
   a single file may contain both NGN and USD rows and is split accordingly. Falls
   back to detecting currency from the filename.
5. **Date-range detection** — `detect_date_range()` derives the `(min, max)` settlement
   date range automatically from the files (no manual date entry).
6. **Validation** — a "Files Required" checklist confirms each gateway/currency
   combination is present; problems are reported before running.
7. **Reconcile** — `run()` loads Pelpay, loads & merges settlements, matches records,
   computes exceptions, and writes the output workbook.
8. **Download** — the formatted `.xlsx` is offered for download.

---

## 4. Reconciliation logic

### 4.1 Matching
Settlement rows are grouped by **(merchant, currency, gateway)** and merged (dedup by
reference). Each settlement reference is matched against the Pelpay export by
normalised reference, merchant, and currency. A match records the amount difference
(settlement − Pelpay).

### 4.2 Exceptions
- **Missing from Settlement** — a successful, in-range Pelpay transaction with no
  settlement record in **any** of that merchant's gateway files.
- **Missing from Pelpay** — a settlement record with no corresponding Pelpay
  transaction.

> **Correctness property:** the "settled" reference set is pooled **across all
> gateways** for each (merchant, currency). A transaction settled via one gateway is
> therefore *not* falsely flagged as missing by another gateway.

### 4.3 "Other Dates" handling
Settlement rows whose dates fall **outside** the requested range are aggregated into a
dedicated **"Other Dates"** row per section rather than dropped, so summary totals
always reconcile to the full file contents.

---

## 5. Output workbook

`run()` produces a single formatted `.xlsx` with these sheets:

| Sheet | Contents |
|-------|----------|
| **SUMMARY** | Per-merchant, per-currency daily breakdown: settlement rows, matched rows, totals, differences, missing counts, coverage. Includes "Other Dates" rows, per-currency totals, grand totals, and a "Successful Pelpay Transaction Summary" block. |
| **All received Transactions** | All successful in-range Pelpay transactions by merchant/currency. |
| **Missing from Settlement** | Successful Pelpay transactions with no settlement record. |
| **Missing from Pelpay** | Settlement records with no Pelpay transaction. |
| **Per-settlement sheets** | One detail sheet per merchant/currency/gateway section. |

All sheets receive number formatting and styling via `apply_formatting()`.

---

## 6. Development & testing

The engine is tested end-to-end with realistic Excel fixtures that match the gateway
schemas exactly; each run's output workbook is opened and asserted against
independently-computed expected values. The current suite covers 9 scenarios (happy
path, metadata preamble skipping, mixed-currency files, out-of-range "Other Dates",
unmatched/exceptions, date-range detection, edge cases, the MPGS gateway, and a
multi-gateway regression) — all passing.

**Not covered by automated tests:** the Streamlit UI layer (not headless-driveable —
covered by manual review), the CSV input path, and multi-sheet workbooks.

See [`TESTING.md`](TESTING.md) for the full test report and bug history.

---

## 7. Deployment

- **Source control:** GitHub. Pushes to `main` auto-trigger a Streamlit redeploy.
- **Hosting:** Streamlit Community Cloud — entry point `app.py`, branch `main`.
- **Dependencies** (`requirements.txt`): `streamlit>=1.59.0`, `openpyxl>=3.0.0`
  (all other imports are Python standard library).
- **Operational note:** if a redeploy doesn't appear automatically, use
  *Manage app → Reboot* on Streamlit Cloud, then hard-refresh the browser tab.

---

## 8. Known limitations & future work

- New processors require a schema addition in `DEFAULT_SCHEMAS`.
- CSV inputs assume the header is the first row (no preamble skipping on CSV).
- No automated UI tests; a lightweight Streamlit end-to-end check would close the gap.
- The public app URL is reachable by anyone with the link — restrict viewers via the
  app's Sharing settings if the data is sensitive.

---

## Appendix — Gateway schema reference

```
CYBERSOURCE : ref=merchant_ref_number  amount=amount
              status=status (BATCHED)  date=batch_date
CHOICEPAY   : ref=Order Reference      amount=Order Amount (amount only)
              status=Order Status (Captured)  date=Order Date
MPGS        : ref=Processor Reference  amount=Settlement amount
              status=Transaction Status (Successful)  date=Transaction Date
```
