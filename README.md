# Pelpay ↔ Gateway Settlement Reconciliation

Reconciles Pelpay transactions against settlement files from Cybersource, ChoicePay, and MPGS. Produces an .xlsx workbook with the settlement file as the source of truth.

## Files

| File | Purpose |
|------|---------|
| `reconcile_core.py` | Core engine — importable `run()` function with no I/O globals |
| `app.py` | Streamlit UI — upload files, run reconciliation, download results |
| `reconciliation_workbook_format.md` | Detailed spec of the output workbook layout |

## Documentation

- [`docs/DOCUMENTATION.md`](docs/DOCUMENTATION.md) — full technical documentation (architecture, pipeline, reconciliation logic, output workbook, deployment)
- [`docs/TESTING.md`](docs/TESTING.md) — test report, coverage, and bug history
- [`reconciliation_workbook_format.md`](reconciliation_workbook_format.md) — detailed output-workbook layout spec

## Quick Start

```bash
pip install -r requirements.txt
streamlit run app.py
```

## API Usage

```python
from reconcile_core import run

result = run(
    pelpay_path="pelpay.xlsx",
    settlement_files=[("CYBERSOURCE", "NGN", "cybersource_ngn.xlsx")],
    settlement_date="2026-07-05",
    output_path="Reconciliation.xlsx",
)
# result: { "settle_rows": ..., "matched": ..., "exceptions": ..., "sheets": [...] }
```

## Supported Gateways

| Gateway | ref field | date field | amount field |
|---------|-----------|------------|--------------|
| Cybersource | `merchant_ref_number` | `batch_date` | `amount` |
| ChoicePay | `Order Reference` | `Order Date` | `Order Amount (amount only)` |
| MPGS | `Processor Reference` | `Transaction Date` | `Settlement amount` |

## Output Workbook

- **SUMMARY** — per-day, per-merchant breakdown: settlement rows, matched rows, totals, differences, missing transactions
- **All Received Transactions** — full Pelpay dump
- **Missing from Settlement** — Pelpay transactions not in any settlement file
- **Missing from Pelpay** — settlement rows with no Pelpay match
- **Settlement Sheets** — one per merchant/currency/gateway
