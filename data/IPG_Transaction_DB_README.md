# IPG Transaction Database

## Source
`IPG_Transaction_Report__28th_July_to_13th_September__2e7d.csv`  
Period: **2026-07-29 → 2026-09-13** (event timestamps)

## What this data is
Internet Payment Gateway (IPG) **event-level** transaction log.  
One payment attempt is usually represented by **multiple rows**:

1. **3DS / payer authentication** events (`transaction_id` like `trans-389`)
2. **Acquirer authorization / capture** attempts (numeric `transaction_id`, with response code / RRN / batch)

## Database
- Path: `data/ipg_transactions.db` (SQLite)
- Main table: `ipg_transactions`
- Lookup: `ipg_response_code_lookup`
- Import audit: `ipg_import_meta`

Reload anytime:

```bash
python3 import_ipg_transactions.py
```

## Key columns
| Column | Meaning |
|--------|---------|
| `order_id` | Gateway order identifier (often `order_reference` + suffix) |
| `order_reference` | Business / merchant order reference |
| `transaction_id` | Event id within the order (`trans-*` = auth step; numeric = acquirer attempt) |
| `transaction_date` | Original `DD-MM-YY HH:MM` string from export |
| `transaction_ts` | Parsed ISO timestamp |
| `payment_method` | Visa / Mastercard / UnionPay SecurePay |
| `authorization_code` | Bank auth code (present on successful approvals) |
| `acquirer_response_code` | Acquirer/ISO response (`0` = approved) |
| `payer_auth_status` | 3DS authentication status |
| `rrn` | Retrieval Reference Number (may be Excel scientific-notation mangled) |
| `acquirer_batch_number` | Acquirer settlement batch |
| `acquirer_id` | e.g. `BANK_PUNJAB_S2I`, `UNIONPAY_SECUREPAY_BOP` |
| `is_approved` | `1` when response code is `0` |
| `is_auth_event` | `1` when `transaction_id` starts with `trans-` |

## Useful queries

```sql
-- Approved payments
SELECT * FROM ipg_transactions WHERE is_approved = 1;

-- Orders with at least one approval
SELECT order_reference, COUNT(*) AS events,
       SUM(is_approved) AS approvals
FROM ipg_transactions
GROUP BY order_reference
HAVING approvals > 0;

-- Response-code breakdown
SELECT acquirer_response_code, acquirer_response_meaning, COUNT(*)
FROM ipg_transactions
WHERE acquirer_response_code <> ''
GROUP BY 1, 2
ORDER BY 3 DESC;
```
