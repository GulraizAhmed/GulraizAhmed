#!/usr/bin/env python3
"""
Import IPG Transaction Report CSV into a local SQLite database.

Source: IPG_Transaction_Report (28 Jul 2026 – 13 Sep 2026)
Each CSV row is a payment-gateway event (3DS auth step and/or acquirer auth/capture).
"""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pandas as pd

CSV_PATH = Path(
    "/home/ubuntu/.cursor/projects/workspace/uploads/"
    "IPG_Transaction_Report__28th_July_to_13th_September__2e7d.csv"
)
DB_PATH = Path("/workspace/data/ipg_transactions.db")
TABLE = "ipg_transactions"

# Human-readable meaning of common ISO/acquirer response codes seen in this feed
RESPONSE_CODE_MEANINGS = {
    "0": "Approved / Successful",
    "5": "Do not honor",
    "12": "Invalid transaction",
    "14": "Invalid card number",
    "15": "No such issuer",
    "30": "Format error",
    "39": "No credit account",
    "51": "Insufficient funds",
    "54": "Expired card",
    "57": "Transaction not permitted",
    "61": "Exceeds withdrawal limit",
    "65": "Exceeds frequency limit",
    "75": "PIN tries exceeded",
    "79": "Already reversed / lifecycle",
    "82": "Negative CAM / CVV failure",
    "83": "Unable to verify PIN / fraud suspect",
    "85": "No reason to decline (AVS / soft decline context)",
    "91": "Issuer / switch inoperative",
    "96": "System malfunction",
}


SCHEMA_SQL = f"""
CREATE TABLE IF NOT EXISTS {TABLE} (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id TEXT NOT NULL,
    order_reference TEXT NOT NULL,
    transaction_id TEXT NOT NULL,
    transaction_date TEXT NOT NULL,
    transaction_ts TEXT,                 -- ISO-8601 when parseable
    payment_method TEXT,
    authorization_code TEXT,
    acquirer_response_code TEXT,
    acquirer_response_meaning TEXT,
    payer_auth_status TEXT,
    rrn TEXT,                            -- Reference Retrieval Number (as exported)
    acquirer_batch_number TEXT,
    acquirer_id TEXT,
    is_approved INTEGER NOT NULL DEFAULT 0,  -- 1 when acquirer_response_code = '0'
    is_auth_event INTEGER NOT NULL DEFAULT 0, -- 1 when transaction_id like 'trans-%'
    source_file TEXT,
    imported_at TEXT DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_ipg_order_ref ON {TABLE}(order_reference);
CREATE INDEX IF NOT EXISTS idx_ipg_order_id ON {TABLE}(order_id);
CREATE INDEX IF NOT EXISTS idx_ipg_txn_id ON {TABLE}(transaction_id);
CREATE INDEX IF NOT EXISTS idx_ipg_txn_ts ON {TABLE}(transaction_ts);
CREATE INDEX IF NOT EXISTS idx_ipg_payment_method ON {TABLE}(payment_method);
CREATE INDEX IF NOT EXISTS idx_ipg_response ON {TABLE}(acquirer_response_code);
CREATE INDEX IF NOT EXISTS idx_ipg_acquirer ON {TABLE}(acquirer_id);
CREATE INDEX IF NOT EXISTS idx_ipg_approved ON {TABLE}(is_approved);

CREATE TABLE IF NOT EXISTS ipg_import_meta (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_file TEXT NOT NULL,
    row_count INTEGER NOT NULL,
    date_from TEXT,
    date_to TEXT,
    notes TEXT,
    imported_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS ipg_response_code_lookup (
    code TEXT PRIMARY KEY,
    meaning TEXT NOT NULL
);
"""


def parse_txn_ts(value: str) -> str | None:
    """Parse 'DD-MM-YY HH:MM' (year 20YY) into ISO timestamp."""
    value = (value or "").strip()
    if not value:
        return None
    try:
        ts = pd.to_datetime(value, format="%d-%m-%y %H:%M", errors="raise")
        return ts.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        try:
            ts = pd.to_datetime(value, dayfirst=True, errors="raise")
            return ts.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return None


def load_frame(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    df.columns = [c.strip() for c in df.columns]

    rename = {
        "Order ID": "order_id",
        "Order Reference": "order_reference",
        "Transaction ID": "transaction_id",
        "Transaction Date": "transaction_date",
        "Payment Method": "payment_method",
        "Authorization Code": "authorization_code",
        "Acquirer Response Code": "acquirer_response_code",
        "Payer Authentication -  Authentication Status": "payer_auth_status",
        "Payer Authentication - Authentication Status": "payer_auth_status",
        "Reference Retrieval Number (RRN)": "rrn",
        "Acquirer Batch Number": "acquirer_batch_number",
        "Acquirer ID": "acquirer_id",
    }
    df = df.rename(columns=rename)

    for col in [
        "order_id",
        "order_reference",
        "transaction_id",
        "transaction_date",
        "payment_method",
        "authorization_code",
        "acquirer_response_code",
        "payer_auth_status",
        "rrn",
        "acquirer_batch_number",
        "acquirer_id",
    ]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str).str.strip()

    df["transaction_ts"] = df["transaction_date"].map(parse_txn_ts)
    df["acquirer_response_meaning"] = df["acquirer_response_code"].map(
        lambda c: RESPONSE_CODE_MEANINGS.get(c, "")
    )
    df["is_approved"] = (df["acquirer_response_code"] == "0").astype(int)
    df["is_auth_event"] = df["transaction_id"].str.startswith("trans-").astype(int)
    df["source_file"] = path.name
    return df


def import_csv(csv_path: Path = CSV_PATH, db_path: Path = DB_PATH) -> dict:
    if not csv_path.exists():
        raise FileNotFoundError(csv_path)

    db_path.parent.mkdir(parents=True, exist_ok=True)
    df = load_frame(csv_path)

    conn = sqlite3.connect(db_path)
    try:
        conn.executescript(SCHEMA_SQL)
        # Fresh load for this report window
        conn.execute(f"DELETE FROM {TABLE}")
        conn.execute("DELETE FROM ipg_response_code_lookup")
        conn.executemany(
            "INSERT INTO ipg_response_code_lookup(code, meaning) VALUES (?, ?)",
            list(RESPONSE_CODE_MEANINGS.items()),
        )

        cols = [
            "order_id",
            "order_reference",
            "transaction_id",
            "transaction_date",
            "transaction_ts",
            "payment_method",
            "authorization_code",
            "acquirer_response_code",
            "acquirer_response_meaning",
            "payer_auth_status",
            "rrn",
            "acquirer_batch_number",
            "acquirer_id",
            "is_approved",
            "is_auth_event",
            "source_file",
        ]
        rows = df[cols].itertuples(index=False, name=None)
        conn.executemany(
            f"""
            INSERT INTO {TABLE} (
                order_id, order_reference, transaction_id, transaction_date, transaction_ts,
                payment_method, authorization_code, acquirer_response_code, acquirer_response_meaning,
                payer_auth_status, rrn, acquirer_batch_number, acquirer_id,
                is_approved, is_auth_event, source_file
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            rows,
        )

        date_from = df["transaction_ts"].dropna().min() if df["transaction_ts"].notna().any() else None
        date_to = df["transaction_ts"].dropna().max() if df["transaction_ts"].notna().any() else None
        notes = (
            "IPG event-level feed: rows include 3DS authentication steps "
            "(transaction_id like trans-*) and acquirer authorization/capture attempts. "
            "Approved payments typically have acquirer_response_code='0' with authorization_code. "
            "RRN values may be Excel scientific-notation mangled in the source CSV."
        )
        conn.execute(
            """
            INSERT INTO ipg_import_meta(source_file, row_count, date_from, date_to, notes)
            VALUES (?, ?, ?, ?, ?)
            """,
            (csv_path.name, len(df), date_from, date_to, notes),
        )
        conn.commit()

        summary = {
            "db_path": str(db_path),
            "rows": len(df),
            "orders": int(df["order_reference"].nunique()),
            "order_ids": int(df["order_id"].nunique()),
            "approved": int(df["is_approved"].sum()),
            "auth_events": int(df["is_auth_event"].sum()),
            "date_from": date_from,
            "date_to": date_to,
            "payment_methods": df["payment_method"].value_counts().to_dict(),
            "acquirers": df["acquirer_id"].replace("", "(blank)").value_counts().to_dict(),
        }
        return summary
    finally:
        conn.close()


def print_summary(summary: dict) -> None:
    print("IPG transactions imported")
    print(f"  database : {summary['db_path']}")
    print(f"  rows     : {summary['rows']:,}")
    print(f"  orders   : {summary['orders']:,} unique order_reference")
    print(f"  order_id : {summary['order_ids']:,} unique")
    print(f"  approved : {summary['approved']:,} (response code 0)")
    print(f"  3DS rows : {summary['auth_events']:,} (transaction_id starts with trans-)")
    print(f"  period   : {summary['date_from']} → {summary['date_to']}")
    print("  payment methods:")
    for k, v in summary["payment_methods"].items():
        print(f"    - {k}: {v:,}")
    print("  acquirers:")
    for k, v in summary["acquirers"].items():
        print(f"    - {k}: {v:,}")


if __name__ == "__main__":
    print_summary(import_csv())
