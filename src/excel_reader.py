"""STEP 1-4: load the workbook, read the Working sheet, validate columns, build investor records."""
import logging

import pandas as pd

import config
from utils import clean_str, normalize, to_float

log = logging.getLogger("drawdown")


def read_working_data(path=None):
    """Read the Working sheet into a DataFrame (header row = row 2). Only real investor rows are kept
    (rows with a Folio No.) so the trailing totals row is ignored."""
    path = path or config.INPUT_WORKBOOK
    df = pd.read_excel(path, sheet_name=config.WORKING_SHEET, header=config.WORKING_HEADER_ROW - 1)
    # remember the Excel row number of each record (header is row 2, data starts row 3)
    df["_excel_row"] = df.index + config.WORKING_HEADER_ROW + 1
    folio_header = _find_header(df, config.FIELD_MAPPING["folio_no"])
    df = df[df[folio_header].notna() & (df[folio_header].astype(str).str.strip() != "")].copy()
    log.info("Workbook loaded: %s | Working rows with a folio: %d", path, len(df))
    return df.reset_index(drop=True)


def _find_header(df, wanted):
    """Find the real column name for a wanted header, ignoring spaces / newlines / case."""
    wanted_n = normalize(wanted)
    for col in df.columns:
        if normalize(col) == wanted_n:
            return col
    return None


def build_column_map(df):
    """logical field -> real column name. Raises if a REQUIRED field is missing."""
    colmap, missing = {}, []
    for field, header in config.FIELD_MAPPING.items():
        real = _find_header(df, header)
        if real is None:
            if field in config.REQUIRED_FIELDS:
                missing.append(f"{field} ('{header}')")
        else:
            colmap[field] = real
    if missing:
        raise ValueError("Required Working columns not found: " + ", ".join(missing))
    colmap["_paid_columns"] = [c for c in df.columns if config.PAID_COLUMN_REGEX.search(str(c))]
    if not colmap["_paid_columns"]:
        raise ValueError("No '<n> Drawdown Amount Paid' columns found in Working")
    return colmap


def get_investor_record(row, colmap):
    """Turn one DataFrame row into a plain dict keyed by logical field names."""
    rec = {"excel_row": int(row["_excel_row"])}
    text_fields = ("class", "folio_no", "notice_format", "investor_name", "address", "email")
    for field in text_fields:
        rec[field] = clean_str(row[colmap[field]]) if field in colmap else ""
    num_fields = ("capital_commitment", "capital_contributed_wb", "drawdown_called", "portfolio_investment",
                  "adjustment", "management_fee", "fund_expenses", "pending_previous",
                  "excess_contribution", "amount_payable_wb", "percentage_wb")
    for field in num_fields:
        rec[field] = to_float(row[colmap[field]]) if field in colmap else 0.0
    d = row[colmap["drawdown_date"]]
    rec["drawdown_date"] = pd.Timestamp(d).to_pydatetime() if pd.notna(d) else None
    rec["paid_amounts"] = {c: to_float(row[c]) for c in colmap["_paid_columns"]}
    return rec
