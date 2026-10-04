"""Small shared helpers."""
import logging
import re
from datetime import datetime

import config


def setup_logging():
    config.LOG_DIR.mkdir(parents=True, exist_ok=True)
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    for h in list(root.handlers):
        root.removeHandler(h)
    fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s")
    fh = logging.FileHandler(config.LOG_FILE, mode="w", encoding="utf-8")
    fh.setFormatter(fmt)
    root.addHandler(fh)
    sh = logging.StreamHandler()
    sh.setFormatter(logging.Formatter("%(message)s"))
    sh.setLevel(logging.WARNING)       # keep the console clean; progress is printed by main.py
    root.addHandler(sh)
    return logging.getLogger("drawdown")


def normalize(text):
    """Trim, collapse whitespace/newlines, lower-case - used to match Working headers robustly."""
    return re.sub(r"\s+", " ", str(text)).strip().lower()


def clean_str(value):
    """Return a trimmed string, or '' for blank/NaN."""
    if value is None:
        return ""
    try:
        import pandas as pd
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def to_float(value):
    """Blank / NaN -> 0.0, otherwise float."""
    if value is None or value == "":
        return 0.0
    try:
        import pandas as pd
        if pd.isna(value):
            return 0.0
    except Exception:
        pass
    return float(value)


def format_date(d):
    return config.DATE_FORMAT.format(d=d)


def fmt_amount(x):
    return f"{x:,.2f}"


def now_stamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def notice_basename(folio, fmt_sheet):
    return f"{folio}_{config.FILE_LABEL[fmt_sheet]}_Drawdown_Notice"


def email_basename(folio, fmt_sheet):
    return f"{folio}_{config.FILE_LABEL[fmt_sheet]}_Drawdown_Email"


from contextlib import contextmanager          # noqa: E402


@contextmanager
def use_workbook(path):
    """Temporarily point the pipeline at another input workbook (used by the Streamlit upload)."""
    from pathlib import Path
    old = config.INPUT_WORKBOOK
    config.INPUT_WORKBOOK = Path(path)
    try:
        yield
    finally:
        config.INPUT_WORKBOOK = old
