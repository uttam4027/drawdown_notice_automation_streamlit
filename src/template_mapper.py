"""STEP 5-6: decide which template sheet an investor's notice is built from."""
import logging

import config
from utils import normalize

log = logging.getLogger("drawdown")

_MAPPING_N = {normalize(k): v for k, v in config.TEMPLATE_MAPPING.items()}


def get_template(format_value, workbook_sheetnames):
    """Working['Formats for DD Notice'] -> template sheet name. Raises with a clear message if unknown."""
    if not format_value:
        raise ValueError("'Formats for DD Notice' is blank in Working")
    sheet = _MAPPING_N.get(normalize(format_value))
    if sheet is None:
        raise ValueError(f"No TEMPLATE_MAPPING entry for format '{format_value}'")
    if sheet not in workbook_sheetnames:
        raise ValueError(f"Template sheet '{sheet}' not found in workbook")
    if sheet not in config.TEMPLATE_LAYOUT or sheet not in config.FILE_LABEL:
        raise ValueError(f"Template '{sheet}' has no layout / file label in config")
    log.info("Template selected: '%s' -> '%s'", format_value, sheet)
    return sheet
