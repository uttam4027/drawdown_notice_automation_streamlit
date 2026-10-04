"""STEP 8-10: populate the chosen template, remove broken / cross-sheet dependencies and save a
self-contained single-notice workbook.

Why this module is needed: the original templates link to each other ('Format - NRIs'!B3 ->
'Format - R Inv ...'!B3:D7 -> ...) and the yellow cells point at Working!#REF!. If a template were
exported as-is, the PDF would show #REF! or lose its text. So for the chosen template we
  1. write the investor values into the yellow cells (positions come from config.TEMPLATE_LAYOUT),
  2. resolve every remaining formula to its final text / number by following the reference chain,
  3. delete all other sheets, so nothing can point outside the notice.
Any formula we cannot resolve raises an error for that investor (it is never silently exported).
"""
import logging
import re

from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.properties import PageSetupProperties

import config
from utils import format_date

log = logging.getLogger("drawdown")

_SHEET = r"(?:'((?:[^']|'')+)'|([A-Za-z0-9_]+))"
_CELL = r"\$?([A-Z]{1,3})\$?(\d+)"
RE_SHEET_REF = re.compile(rf"^=\s*{_SHEET}!{_CELL}$")
RE_LOCAL_REF = re.compile(rf"^=\s*{_CELL}$")
RE_ARRAY_RANGE = re.compile(rf"^=\s*{_SHEET}!{_CELL}:\$?[A-Z]{{1,3}}\$?\d+$")
RE_ANCHOR = re.compile(rf"^=_xlfn\.ANCHORARRAY\({_SHEET}!{_CELL}\)$")
RE_CONCAT = re.compile(r"^=\s*([A-Z]{1,3}\d+)\s*&\s*([A-Z]{1,3}\d+)$")
RE_TEXT_DATE = re.compile(r'^=TEXT\(([A-Z]{1,3}\d+),\s*"dd mmmm yyyy"\)$')


class UnresolvedFormula(Exception):
    pass


def _sheet_name(g1, g2):
    return (g1.replace("''", "'") if g1 else g2)


def _is_formula(v):
    return isinstance(v, ArrayFormula) or (isinstance(v, str) and v.startswith("="))


def resolve_value(wb, sheet, coord, depth=0):
    """Follow a cell's formula chain across template sheets and return the final literal value."""
    if depth > 40:
        raise UnresolvedFormula(f"Reference chain too deep at {sheet}!{coord}")
    v = wb[sheet][coord].value

    if isinstance(v, ArrayFormula):
        text = v.text
        m = RE_ARRAY_RANGE.match(text) or RE_ANCHOR.match(text)
        if not m:
            raise UnresolvedFormula(f"Array formula not understood at {sheet}!{coord}: {text}")
        g1, g2, col, row = m.groups()
        return resolve_value(wb, _sheet_name(g1, g2), f"{col}{row}", depth + 1)

    if isinstance(v, str) and v.startswith("="):
        if "#REF!" in v or "#NAME?" in v:
            raise UnresolvedFormula(f"Broken reference at {sheet}!{coord}: {v}")
        m = RE_SHEET_REF.match(v)
        if m:
            g1, g2, col, row = m.groups()
            return resolve_value(wb, _sheet_name(g1, g2), f"{col}{row}", depth + 1)
        m = RE_LOCAL_REF.match(v)
        if m:
            return resolve_value(wb, sheet, f"{m.group(1)}{m.group(2)}", depth + 1)
        m = RE_CONCAT.match(v)
        if m:
            parts = [resolve_value(wb, sheet, c, depth + 1) for c in m.groups()]
            return "".join("" if p is None else str(p) for p in parts)
        m = RE_TEXT_DATE.match(v)
        if m:
            d = resolve_value(wb, sheet, m.group(1), depth + 1)
            return f"{d.day:02d} {d:%B %Y}"          # Excel "dd mmmm yyyy"
        raise UnresolvedFormula(f"Formula not supported at {sheet}!{coord}: {v}")
    return v


def _check_target(ws, coord):
    cell = ws[coord]
    if cell.__class__.__name__ == "MergedCell":
        raise ValueError(f"Layout cell {coord} is inside a merged range but is not its top-left cell")
    return cell


def _block_height(ws, coord):
    """Height in points of the merged range starting at coord (15pt default per row)."""
    for rng in ws.merged_cells.ranges:
        if coord in rng:
            return sum((ws.row_dimensions[r].height or 15) for r in range(rng.min_row, rng.max_row + 1))
    return ws.row_dimensions[ws[coord].row].height or 15


def _letterhead(ws, coord):
    big, small = (20, 11) if _block_height(ws, coord) >= 70 else (14, 9)
    font = config.LETTERHEAD_FONT
    return CellRichText(
        TextBlock(InlineFont(rFont=font, sz=big, b=True), config.CLIENT_NAME + "\n"),
        TextBlock(InlineFont(rFont=font, sz=small), "\n".join(config.CLIENT_ADDRESS_LINES)),
    )


def display_name(rec):
    """Name printed on the notice: Working value, else (DEMO_MODE only) a labelled dummy, else blank."""
    if rec["investor_name"]:
        return rec["investor_name"]
    return config.DEMO_INVESTOR_NAME.format(folio=rec["folio_no"]) if config.DEMO_MODE else None


def display_address(rec):
    if rec["address"]:
        return rec["address"]
    return config.DEMO_INVESTOR_ADDRESS if config.DEMO_MODE else None


def populate_notice(wb, sheet, rec, calc):
    """Write dynamic investor values into the yellow cells of the chosen template."""
    ws = wb[sheet]
    layout = config.TEMPLATE_LAYOUT[sheet]
    texts = {
        "class": config.CLASS_PREFIX + rec["class"],
        "folio_no": config.FOLIO_PREFIX + rec["folio_no"],
        "date": config.DATE_PREFIX + format_date(rec["drawdown_date"]),
        "investor_name": display_name(rec),
        "address": display_address(rec),
    }
    for field, value in texts.items():
        _check_target(ws, layout[field]).value = value
    for field in config.AMOUNT_FIELDS:
        if field in layout:
            cell = _check_target(ws, layout[field])
            cell.value = calc[field]
            cell.number_format = config.AMOUNT_NUMBER_FORMAT
    pct_cell = _check_target(ws, layout["percentage"])
    pct_cell.value = calc["percentage"]
    pct_cell.number_format = config.PERCENT_NUMBER_FORMAT


def apply_text_substitutions(ws):
    """Replace the template's placeholder client e-mail and blank signatory line (plain-text cells only).
    The investor's own e-mail field is never touched."""
    for row in ws.iter_rows(max_col=config.HELPER_COLUMNS_FROM - 1):
        for cell in row:
            v = cell.value
            if not isinstance(v, str):
                continue
            new = v.replace(config.TEMPLATE_EMAIL_PLACEHOLDER, config.CLIENT_EMAIL)
            new = re.sub(config.SIGNATORY_PLACEHOLDER_RE, config.SIGNATORY_TEXT, new)
            new = re.sub(config.SIGNATORY_ON_BEHALF_RE, config.SIGNATORY_ON_BEHALF_TEXT, new)
            if new != v:
                cell.value = new


def make_self_contained(wb, sheet):
    """Resolve every formula in the printable area to a literal, add the letterhead, clear helper columns."""
    ws = wb[sheet]
    # 1) resolve formulas (printable columns only; helper columns are cleared in step 3)
    for row in ws.iter_rows(max_col=config.HELPER_COLUMNS_FROM - 1):
        for cell in row:
            if cell.__class__.__name__ == "MergedCell":
                continue
            if _is_formula(cell.value):
                cell.value = resolve_value(wb, sheet, cell.coordinate)
    # 1b) client e-mail + signature placeholders
    apply_text_substitutions(ws)
    # 2) client name / address block(s)
    for row in ws.iter_rows(max_col=config.HELPER_COLUMNS_FROM - 1):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip() == config.CLIENT_PLACEHOLDER_TEXT:
                cell.value = _letterhead(ws, cell.coordinate)
    # 3) clear helper / build-note columns (G onwards)
    for row in ws.iter_rows(min_col=config.HELPER_COLUMNS_FROM):
        for cell in row:
            if cell.__class__.__name__ != "MergedCell":
                cell.value = None
    # 3b) yellow "to be filled" marker is a build aid, not part of the notice
    if config.REMOVE_YELLOW_HIGHLIGHT:
        for row in ws.iter_rows():
            for cell in row:
                f = cell.fill
                if f is not None and f.fill_type == "solid" and getattr(f.fgColor, "rgb", None) == "FFFFFF00":
                    cell.fill = PatternFill(fill_type=None)
    # 4) nothing formula-like may remain anywhere
    for row in ws.iter_rows():
        for cell in row:
            if _is_formula(cell.value):
                raise UnresolvedFormula(f"Formula left in {sheet}!{cell.coordinate}")


def _print_area(ws):
    last_row, last_col = 1, 5
    for row in ws.iter_rows(max_col=config.HELPER_COLUMNS_FROM - 1):
        for cell in row:
            if cell.value not in (None, ""):
                last_row = max(last_row, cell.row)
                last_col = max(last_col, cell.column)
    for rng in ws.merged_cells.ranges:
        if rng.min_col < config.HELPER_COLUMNS_FROM:
            last_row = max(last_row, rng.max_row)
            last_col = max(last_col, min(rng.max_col, config.HELPER_COLUMNS_FROM - 1))
    return f"A1:{get_column_letter(last_col)}{last_row}"


def set_page_setup(ws):
    ws.print_area = _print_area(ws)
    ws.page_setup.paperSize = 9                    # A4
    ws.page_setup.orientation = "portrait"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.scale = None
    ws.page_setup.fitToHeight = 0                # fit to 1 page WIDE; the template's own page break is kept
    ws.print_options.gridLines = False
    ws.print_options.horizontalCentered = True
    ws.sheet_view.showGridLines = False
    ws.page_margins.left = ws.page_margins.right = 0.6
    ws.page_margins.top = ws.page_margins.bottom = 0.5


def expected_pages(ws):
    """The templates carry a manual page break per page: expected pages = breaks + 1."""
    return len(ws.row_breaks.brk) + 1


def keep_only(wb, sheet):
    for name in list(wb.sheetnames):
        if name != sheet:
            del wb[name]
    wb.active = 0
    # drop defined names that still point at removed sheets
    for key in list(wb.defined_names.keys()):
        if wb.defined_names[key].attr_text and any(
                n in wb.defined_names[key].attr_text for n in ("Working!",)):
            del wb.defined_names[key]


def generate_notice(rec, calc, sheet, out_path, source_path=None):
    """Build one self-contained single-notice workbook. Returns {"path", "expected_pages"}."""
    wb = load_workbook(source_path or config.INPUT_WORKBOOK)       # fresh copy per investor
    populate_notice(wb, sheet, rec, calc)
    make_self_contained(wb, sheet)
    set_page_setup(wb[sheet])
    pages = expected_pages(wb[sheet])
    keep_only(wb, sheet)
    wb.save(out_path)
    log.info("Notice workbook created: %s (expected pages: %d)", out_path.name, pages)
    return {"path": out_path, "expected_pages": pages}


def data_warnings(rec):
    """Missing investor data is never invented - it is reported instead."""
    w = []
    suffix = " (labelled demo value used)" if config.DEMO_MODE else ""
    if not rec["investor_name"]:
        w.append("Investor name blank in Working" + suffix)
    if not rec["address"]:
        w.append("Investor address blank in Working" + suffix)
    if not rec["email"]:
        w.append("Investor email blank in Working" + (" (dummy recipient used)" if config.DEMO_MODE else ""))
    return w


# name used in the assignment brief
populate_template = populate_notice
