"""All configuration lives here - no business values are hardcoded in the logic modules."""
import os
import re
import shutil
from pathlib import Path

# ----------------------------------------------------------------- paths
PROJECT_ROOT = Path(__file__).resolve().parent.parent
INPUT_WORKBOOK = PROJECT_ROOT / "input" / "Drawdown working and notices.xlsx"
OUT_DIR = PROJECT_ROOT / "output"
EXCEL_DIR = OUT_DIR / "excel"
PDF_DIR = OUT_DIR / "pdf"
EMAIL_DIR = OUT_DIR / "emails"
REPORT_DIR = OUT_DIR / "reports"
LOG_DIR = PROJECT_ROOT / "logs"
LOG_FILE = LOG_DIR / "drawdown_automation.log"

# ----------------------------------------------------------------- run modes
DEMO_MODE = True        # blank investor name/address get labelled dummy text; emails go to @example.com
DRY_RUN = True          # True = never send a real email (email_service.send_email_outlook refuses)

# Dummy investor details (DEMO_MODE only, used only where Working is blank). {folio} = the investor's own folio.
DEMO_INVESTOR_NAME = "Demo Investor - {folio}"
DEMO_INVESTOR_ADDRESS = "Demo Address, Mumbai, Maharashtra - 400013"

# ----------------------------------------------------------------- Working sheet
WORKING_SHEET = "Working"
WORKING_HEADER_ROW = 2          # row holding the column headers (1-based)
PAID_COLUMN_REGEX = re.compile(r"drawdown amount paid", re.I)   # every "<n> Drawdown Amount Paid" column

# Logical field -> header text in the Working sheet (matched after trimming/collapsing spaces, case-insensitive).
# NOTE: the sheet has two "Class" columns. Column A repeats the folio (A1-001); column C holds the class (A1).
# pandas names the second one "Class.1".
FIELD_MAPPING = {
    "class": "Class.1",
    "folio_no": "Folio No.",
    "notice_format": "Formats for DD Notice",
    "investor_name": "Name of the contributor",
    "address": "Address",
    "email": "Email ID's",
    "capital_commitment": "Capital commitment",
    "capital_contributed_wb": "Capital contributed till date",     # workbook value, used only for cross-checking
    "drawdown_date": "Date for the 15th Drawdown",
    "drawdown_called": "15th Drawdown Amount Called",
    "portfolio_investment": "Capital contribution for Portfolio Investment (INR)",
    "adjustment": "Adjustment pursuant to previous drawdown (INR)",
    "management_fee": "Capital contribution for Management fees (INR)",
    "fund_expenses": "Capital contribution for Establishment and Fund Expenses (INR)",
    "pending_previous": "Pending from previous contribution",
    "excess_contribution": "Excess Contribution",
    "amount_payable_wb": "15th Drawdown Amount Payable",          # workbook value, used only for cross-checking
    "percentage_wb": "% on the Capital Commitment",               # workbook value, used only for cross-checking
}
REQUIRED_FIELDS = [
    "class", "folio_no", "notice_format", "capital_commitment", "drawdown_date",
    "drawdown_called", "pending_previous", "excess_contribution",
]

# ----------------------------------------------------------------- template selection
# Value in Working["Formats for DD Notice"]  ->  template sheet name in the workbook.
# (Working says "2nd close NR inves" but the sheet is called "2nd close NR investors" - mapped here.)
TEMPLATE_MAPPING = {
    "Format - Resident investors": "Format - Resident investors",
    "Format - R Inv (Empl & Sponsor)": "Format - R Inv (Empl & Sponsor)",
    "Format - NRIs": "Format - NRIs",
    "Format - LDC's": "Format - LDC's",
    "Format - 2nd close R investors": "Format - 2nd close R investors",
    "Format - 2nd close NR investors": "Format - 2nd close NR investors",
    "Format - 2nd close NR inves": "Format - 2nd close NR investors",
    "Format - 3rd close NR inves": "Format - 3rd close NR inves",
    "Format - 3rd close R invest": "Format - 3rd close R invest",
    "Format - Final close Employee": "Format - Final close Employee",
    "Format - Institutional Inv": "Format - Institutional Inv",
}

# Short label used in output file names:  <folio>_<label>_Drawdown_Notice.pdf
FILE_LABEL = {
    "Format - Resident investors": "Resident",
    "Format - R Inv (Empl & Sponsor)": "Employee_Sponsor",
    "Format - NRIs": "NRI",
    "Format - LDC's": "LDC",
    "Format - 2nd close R investors": "2ndClose_Resident",
    "Format - 2nd close NR investors": "2ndClose_NRI",
    "Format - 3rd close NR inves": "3rdClose_NRI",
    "Format - 3rd close R invest": "3rdClose_Resident",
    "Format - Final close Employee": "FinalClose_Employee",
    "Format - Institutional Inv": "Institutional",
}


# ----------------------------------------------------------------- template layouts
# Which template cell receives which value. (The #REF! cells in the original templates lost
# this link, so it is restored here - one place, easy to audit.)
def _std(table_row=28, remit="B36"):
    r = table_row
    return {
        "class": "A10", "folio_no": "A11", "date": "A13", "investor_name": "A16", "address": "A17",
        "capital_commitment": f"A{r}", "capital_contributed": f"B{r}", "drawdown_called": f"C{r}",
        "amount_payable": f"D{r}", "percentage": f"E{r}", "remit_amount": remit,
    }


TEMPLATE_LAYOUT = {
    "Format - Resident investors": _std(),
    "Format - R Inv (Empl & Sponsor)": _std(),
    "Format - NRIs": _std(),
    "Format - 2nd close R investors": _std(),
    "Format - 2nd close NR investors": _std(),
    "Format - 3rd close NR inves": _std(),
    "Format - Institutional Inv": _std(remit="B37"),
    "Format - 3rd close R invest": _std(table_row=27, remit="B35"),
    "Format - Final close Employee": _std(table_row=27, remit="B35"),
    # LDC: extra adjustment column, 6-column table and a portfolio / management fee / fund expense split
    "Format - LDC's": {
        "class": "A10", "folio_no": "A11", "date": "A13", "investor_name": "A16", "address": "A17",
        "capital_commitment": "A28", "capital_contributed": "B28", "drawdown_called": "C28",
        "adjustment": "D28", "amount_payable": "E28", "percentage": "F28",
        "portfolio_investment": "A33", "management_fee": "C33", "fund_expenses": "E33",
        "remit_amount": "B40",
    },
}
AMOUNT_FIELDS = ["capital_commitment", "capital_contributed", "drawdown_called", "adjustment",
                 "amount_payable", "portfolio_investment", "management_fee", "fund_expenses", "remit_amount"]

# Labels in front of the dynamic value in the template cells
CLASS_PREFIX = "Class: "
FOLIO_PREFIX = "Folio No.: "
DATE_PREFIX = "Date: "
DATE_FORMAT = "{d.day} {d:%B %Y}"            # 11 June 2026

# Number formats (templates showed whole rupees and a wrongly formatted % cell; paise matter for a remittance)
AMOUNT_NUMBER_FORMAT = '_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_ '
PERCENT_NUMBER_FORMAT = "0.00%"

# Print: A4 portrait, fit to 1 page wide. Every original template has ONE deliberate manual page break
# (page 1 = notice + bank details, page 2 = repeated letterhead + KYC note + remittance confirmation + signature),
# so each notice is 2 pages. (Squeezing it onto one page made LibreOffice clip the signature block, so the
# template's own layout is respected.) The expected page count is read from the template's page breaks.
REMOVE_YELLOW_HIGHLIGHT = True      # yellow = "to be filled" marker in the templates; not wanted in the final notice

# Helper columns in the templates hold build notes ("reflects Class") and date helpers - never printed
HELPER_COLUMNS_FROM = 7                      # column G onwards is cleared in the generated notice

# ----------------------------------------------------------------- letterhead / client block
CLIENT_PLACEHOLDER_TEXT = "CLIENT NAME AND ADDRESS"
CLIENT_NAME = "Nangia & Co LLP"
CLIENT_ADDRESS = ("4th Floor, Iconic Tower, URMI Estate, Ganpat Rao Kadam Marg,\n"
                  "Lower Parel - 400013, India")
CLIENT_ADDRESS_LINES = CLIENT_ADDRESS.split("\n")

# 2) client e-mail shown in "Once the payment is done, kindly inform us on our email id ..."
#    (this is NOT an investor e-mail - the investor's own email field is never touched)
CLIENT_EMAIL = "drawdown.demo@example.com"
TEMPLATE_EMAIL_PLACEHOLDER = "abc@gmail.com"

# 3) signature block: template text  "For ___________________LLP\n(On behalf of  XYZ FUND)"
SIGNATORY_PLACEHOLDER_RE = r"For\s*_+\s*LLP"
SIGNATORY_TEXT = "For " + CLIENT_NAME
SIGNATORY_ON_BEHALF_RE = r"\(On behalf of\s+XYZ FUND\)"
SIGNATORY_ON_BEHALF_TEXT = "(On behalf of XYZ FUND)"
SIGNATURE_LINES_EXPECTED = ["Sd/-", "Designated Partner"]       # template text, kept as is

# Things that must never appear in a finished notice (checked in PDF QA)
BANNED_TEXT_RE = {
    "template client e-mail": r"abc@gmail\.com",
    "blank signatory": r"For\s*_{3,}\s*LLP",
    "client placeholder": r"CLIENT NAME AND ADDRESS|Input Client Name",
    "build notes": r"reflects (?:Class|Folio)|of working tab|to be changed",
}

# Warnings that cannot be fixed in code because the data is not in the workbook (listed in summary.txt)
KNOWN_WARNINGS = [
    "Bank details (bank, account no., IFSC, SWIFT, correspondent bank) are blank in every template - not invented.",
    "Investor name/address/email are blank in Working - dummy labelled values used in DEMO_MODE.",
    "Template due date is 30 June 2026 but the wording says 15 Business Days from 11 June 2026 (= 2 July 2026).",
]
LETTERHEAD_FONT = "Times New Roman"

# ----------------------------------------------------------------- expected business values
EXPECTED_DRAWDOWN_TEXT = "15th"              # sanity check: the notice text must mention this drawdown

# ----------------------------------------------------------------- email
EMAIL_SUBJECT = "15th Drawdown Notice - {folio}"
EMAIL_BODY = (
    "Dear Investor,\n\n"
    "Please find attached the 15th Drawdown Notice for your reference.\n\n"
    "Kindly review the notice and arrange the payment as mentioned in the notice.\n\n"
    "Regards,\nNangia & Co LLP\n"
)
DEMO_EMAIL_DOMAIN = "example.com"
DEMO_SENDER = "demo.sender@example.com"

# ----------------------------------------------------------------- numeric tolerance
MONEY_TOLERANCE = 0.01
PCT_TOLERANCE = 1e-6


# ----------------------------------------------------------------- PDF conversion
# "auto"        = LibreOffice if it is installed, otherwise the built-in pure-Python renderer (pdf_renderer.py)
# "libreoffice" = LibreOffice only (error if missing)      "builtin" = always the built-in renderer
PDF_ENGINE = os.environ.get("PDF_ENGINE", "auto").lower()


def find_soffice():
    """Return the command (list) used to run LibreOffice headless, or None."""
    env = os.environ.get("SOFFICE_PATH")
    if env:
        return [env]
    wrapper = Path("/mnt/skills/public/docx/scripts/office/soffice.py")   # sandbox helper (bare soffice hangs there)
    if wrapper.exists():
        return ["python3", str(wrapper)]
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return [found]
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if Path(p).exists():
            return [p]
    return None
