"""STEP 12: automatic QA on every generated PDF / email, batch-level checks, and the A1-001 test case.

Check results are PASS, FAIL (makes the investor's status QA_FAIL) or WARN (reported, does not fail).
"""
import email
import hashlib
import logging
import re
from email import policy
from pathlib import Path

import pymupdf

import config
from calculation_engine import calculate_values
from notice_generator import display_address, display_name
from utils import email_basename, fmt_amount, format_date, notice_basename

log = logging.getLogger("drawdown")

ERROR_VALUE_RE = re.compile(r"#REF!|#NAME\?|#VALUE!|#DIV/0!")
CROSS_SHEET_RE = re.compile(r"Working!|='|\bFormat - ")
TO_BLOCK_RE = re.compile(r"\bTo,\s*(.*?)\s*Re:\s*Notice")


def _norm(text):
    return re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()


def validate_pdf(pdf_path, rec, calc, expected_pages, expected_stem):
    """Returns {"checks": {name: PASS/FAIL/WARN}, "pages": n, "problems": [FAIL msgs], "warnings": [WARN msgs]}"""
    problems, warnings, checks = [], [], {}
    pdf_path = Path(pdf_path)
    doc = pymupdf.open(pdf_path)
    page_texts = [_norm(p.get_text()) for p in doc]
    full = " ".join(page_texts)
    pages = len(doc)
    doc.close()

    def record(name, ok, msg="", warn_only=False):
        if ok:
            checks[name] = "PASS"
        elif warn_only:
            checks[name] = "WARN"
            warnings.append(f"{name}: {msg}")
        else:
            checks[name] = "FAIL"
            problems.append(f"{name}: {msg}")

    # --- structure ---------------------------------------------------------------------------------
    record("pages", pages == expected_pages, f"{pages} page(s), expected {expected_pages}")
    record("empty_pages", all(len(t) >= 20 for t in page_texts), "an empty/near-empty page was found")
    hashes = [hashlib.md5(t.encode()).hexdigest() for t in page_texts]
    record("duplicate_pages", len(set(hashes)) == len(hashes), "two pages have identical content")

    # --- broken references / placeholders --------------------------------------------------------------
    err_vals = sorted(set(ERROR_VALUE_RE.findall(full)))
    record("error_values", not err_vals, f"Excel error text in PDF: {err_vals}")
    xrefs = sorted(set(CROSS_SHEET_RE.findall(full)))
    record("cross_sheet_references", not xrefs, f"visible sheet references: {xrefs}")
    checks["broken_references"] = "PASS" if not (err_vals or xrefs) else "FAIL"
    banned = [label for label, rx in config.BANNED_TEXT_RE.items() if re.search(rx, full, re.I)]
    record("banned_text", not banned, f"template placeholder text still present: {banned}")

    # --- letterhead / client block / client e-mail / signature -------------------------------------------
    record("client_block", config.CLIENT_NAME in full and all(_norm(l) in full for l in config.CLIENT_ADDRESS_LINES),
           "client name / exact address (incl. URMI Estate) missing")
    record("client_email", config.CLIENT_EMAIL in full, f"{config.CLIENT_EMAIL} not shown in the notice")
    sig_ok = (config.SIGNATORY_TEXT in full and config.SIGNATORY_ON_BEHALF_TEXT in full
              and all(s in full for s in config.SIGNATURE_LINES_EXPECTED))
    record("signature", sig_ok, "signature block incomplete (For <client> / On behalf of / Sd/- / Designated Partner)")

    # --- "To," section -------------------------------------------------------------------------------
    m = TO_BLOCK_RE.search(full)
    to_text = m.group(1).strip() if m else ""
    exp_name, exp_addr = display_name(rec), display_address(rec)
    if config.DEMO_MODE:
        record("to_section", bool(to_text), '"To," section is blank')
        record("investor_details", bool(exp_name) and bool(exp_addr) and _norm(exp_name) in to_text
               and _norm(exp_addr) in to_text, f"expected '{exp_name}' and the address in the To section")
    else:
        record("to_section", bool(to_text), '"To," section is blank (investor name/address blank in Working)',
               warn_only=True)
        if exp_name and exp_addr:
            record("investor_details", _norm(exp_name) in to_text and _norm(exp_addr) in to_text,
                   "investor name/address from Working not found in the To section")

    # --- required values -------------------------------------------------------------------------------
    expected_strings = {
        "class": config.CLASS_PREFIX + rec["class"],
        "folio": config.FOLIO_PREFIX + rec["folio_no"],
        "drawdown_date": config.DATE_PREFIX + format_date(rec["drawdown_date"]),
        "amount_payable": fmt_amount(calc["amount_payable"]),
        "capital_commitment": fmt_amount(calc["capital_commitment"]),
        "capital_contributed": fmt_amount(calc["capital_contributed"]),
        "percentage": f"{calc['percentage']:.2%}",
        "drawdown_number": config.EXPECTED_DRAWDOWN_TEXT,
    }
    missing = [k for k, s in expected_strings.items() if s not in full]
    record("financial_values", not missing, f"not found in PDF text: {missing}")

    blank_required = []
    for label, rx in {"class": r"Class:\s*\S", "folio": r"Folio No\.:\s*\S", "date": r"Date:\s*\d"}.items():
        if not re.search(rx, full):
            blank_required.append(label)
    # remittance box: the label can wrap onto two lines in some templates, so instead of matching
    # "label + number" require the payable figure twice (summary table + "Amount to be remitted" box)
    if "Amount to be remitted" not in full or full.count(fmt_amount(calc["amount_payable"])) < 2:
        blank_required.append("amount to be remitted")
    record("required_fields", not blank_required, f"blank required field(s): {blank_required}")

    record("filename", pdf_path.stem == expected_stem and pdf_path.suffix == ".pdf",
           f"file name {pdf_path.name} != {expected_stem}.pdf")
    return {"checks": checks, "pages": pages, "problems": problems, "warnings": warnings}


def validate_email(eml_path, pdf_path, rec, expected_to, expected_subject):
    """The .eml must have the right recipient/subject and exactly one attachment: the byte-identical PDF
    that belongs to this investor."""
    problems = []
    with open(eml_path, "rb") as fh:
        msg = email.message_from_binary_file(fh, policy=policy.default)
    if msg["To"] != expected_to:
        problems.append(f"email: To is {msg['To']}, expected {expected_to}")
    if not msg["To"]:
        problems.append("email: no recipient")
    elif config.DEMO_MODE and not str(msg["To"]).endswith("@" + config.DEMO_EMAIL_DOMAIN):
        problems.append(f"email: recipient {msg['To']} is not a dummy @{config.DEMO_EMAIL_DOMAIN} address")
    if msg["Subject"] != expected_subject or rec["folio_no"] not in str(msg["Subject"]):
        problems.append(f"email: subject '{msg['Subject']}' does not carry folio {rec['folio_no']}")
    atts = list(msg.iter_attachments())
    if len(atts) != 1:
        problems.append(f"email: {len(atts)} attachments (expected 1)")
    else:
        name = atts[0].get_filename()
        if name != Path(pdf_path).name or not name.startswith(rec["folio_no"] + "_"):
            problems.append(f"email: attachment name {name} does not match investor {rec['folio_no']}")
        if atts[0].get_content() != Path(pdf_path).read_bytes():
            problems.append("email: attached PDF differs from the generated PDF")
    return problems


def verify_batch(records, qa_rows):
    """Batch-level verification over the whole run. Returns (list of failure strings, dict of facts)."""
    fails = []
    n = len(records)
    excel_files = sorted(config.EXCEL_DIR.glob("*.xlsx"))
    pdf_files = sorted(config.PDF_DIR.glob("*.pdf"))
    eml_files = sorted(config.EMAIL_DIR.glob("*.eml"))
    ok_rows = [r for r in qa_rows if r["Status"] != "ERROR"]

    for label, files in (("Excel notices", excel_files), ("PDFs", pdf_files), ("Dummy emails", eml_files)):
        if len(files) != n:
            fails.append(f"{label}: {len(files)} files, expected {n}")

    # exactly one .eml per investor, named after that investor
    for r in qa_rows:
        if r["Status"] == "ERROR":
            continue
        mine = [f for f in eml_files if f.name.startswith(r["Folio No."] + "_")]
        if len(mine) != 1:
            fails.append(f"{r['Folio No.']}: {len(mine)} .eml files (expected 1)")

    templates_used = sorted({r["Format"] for r in ok_rows})
    mapped = sorted(set(config.TEMPLATE_MAPPING.values()))
    templates_seen = sorted({config.TEMPLATE_MAPPING.get(t, t) for t in templates_used})
    if len(templates_seen) != len(mapped):
        fails.append(f"templates processed: {len(templates_seen)} of {len(mapped)} "
                     f"(missing: {sorted(set(mapped) - set(templates_seen))})")

    facts = {
        "Excel notices": len(excel_files), "PDFs": len(pdf_files), "Dummy emails": len(eml_files),
        "Templates processed": f"{len(templates_seen)}/{len(mapped)}",
        "Pages per template": _pages_per_template(ok_rows),
    }
    return fails, facts


def _pages_per_template(rows):
    out = {}
    for r in rows:
        out.setdefault(r["Format"], set()).add(f"{r['Pages']} (expected {r['Expected Pages']})")
    return {k: ", ".join(sorted(v)) for k, v in sorted(out.items())}


# ------------------------------------------------------------------ A1-001 test case
A1_001_EXPECTED = {
    "folio_no": "A1-001",
    "notice_format": "Format - NRIs",
    "capital_commitment": 12_000_000.00,
    "capital_contributed": 6_508_925.00,
    "drawdown_called": 699_633.00,
    "pending_previous": 209_877.00,
    "excess_contribution": 22_948.90,
    "amount_payable": 886_561.10,
    "percentage_rounded": 0.6182,
    "date_text": "11 June 2026",
}


def run_a1_001_calc_test(records):
    """Calculation-level test, run BEFORE processing everybody. Returns a list of failures."""
    exp = A1_001_EXPECTED
    rec = next((r for r in records if r["folio_no"] == exp["folio_no"]), None)
    if rec is None:
        return ["A1-001 not found in Working"]
    calc = calculate_values(rec)
    fails = []
    if rec["notice_format"] != exp["notice_format"]:
        fails.append(f"format {rec['notice_format']} != {exp['notice_format']}")
    for key in ("capital_commitment", "capital_contributed", "drawdown_called", "pending_previous",
                "excess_contribution", "amount_payable"):
        if abs(calc[key] - exp[key]) > config.MONEY_TOLERANCE:
            fails.append(f"{key}: {calc[key]:,.2f} != {exp[key]:,.2f}")
    if round(calc["percentage"], 4) != exp["percentage_rounded"]:
        fails.append(f"percentage {calc['percentage']:.4%} != {exp['percentage_rounded']:.2%}")
    if format_date(rec["drawdown_date"]) != exp["date_text"]:
        fails.append(f"date {format_date(rec['drawdown_date'])} != {exp['date_text']}")
    return fails


def check_a1_001_outputs():
    """Output-level test, run AFTER generation: files exist and the PDF shows the expected content."""
    exp = A1_001_EXPECTED
    stem = notice_basename(exp["folio_no"], "Format - NRIs")
    pdf = config.PDF_DIR / f"{stem}.pdf"
    eml = config.EMAIL_DIR / f"{email_basename(exp['folio_no'], 'Format - NRIs')}.eml"
    fails = []
    if not pdf.exists():
        return [f"{pdf.name} missing"]
    if not eml.exists():
        fails.append(f"{eml.name} missing")
    doc = pymupdf.open(pdf)
    text = _norm(" ".join(p.get_text() for p in doc))
    doc.close()
    wanted = {"commitment": "12,000,000.00", "previous": "6,508,925.00", "called": "699,633.00",
              "payable": "886,561.10", "pct": "61.82%", "date": "11 June 2026",
              "client": config.CLIENT_NAME, "client address": _norm(config.CLIENT_ADDRESS_LINES[0]),
              "client email": config.CLIENT_EMAIL, "signature": config.SIGNATORY_TEXT}
    if config.DEMO_MODE:
        wanted["demo investor"] = config.DEMO_INVESTOR_NAME.format(folio=exp["folio_no"])
        wanted["demo address"] = config.DEMO_INVESTOR_ADDRESS
    for label, s in wanted.items():
        if s not in text:
            fails.append(f"PDF is missing {label} ({s})")
    if ERROR_VALUE_RE.search(text) or CROSS_SHEET_RE.search(text):
        fails.append("PDF contains a broken / cross-sheet reference")
    for label, rx in config.BANNED_TEXT_RE.items():
        if re.search(rx, text, re.I):
            fails.append(f"banned text present: {label}")
    return fails


# name used in the assignment brief
validate_notice = validate_pdf
