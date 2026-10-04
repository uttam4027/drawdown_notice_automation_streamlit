# Drawdown Notice Automation

Turns the **Working** sheet of `Drawdown working and notices.xlsx` into one PDF drawdown notice per investor, plus a dummy email (with the PDF attached) per investor and a QA report. Run with a single command: `python run.py`.

## 1. Project objective
Automate the 15th-drawdown notice process end to end: pick the right notice format for each investor, fill it from the Working data, produce a clean self-contained PDF, prepare the email, and prove with a QA report that every notice is correct.

## 2. Business problem
The fund calls capital from 58 investors. Each investor needs a notice in one of 10 formats (resident, NRI, LDC, institutional, 2nd/3rd/final close, employee/sponsor). Done by hand this means copying figures into 58 notices, exporting 58 PDFs and writing 58 emails, which is slow and error-prone. The template sheets also contain broken links (`Working!#REF!`) and link to each other.

## 3. Input workbook
`input/Drawdown working and notices.xlsx`

| Sheet | Used for |
|---|---|
| Instructions | The brief |
| Working | Source of truth: 58 investor rows (header row = row 2) |
| 10 x `Format - ...` | Notice templates (never renamed or modified on disk) |

## 4. Workflow
```
Working sheet -> read & validate columns -> investor record
   -> choose template (TEMPLATE_MAPPING) -> calculate + validate figures
   -> fill template -> resolve broken/cross-sheet formulas -> single-notice .xlsx
   -> PDF (LibreOffice) -> PDF QA -> dummy .eml with PDF attached -> QA report
```

## 5. Technology
Python, pandas (read Working), openpyxl (fill templates), LibreOffice headless (xlsx -> pdf), PyMuPDF (read PDFs for QA), standard-library `email` (.eml), `logging`, `pathlib`. No UI, no extra frameworks.

## 6. Data mapping
* **Working column -> logical field**: `FIELD_MAPPING` in `src/config.py`. Columns are found by **header text** (spaces/newlines/case ignored), not by position. The sheet has two "Class" columns; column C (the second, `Class.1`) holds the class (A1), column A repeats the folio.
* **Logical field -> template cell**: `TEMPLATE_LAYOUT` in `src/config.py` (e.g. Class -> A10, Folio -> A11, Date -> A13, Name -> A16, Address -> A17, table row 28 or 27, remit amount B36/B35/B37/B40). The original yellow cells lost this link (`#REF!`), so it is restored in one auditable place.
* **Working "Formats for DD Notice" -> template sheet**: `TEMPLATE_MAPPING` (also maps the Working label `Format - 2nd close NR inves` to the sheet `Format - 2nd close NR investors`).

## 7. Calculation logic
Recomputed in Python from Working inputs (not from cached or broken Excel formulas) and cross-checked against the workbook's own numbers:

* Previous contribution = sum of all "1st ... 14th Drawdown Amount Paid" columns
* Amount payable = 15th drawdown called + pending previous contribution - excess contribution
* % of commitment = (amount payable + previous contribution + excess contribution) / capital commitment

`validate_financial_calculations()` checks commitment, previous contribution, amount called, pending, excess, payable and % (recomputed vs workbook, no negatives, % not over 100%, portfolio + management fee + fund expenses = amount called). Any mismatch is reported in the QA report.

A1-001 test case: commitment 12,000,000.00, previous 6,508,925.00, called 699,633.00, pending 209,877.00, excess 22,948.90, payable 886,561.10, 61.82%, date 11 June 2026. It runs automatically before the batch (calculation level) and after it (PDF level).

## 8. PDF generation
For each investor the chosen template sheet is loaded, filled, and made self-contained: every formula is resolved to its final text/number by following the reference chain, other sheets are deleted, helper/note columns (G onwards) are cleared, and the yellow "to be filled" highlight is removed. Any formula that cannot be resolved raises an error for that investor, so a `#REF!` can never reach a PDF. A4 portrait, fit to 1 page wide, no gridlines. File name: `<folio>_<label>_Drawdown_Notice.pdf`.

**Pages:** each template has one deliberate manual page break and a repeated letterhead for page 2, so notices are **2 pages by design** (page 1: notice, amounts, bank details; page 2: KYC note, remittance confirmation, signature). The expected page count is read from the template and checked in QA.

## 9. Email generation
`.eml` per investor (`<folio>_<label>_Drawdown_Email.eml`) with the subject `15th Drawdown Notice - <folio>`, the standard body, and the matching PDF attached. `DEMO_MODE = True` sends every email to a clearly labelled dummy address (`a1-001.demo@example.com`), never to a real person. The client e-mail printed in the notice ("kindly inform us on our email id ...") is the separate config value `CLIENT_EMAIL` (`drawdown.demo@example.com`); it never touches an investor's email field. Nothing is sent: `DRY_RUN = True`; `send_email_outlook()` (pywin32) refuses unless DRY_RUN is turned off and `enabled=True` is passed.

## 10. Validation
Results are PASS, FAIL (investor status becomes QA_FAIL) or WARN (reported only).

**Per PDF:** page count vs the template's expected pages, empty pages, duplicate pages, `#REF!/#NAME?/#VALUE!/#DIV/0!`, visible cross-sheet references, banned text (`abc@gmail.com`, `For ____LLP`, `CLIENT NAME AND ADDRESS`, template build notes), exact client name and address (including "URMI Estate"), `CLIENT_EMAIL`, signature block (For <client> / On behalf of / Sd/- / Designated Partner), **"To," section not blank**, dummy investor name/address present in DEMO_MODE, required fields not blank (class, folio, date, remittance amount), plus class, folio, date, commitment, contributed, payable, % and the drawdown number, and the file name.

**Per email:** recipient exists and is a dummy `@example.com` address in DEMO_MODE, subject carries the investor's folio, exactly one attachment, attachment name belongs to the investor and is byte-identical to the generated PDF.

**Whole batch:** 58 Excel notices, 58 PDFs, 58 emails; exactly one `.eml` per investor; all 10 templates processed; pages per template. Results go to `final_qa_report.xlsx` and `summary.txt`.

## 11. Folder structure
```
drawdown_notice_automation/
  input/   Drawdown working and notices.xlsx
  output/  excel/ (single-notice workbooks)  pdf/  emails/  reports/
  src/     main, config, excel_reader, template_mapper, notice_generator,
           calculation_engine, pdf_generator, email_service, validator, reporting, utils
  logs/    drawdown_automation.log
  app.py   run.py   requirements.txt   README.md   WALKTHROUGH.md
  .streamlit/config.toml
```

## 12. How to run
```
pip install -r requirements.txt        # LibreOffice is optional (see section 17)
python run.py                          # add --skip-test to skip the A1-001 test
streamlit run app.py                   # optional UI over the same pipeline
```
If LibreOffice is not on the PATH, set the `SOFFICE_PATH` environment variable to `soffice.exe`. Settings (DEMO_MODE, DRY_RUN, mappings, client address, number formats) are in `src/config.py`.

## 13. Sample output
```
Total investors: 58
Processed: 58  Successful: 58  Failed: 0
PDF generated: 58  Dummy emails generated: 58
QA errors: 0
```
`output/pdf/A1-001_NRI_Drawdown_Notice.pdf`, `output/emails/A1-001_NRI_Drawdown_Email.eml`, `output/reports/final_qa_report.xlsx`, `error_report.xlsx`, `summary.txt`.

## 14. Known limitations / things to confirm with the client
* **2 pages, not 1:** follows the template's own page break. Squeezing onto one page made LibreOffice clip the signature block, so it was not done.
* **Investor name, address and email are blank in Working.** In `DEMO_MODE` the notice shows clearly labelled dummy details (`Demo Investor - <folio>`, `Demo Address, Mumbai, Maharashtra - 400013`) and emails go to `<folio>.demo@example.com`. These are not real investor data. With `DEMO_MODE = False` the "To," block stays blank and QA raises a WARN.
* **Bank details are blank in every template** (bank, account, IFSC, SWIFT, correspondent bank). They were not invented; the client must supply them.
* **Due date:** the notice says "within 15 Business Days" but the template's due date is 30 June 2026. 15 business days from 11 June 2026 is 2 July 2026 (ignoring holidays). The template date was kept; please confirm.
* **Placeholders replaced with demo/assignment values:** client e-mail (`drawdown.demo@example.com`), signatory ("For Nangia & Co LLP (On behalf of XYZ FUND)"), client address (assignment address incl. URMI Estate). All three are in `src/config.py`.
* **Small formatting changes:** amounts show 2 decimals (templates showed whole rupees, which hid the paise in 886,561.10); the % cell uses a % format; the LDC template's second line said "Class:" twice, it now shows "Folio No."; the yellow highlight is removed.
* PDFs are rendered by LibreOffice, so fonts can differ slightly from Excel's own PDF export.

## 15. Final QA result (full 58-investor run)
```
Total investors: 58        Processed: 58   Successful: 58   Failed: 0
Excel notices: 58          PDFs: 58        Dummy emails: 58
Templates processed: 10/10                 QA failures: 0   Errors: 0
Pages: every template 2 (expected 2)       PDFs with QA warnings: 0
Batch verification: PASS                   A1-001 test: PASS
```
Verified twice: by the built-in QA and by an independent scan of all 58 PDFs and 58 emails (no `abc@gmail.com`, no blank signatory, no blank "To," block, no error values, no placeholder text, URMI Estate present). The QA was also tested by deliberately reintroducing each bug; it caught all of them.

## 16. Streamlit UI (`app.py`)
`streamlit run app.py` opens **Drawdown Notice Automation & QA Platform**: Overview, Process, Investor Notices (figures + PDF preview + download), Email Automation (dummy .eml viewer), QA Dashboard, Analytics and Documentation.

* `app.py` is only the UI/control layer. It calls `main.main(progress=...)`, the same workflow as `python run.py`, and displays what the existing modules produced. It contains no calculations or business rules.
* Changes to `src/` for the UI (everything else is byte-identical to before): `main.py` got an optional `progress` callback (default `None`, so the command line behaves as before); `utils.py` got `use_workbook()` (point the pipeline at an uploaded workbook); new `reporting.py` (read-side helpers: summary, QA report, investor table, check-level QA that re-uses `validator.validate_pdf` / `validate_email`).
* Demo safety: DEMO_MODE and DRY_RUN stay on, e-mails are `.eml` files to `@example.com`, and there is no sending code path in the app. **Reset Demo** clears the generated files so a run can be shown from scratch (about 17 seconds).
* Data note: the QA Dashboard's check table groups `#REF!`, `#VALUE!`, `#NAME?`, `#DIV/0!` because the validator catches them with one combined scan.

## 17. PDF engines and Streamlit Community Cloud deployment
`config.PDF_ENGINE` (or the environment variable `PDF_ENGINE`) selects how a notice workbook becomes a PDF:

| Value | Behaviour |
|---|---|
| `auto` (default) | LibreOffice if it is installed, otherwise the built-in renderer |
| `libreoffice` | LibreOffice only; error if missing |
| `builtin` | Always the built-in renderer `src/pdf_renderer.py` (pure Python, ReportLab) |

The built-in renderer reads the same self-contained notice `.xlsx` and draws it: column widths, row heights, merged cells, borders, fonts, alignment, number formats and the template's page break. No LibreOffice, Excel or system package is needed. The same `validator.py` checks run on its output: with it, all 58 notices pass (2 pages each, 0 failures). Layout is very close to the LibreOffice version but not pixel-identical (fonts are ReportLab's built-in Times, line breaks can differ slightly). Use `PDF_ENGINE=libreoffice` if you need the original look.

**Deploy on Streamlit Community Cloud**
1. Push the project folder to a GitHub repo (keep `input/Drawdown working and notices.xlsx`; `output/` can be left out, the app recreates it).
2. share.streamlit.io -> New app -> pick the repo, branch and main file `app.py`.
3. `requirements.txt` is installed automatically; no `packages.txt` is needed.
4. Click **▶ Run Automation** in the app (about 20-30 s on the free tier).

Notes: the free tier's disk is temporary, so outputs vanish when the app restarts (just run again); everyone using the same app instance shares one `output/` folder, which is fine for a demo; keep the data sanitized because the repo/app may be public.
