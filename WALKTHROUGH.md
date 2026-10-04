# Walkthrough (interview notes)

## 1. What was the problem?
The fund is calling its 15th drawdown from 58 investors. Every investor needs a notice in one of 10 formats, then a PDF, then an email. Doing this by hand means copying numbers 58 times. The template sheets were also broken: their key cells pointed to `Working!#REF!`, and the templates were linked to each other.

## 2. What did I automate?
The whole chain: read the Working sheet, choose the right format per investor, fill it, check the numbers, make one PDF per investor, make a dummy email with that PDF attached, and write a QA report. One command: `python run.py`.

## 3. How does the system work?
Excel -> reader -> validation -> template mapper -> calculation engine -> notice generator -> single-notice Excel -> PDF export -> PDF QA -> dummy email -> QA report. Each step is its own small module in `src/`, and all settings sit in `config.py`.

## 4. How do I select the template?
Each investor row has a "Formats for DD Notice" value. A `TEMPLATE_MAPPING` dictionary turns that into a template sheet name. If the value is unknown, only that investor fails, with a clear message. One real catch: Working says "2nd close NR inves" but the sheet is called "2nd close NR investors", so I mapped both names.

## 5. How are calculations handled?
I recompute from the raw inputs in Python instead of trusting cached or broken Excel formulas: previous contribution = sum of the 1st to 14th drawdown payments; amount payable = called + pending - excess; percentage = (payable + previous + excess) / commitment. Then I compare with the workbook's own figures and flag any mismatch. For A1-001: 886,561.10 payable and 61.82%, matching the expected values. All 58 rows matched the workbook.

## 6. How are broken Excel references handled?
For the chosen template I (a) write investor values into the cells that used to say `#REF!`, (b) follow every remaining formula chain (for example NRIs -> R Inv -> Resident) to its final text, (c) delete all other sheets and clear the helper-notes columns. If any formula cannot be resolved, the investor errors out instead of exporting a bad PDF. The result is a notice with no formulas at all.

## 7. How are PDFs generated?
LibreOffice in headless mode converts each single-notice workbook (A4, portrait, fit to one page wide). Files are converted in batches, so 58 notices take about 17 seconds. Each PDF has only that investor's notice. Every template has its own page break, so each notice is 2 pages by design (page 1: figures and bank details; page 2: KYC note, remittance confirmation, signature).

## 8. How are emails generated?
Python's `email` module writes a `.eml` file per investor: subject "15th Drawdown Notice - <folio>", the standard body, and the matching PDF attached. In demo mode the recipient is a dummy `@example.com` address. Nothing is sent: `DRY_RUN = True`, and the Outlook function refuses to run unless it is explicitly enabled.

## 9. How do I validate the output?
I read each PDF back with PyMuPDF and check: page count for that template, no `#REF!`/`#NAME?`/`#VALUE!`/`#DIV/0!`, no sheet references, no placeholder text (`abc@gmail.com`, `For ____ LLP`, `CLIENT NAME AND ADDRESS`), the exact client name and address, the client e-mail, the signature block, a non-blank "To," section, and that class, folio, date, commitment, contribution, amount payable and percentage appear. For each `.eml` I check the recipient, that the subject carries the folio, and that the one attachment is byte-identical to that investor's PDF. Then batch checks: 58 / 58 / 58 files, one email per investor, all 10 templates used. I also tested the QA itself by putting each bug back on purpose; it caught every one. Results go to `final_qa_report.xlsx`.

## 10. What happens if one investor has an error?
It is logged, written to `error_report.xlsx` with folio, format, error, stage and timestamp, and the run continues. I tested this by removing one format from the mapping: 5 investors failed, 53 were still produced, and the error report listed the 5.

## 11. How can this be scaled?
More investors or drawdowns only need new Working rows. A new format needs one entry in the three config dictionaries (mapping, layout, file label). The 16-second run time scales linearly, and PDF conversion can run in parallel batches.

## 12. What would I improve in production?
* Get real names, addresses and emails from the client (they are blank in Working) and send via Outlook or an email API after a manual approval step.
* Fill the blank bank details and the "For ____ LLP" signatory from a master data sheet.
* Move the data into a database or Power Query and show the QA results in a Power BI dashboard.
* Make the due date a formula (business-day calendar) and confirm the 30 June vs 15-business-day question.
* Schedule runs with Power Automate and keep an audit trail of who approved what.

## Final QA result
58 investors processed, 58 successful, 0 failed. 58 Excel notices, 58 PDFs, 58 dummy emails, all 10 templates processed, every PDF has its expected 2 pages, 0 QA warnings, A1-001 test passes. Open remaining points: bank details blank in the templates, due date 30 June vs 15 business days, and investor name/address/email blank in Working (dummy labelled values used in DEMO_MODE).

---

## 2-minute explanation (spoken)

"First, I read the Working sheet from Excel and check that all the required columns are there.

Then, for each investor, I look at the notice format and pick the matching template using a mapping dictionary.

I calculate the drawdown figures in Python, previous contribution, amount payable and the percentage, and I compare them with the workbook's own numbers. For A1-001 I got 886,561.10 and 61.82%, exactly as expected.

Then I fill the investor data into the template. The original templates had broken `#REF!` links and linked to each other, so I resolve every formula to its final value and keep only that one sheet. That makes each notice self-contained.

I convert each notice to a PDF with LibreOffice. The templates are designed as two pages, so each PDF is two pages. The investor details are blank in the file, so in demo mode the notice shows clearly labelled dummy details instead of leaving the "To" section empty.

Then I run validation on every PDF: page count, broken references, placeholders, and the key values. I also create a dummy email for each investor with the right PDF attached, and I never send anything automatically.

If one investor fails, the run continues and the error goes into an error report. At the end I get a QA report and a summary: 58 processed, 58 successful.

The main idea was to turn a manual Excel and email process into a controlled, repeatable automation. Points I'd raise with the client: the investor names and bank details are blank in the file, and the due date doesn't match the 15-business-day wording."


---

## 5-minute demo script (Streamlit app: `streamlit run app.py`)

1. **Problem (30 s)** - Overview page, "5-minute story" strip: 58 investors, 10 notice formats, templates with broken `#REF!` links.
2. **Automation (90 s)** - click **Reset Demo**, then **Start Full Automation**. Watch the nine steps turn green while "Processing investor n / 58" counts up. About 17 seconds.
3. **Output (60 s)** - Investor Notices: pick **A1-001**; amounts (payable 886,561.10) match the PDF preview; download it. Email Automation: the dummy .eml, To `a1-001.demo@example.com`, attachment `A1-001_NRI_Drawdown_Notice.pdf`, DEMO MODE banner.
4. **QA (90 s)** - QA Dashboard: 58/58 pass, check table (`#REF!`... duplicate PDF), A1-001 reference test, per-investor filter. Mention *Known data gaps*: bank details and investor names are blank in the source file and were not invented.
5. **Value and next steps (30 s)** - repeatable, auditable, one failure never stops the batch; next: real master data, approval step before sending, Power BI on `final_qa_report.xlsx`.

### Likely questions about the UI
* *Did you rewrite the logic?* No. `app.py` calls `main.main()`; only an optional progress callback was added to `main.py`. The other eight modules are unchanged.
* *Can it send real e-mail?* No. The app has no sending path; `DRY_RUN = True` and `send_email_outlook()` is disabled by default.

### Q: Why not just use LibreOffice for the PDFs?
It is the most faithful converter, but it has to be installed, and Streamlit Community Cloud does not provide it. So the project now has a second engine, `src/pdf_renderer.py`, that draws the finished notice workbook with ReportLab. `PDF_ENGINE=auto` uses LibreOffice when it exists and falls back to the built-in renderer otherwise. The same QA runs on both: A1-001 reference test, page counts, placeholders, figures, signature, To-block.
