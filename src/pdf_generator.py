"""STEP 11: convert single-notice workbooks to PDF with LibreOffice (headless)."""
import logging
import subprocess
from pathlib import Path

import config
import pdf_renderer

log = logging.getLogger("drawdown")


def _convert(cmd_base, files, pdf_dir, timeout):
    cmd = cmd_base + ["--headless", "--convert-to", "pdf", "--outdir", str(pdf_dir)] + [str(f) for f in files]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    if result.returncode != 0:
        log.warning("LibreOffice returned %s: %s", result.returncode, result.stderr.strip()[:300])


def export_pdf(xlsx_files, pdf_dir=None, batch_size=20):
    """Convert each single-notice .xlsx to a .pdf with the same file name.

    One LibreOffice start-up handles a whole batch (much faster than one start per file); anything that
    did not come out is retried individually. Returns {xlsx_stem: pdf_path or None}.
    """
    pdf_dir = Path(pdf_dir or config.PDF_DIR)
    pdf_dir.mkdir(parents=True, exist_ok=True)
    cmd_base = config.find_soffice() if config.PDF_ENGINE in ("auto", "libreoffice") else None
    if not cmd_base:
        if config.PDF_ENGINE == "libreoffice":
            raise RuntimeError("LibreOffice not found. Install it or set the SOFFICE_PATH environment variable.")
        # no LibreOffice (e.g. Streamlit Community Cloud): draw the PDFs with the built-in pure-Python renderer
        log.info("LibreOffice not used - exporting PDFs with the built-in renderer")
        return pdf_renderer.export_pdf(xlsx_files, pdf_dir)

    files = [Path(f) for f in xlsx_files]
    for i in range(0, len(files), batch_size):
        batch = files[i:i + batch_size]
        try:
            _convert(cmd_base, batch, pdf_dir, timeout=600)
        except subprocess.TimeoutExpired:
            log.error("LibreOffice batch timed out (%d files)", len(batch))

    results = {}
    for f in files:
        pdf = pdf_dir / (f.stem + ".pdf")
        if not pdf.exists():                                   # retry one by one
            try:
                _convert(cmd_base, [f], pdf_dir, timeout=180)
            except subprocess.TimeoutExpired:
                log.error("LibreOffice timed out on %s", f.name)
        results[f.stem] = pdf if pdf.exists() else None
        if results[f.stem]:
            log.info("PDF created: %s", pdf.name)
        else:
            log.error("PDF NOT created for %s", f.name)
    return results


def active_engine():
    """Name of the PDF engine export_pdf() will use right now (shown in the UI)."""
    if config.PDF_ENGINE in ("auto", "libreoffice") and config.find_soffice():
        return "LibreOffice"
    return "Built-in (ReportLab, no LibreOffice needed)"
